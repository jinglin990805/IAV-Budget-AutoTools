# -*- coding: utf-8 -*-
"""
GR WP 合并器 — 按季度 + PO + 人员合并审核后的工时表 PDF

数据来源：data/iav_facts.sqlite3 中 approval_status='APPROVED' 且 is_current=1
的 timesheet_facts（审核通过后才算数，candidates 不参与）。

输出（参考 Q2\\WP 结构，一个 PO 一个文件夹，里面每人一个合并 PDF）：
    CNTS-25165\\03_GR Documents\\2026\\VCTC\\<Q>\\WP\\<PO 文件夹>\\<每人合并 PDF>

命名：
    文件夹  IAV_<人员1>&<人员2>..._<MM-MM>_PO<po>
    文件    IAV_<人员无空格全名>_<MM-MM>_PO<po>.pdf

电子签处理：
    若某「PO+人员」分组中的任一源 PDF 含数字签名（adbe.pkcs7.detached / SigFlags /
    DocMDP）或已加密（Encrypt）或文件缺失，则该分组不自动合并，进入 manual 清单，
    由用户手动处理。

用法：
    python timesheet_wp_merger.py            # Q3 实际合并
    python timesheet_wp_merger.py --dry-run  # 仅预览，不写文件
    python timesheet_wp_merger.py --quarter Q3 --json
"""

import os
import sys
import io
import json
import shutil
import hashlib
import re
import sqlite3
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = PdfWriter = None

FACT_STORE = config.FACT_STORE_PATH
PROJECT = "CNTS-25165"

# 数字签名 / 加密标记（PDF 原始字节中查找）
_SIGNATURE_MARKERS = (b"adbe.pkcs7.detached", b"/SigFlags", b"/DocMDP")
_ENCRYPT_MARKER = b"/Encrypt"


# ------------------------------------------------------------
# 命名与路径
# ------------------------------------------------------------
def quarter_months(quarter: str) -> list:
    """返回季度对应的财务月份列表，例如 Q3 -> ['2026-07','2026-08','2026-09']。"""
    q = (quarter or "Q3").upper()
    months = config.QUARTER_TO_MONTHS.get(q)
    if not months:
        raise ValueError(f"未知季度: {quarter}，可选: {list(config.QUARTER_TO_MONTHS)}")
    return months


def wp_output_dir(quarter: str) -> str:
    """返回该季度的 WP 输出根目录，例如 ...\\VCTC\\Q3\\WP。"""
    return os.path.join(config.VCTC_GR_YEAR_DIR, quarter.upper(), "WP")


def canonical_name(person_key: str) -> str:
    """把 person_key 归一化为标准姓名（如 'He, Bo' / 'He Bo' -> 'He Bo'）。"""
    lookup = config.build_name_lookup(config.CNTS25165_PERSONNEL)
    key = person_key.replace(", ", " ").replace(",", " ").strip()
    low = key.lower()
    return lookup.get(low, key)


def name_token(name: str) -> str:
    """去掉姓名中的空格，用于文件名/文件夹名（'He Bo' -> 'HeBo'）。"""
    return name.replace(" ", "").replace(",", "").replace("&", "")


def name_abbr(name: str) -> str:
    """路径过长时使用的姓名缩写（'PanChengchao' -> 'PanCC'），未配置时退回全名。"""
    return config.CNTS25165_NAME_ABBR.get(name) or name_token(name)


# Windows 传统 MAX_PATH 上限为 260；超过该阈值时改用缩写，保证文件仍可在资源管理器中改名。
MAX_OUTPUT_PATH = 240


def month_short(month: str) -> str:
    """'2026-07' -> '07'。"""
    return month[5:7]


def month_label(months) -> str:
    """['07','08','09'] -> '07-09'，单月 -> '07'。"""
    ms = sorted(set(months))
    if not ms:
        return ""
    if len(ms) == 1:
        return ms[0]
    return f"{ms[0]}-{ms[-1]}"


# ------------------------------------------------------------
# 数据读取与文件定位
# ------------------------------------------------------------
def _month_folder(month: str) -> str:
    """'2026-07' -> '2607'（源 PDF 所在月份目录名）。"""
    return f"{month[2:4]}{month[5:7]}"


def load_facts(quarter: str) -> list:
    """读取该季度审核通过且为当前版本的工时事实。"""
    months = quarter_months(quarter)
    conn = sqlite3.connect(FACT_STORE)
    conn.row_factory = sqlite3.Row
    placeholders = ",".join("?" * len(months))
    rows = conn.execute(
        f"SELECT person_key, month, po_id, hours, source_id "
        f"FROM timesheet_facts "
        f"WHERE approval_status='APPROVED' AND is_current=1 "
        f"AND project_id=? AND month IN ({placeholders}) "
        f"ORDER BY month, po_id, person_key",
        [PROJECT] + months,
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def _load_source_files(source_ids) -> dict:
    """source_id -> {'path':.., 'file_name':..}。"""
    if not source_ids:
        return {}
    conn = sqlite3.connect(FACT_STORE)
    conn.row_factory = sqlite3.Row
    placeholders = ",".join("?" * len(source_ids))
    rows = conn.execute(
        f"SELECT source_id, path, file_name FROM source_files "
        f"WHERE source_id IN ({placeholders})",
        list(source_ids),
    ).fetchall()
    conn.close()
    return {r["source_id"]: dict(r) for r in rows}


def resolve_source_path(fact: dict, source_map: dict, timesheet_root: str) -> str | None:
    """按数据库登记的文件名/路径定位源 PDF（未改名时的快速路径）。"""
    s = source_map.get(fact["source_id"])
    if not s:
        return None
    file_name = s.get("file_name")
    if file_name:
        candidate = os.path.join(timesheet_root, _month_folder(fact["month"]), file_name)
        if os.path.isfile(candidate):
            return candidate
    db_path = s.get("path")
    if db_path and os.path.isfile(db_path):
        return db_path
    if file_name:
        for dirpath, _dirs, files in os.walk(timesheet_root):
            if file_name in files:
                return os.path.join(dirpath, file_name)
    return None


# ------------------------------------------------------------
# 源文件被重命名后的兜底匹配（按人员 + OCR 读出的 PO）
# ------------------------------------------------------------
SOURCE_INDEX_PATH = os.path.join(config.AUTOTOOLS_DIR, "data", "wp_source_index.json")


def _norm_token(value: str) -> str:
    """去掉空格/分隔符并转小写，用于文件名匹配。"""
    return re.sub(r"[\s\-_(),.（）()]", "", value or "").lower()


def _person_aliases(person: str) -> list:
    """返回某人的全部姓名变体（英文 + 中文）。"""
    info = config.CNTS25165_PERSONNEL.get(person, {})
    return [person, *info.get("aliases", [])]


def _file_matches_person(path: str, person: str) -> bool:
    stem = _norm_token(os.path.splitext(os.path.basename(path))[0])
    for alias in _person_aliases(person):
        token = _norm_token(alias)
        if token and token in stem:
            return True
    return False


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_index() -> dict:
    try:
        with open(SOURCE_INDEX_PATH, encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_index(index: dict) -> None:
    os.makedirs(os.path.dirname(SOURCE_INDEX_PATH), exist_ok=True)
    tmp = SOURCE_INDEX_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(index, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, SOURCE_INDEX_PATH)


def extract_source_fields(path: str, index: dict) -> dict:
    """读取 PDF 内的姓名/月份/PO/工时；结果按 SHA-256 缓存在 data/。"""
    digest = _sha256(path)
    cached = index.get(digest)
    if cached:
        return cached
    try:
        from timesheet_ingest import parse_pdf
        parsed = parse_pdf(path, use_ocr=True)
        fields = {
            "person": parsed.canonical_name or "",
            "month": parsed.month,
            "po": parsed.po_id or "",
            "hours": parsed.total_hours,
            "method": parsed.parse_method,
        }
    except Exception as exc:
        fields = {"error": str(exc)}
    index[digest] = fields
    return fields


def _month_dir_files(timesheet_root: str, month: str) -> list:
    month_dir = os.path.join(timesheet_root, _month_folder(month))
    if not os.path.isdir(month_dir):
        return []
    return [os.path.join(month_dir, f) for f in sorted(os.listdir(month_dir))
            if f.lower().endswith(".pdf")]


def resolve_group_files(person: str, month: str, po_targets, timesheet_root: str,
                        index: dict, used: set) -> tuple:
    """为该人员的某个月找到每个 PO 对应的源文件。

    返回 (found: {po: [path]}, notes: [str])。仅当数据库文件名失效时才走
    文件名/OCR 兜底；一个文件只会分配给一个 PO。
    """
    found = defaultdict(list)
    notes = []
    candidates = [p for p in _month_dir_files(timesheet_root, month)
                  if p not in used and _file_matches_person(p, person)]
    if not candidates:
        return found, notes

    if len(candidates) == 1 and len(po_targets) == 1:
        found[next(iter(po_targets))].append(candidates[0])
        return found, notes

    for path in candidates:
        fields = extract_source_fields(path, index)
        if fields.get("error"):
            notes.append(f"{os.path.basename(path)} 解析失败：{fields['error']}")
            continue
        po = fields.get("po") or ""
        # 姓名交叉校验，避免同名/误匹配
        parsed_person = fields.get("person") or ""
        if parsed_person:
            lookup = config.build_name_lookup(config.CNTS25165_PERSONNEL)
            if lookup.get(parsed_person.lower(), parsed_person) != person:
                continue
        if po in po_targets and not found[po]:
            found[po].append(path)
        elif po in po_targets:
            found[po].append(path)
        else:
            notes.append(f"{os.path.basename(path)} 的 PO={po or '未识别'} 与事实不符")
    return found, notes


def detect_signature(path: str) -> str | None:
    """返回数字签名/加密原因，或 None（可安全合并）。"""
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError:
        return "文件不可读"
    if _ENCRYPT_MARKER in data:
        return "已加密"
    for marker in _SIGNATURE_MARKERS:
        if marker in data:
            return "含电子签名"
    return None


# ------------------------------------------------------------
# 合并
# ------------------------------------------------------------
def merge_pdfs(source_paths: list, out_path: str) -> None:
    """按给定顺序把多个 PDF 合并为一个。"""
    if not PdfReader:
        raise RuntimeError("缺少 pypdf 依赖，无法合并 PDF")
    writer = PdfWriter()
    for p in source_paths:
        reader = PdfReader(p)
        for page in reader.pages:
            writer.add_page(page)
    with open(out_path, "wb") as fh:
        writer.write(fh)


# ------------------------------------------------------------
# 主流程
# ------------------------------------------------------------
def build_plan(quarter: str) -> list:
    """读取事实并按 (po, person) 分组，返回分组明细（不写文件）。"""
    facts = load_facts(quarter)
    source_map = _load_source_files({f["source_id"] for f in facts})
    timesheet_root = config.PROJECT_PROFILE_OVERRIDES[PROJECT]["source_dirs"]["timesheet_pdf"]
    index = _load_index()
    index_dirty = False

    # 先做数据库文件名快速定位，记录失败的 (person, month) 需要兜底匹配
    resolved = {}
    unresolved_keys = set()
    for f in facts:
        person = canonical_name(f["person_key"])
        path = resolve_source_path(f, source_map, timesheet_root)
        resolved[f["source_id"]] = path
        if not path:
            unresolved_keys.add((person, f["month"]))

    # 对缺失的 (person, month) 做文件名 + OCR 兜底配对
    fallback_notes = defaultdict(list)
    for person, month in sorted(unresolved_keys):
        targets = {f["po_id"] for f in facts
                   if canonical_name(f["person_key"]) == person and f["month"] == month}
        used = {p for p in resolved.values() if p}
        before = {k: dict(v) for k, v in index.items()}
        found, notes = resolve_group_files(person, month, targets, timesheet_root, index, used)
        if index != before:
            index_dirty = True
        for po, paths in found.items():
            for f in facts:
                if (f["month"] == month and f["po_id"] == po
                        and canonical_name(f["person_key"]) == person
                        and not resolved.get(f["source_id"])):
                    resolved[f["source_id"]] = paths[0]
                    break
        for n in notes:
            fallback_notes[(person, month)].append(n)

    groups = defaultdict(list)
    for f in facts:
        po = f["po_id"] or "NOPO"
        person = canonical_name(f["person_key"])
        groups[(po, person)].append(f)

    plan_groups = []
    for (po, person), fs in sorted(groups.items()):
        entries = []
        for f in sorted(fs, key=lambda x: x["month"]):
            entries.append({"month": f["month"], "source_id": f["source_id"],
                            "path": resolved.get(f["source_id"])})
        months = sorted({e["month"] for e in entries})
        manual_reasons = []
        for e in entries:
            if not e["path"]:
                manual_reasons.append(f"{month_short(e['month'])}月源文件缺失")
                continue
            reason = detect_signature(e["path"])
            if reason:
                manual_reasons.append(f"{month_short(e['month'])}月源文件{reason}")
        notes = []
        for e in entries:
            notes.extend(fallback_notes.get((person, e["month"]), []))
        plan_groups.append({
            "po": po,
            "person": person,
            "token": name_token(person),
            "abbr": name_abbr(person),
            "months": months,
            "entries": entries,
            "manual_reasons": sorted(set(manual_reasons)),
            "notes": sorted(set(notes)),
        })

    if index_dirty:
        try:
            _save_index(index)
        except OSError:
            pass
    return plan_groups


def run(quarter: str = "Q3", dry_run: bool = False) -> dict:
    """执行 WP 合并，返回结构化结果。"""
    quarter = (quarter or "Q3").upper()
    plan_groups = build_plan(quarter)
    out_root = wp_output_dir(quarter)

    # 按 PO 归集，计算每个 PO 的文件夹名与月份范围
    by_po = defaultdict(list)
    for g in plan_groups:
        by_po[g["po"]].append(g)

    def _person_tokens(gs, key: str) -> dict:
        """person -> token；缩写后重名的人员回退到完整姓名，避免文件互相覆盖。"""
        by_token = defaultdict(list)
        for g in gs:
            by_token[g[key]].append(g["person"])
        tokens = {}
        for g in gs:
            token = g[key]
            if len(set(by_token[token])) > 1:
                token = g["token"]
            tokens[g["person"]] = token
        return tokens

    def _folder_name(po: str, months, tokens: dict) -> str:
        persons = sorted(set(tokens.values()))
        return f"IAV_{'&'.join(persons)}_{month_label(month_short(m) for m in months)}_PO{po}"

    # 先用完整姓名拼路径；超过 MAX_OUTPUT_PATH 时整组改用缩写姓名。
    folder_names, token_maps = {}, {}
    for po, gs in by_po.items():
        all_months = sorted({m for g in gs for m in g["months"]})
        full_tokens = _person_tokens(gs, "token")
        full_folder = _folder_name(po, all_months, full_tokens)
        longest_path = max(
            len(os.path.join(
                out_root, full_folder,
                f"IAV_{g['token']}_{month_label(month_short(m) for m in g['months'])}_PO{po}.pdf",
            ))
            for g in gs
        )
        if longest_path > MAX_OUTPUT_PATH:
            token_maps[po] = _person_tokens(gs, "abbr")
            folder_names[po] = _folder_name(po, all_months, token_maps[po])
        else:
            token_maps[po] = full_tokens
            folder_names[po] = full_folder

    merged, manual, errors = [], [], []
    for g in plan_groups:
        po, person = g["po"], g["person"]
        folder = folder_names[po]
        file_label = month_label(month_short(m) for m in g["months"])
        out_file = f"IAV_{token_maps[po][person]}_{file_label}_PO{po}.pdf"
        out_path = os.path.join(out_root, folder, out_file)

        if g["manual_reasons"]:
            manual.append({
                "po": po,
                "person": person,
                "reason": "；".join(g["manual_reasons"]),
                "source_files": [os.path.basename(e["path"]) if e["path"] else "缺失" for e in g["entries"]],
                "notes": g.get("notes", []),
                "expected_file": os.path.join(folder, out_file),
            })
            continue

        source_paths = [e["path"] for e in g["entries"]]
        if dry_run:
            merged.append({"po": po, "person": person, "folder": folder, "file": out_file,
                           "months": g["months"], "sources": [os.path.basename(p) for p in source_paths],
                           "output": out_path})
            continue

        try:
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            if len(source_paths) == 1:
                shutil.copyfile(source_paths[0], out_path)
            else:
                merge_pdfs(source_paths, out_path)
            merged.append({"po": po, "person": person, "folder": folder, "file": out_file,
                           "months": g["months"], "sources": [os.path.basename(p) for p in source_paths],
                           "output": out_path})
        except Exception as exc:
            errors.append({"po": po, "person": person, "reason": str(exc)})

    return {
        "quarter": quarter,
        "output_root": out_root,
        "dry_run": dry_run,
        "groups_total": len(plan_groups),
        "merged": merged,
        "manual": manual,
        "errors": errors,
    }


def _print_report(result: dict) -> None:
    print("=" * 60)
    print(f"  GR WP 合并结果（{result['quarter']}）")
    print(f"  输出目录: {result['output_root']}")
    print(f"  分组总数: {result['groups_total']} | 成功 {len(result['merged'])} | "
          f"需手动 {len(result['manual'])} | 失败 {len(result['errors'])}")
    print("=" * 60)
    if result["merged"]:
        print("\n[已生成]")
        for m in result["merged"]:
            print(f"  ✓ {m['folder']}/{m['file']}")
    if result["manual"]:
        print("\n[需手动处理]")
        for m in result["manual"]:
            print(f"  ✗ PO {m['po']} / {m['person']}：{m['reason']}")
            for f in m["source_files"]:
                print(f"       - {f}")
    if result["errors"]:
        print("\n[合并失败]")
        for e in result["errors"]:
            print(f"  ⚠ PO {e['po']} / {e['person']}：{e['reason']}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="GR WP 合并器")
    parser.add_argument("--quarter", default="Q3", help="季度（默认 Q3）")
    parser.add_argument("--dry-run", action="store_true", help="仅预览，不写文件")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出")
    args = parser.parse_args()

    if sys.stdout.encoding != "utf-8":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    result = run(quarter=args.quarter, dry_run=args.dry_run)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        _print_report(result)
