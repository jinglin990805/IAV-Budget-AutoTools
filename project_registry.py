# -*- coding: utf-8 -*-
"""跨项目注册与健康检查。

该模块只负责把 config.py 中的项目差异整理成稳定的 ProjectProfile，
并发现输入文件；不会导入事实，也不会修改任何业务 Excel。
"""

from __future__ import annotations

import argparse
import json
import os
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import config


DEFAULT_REQUIRED_SOURCES = (
    "timesheet_pdf",
    "monthly_closing",
    "po_file",
)


@dataclass(frozen=True)
class ProjectProfile:
    """项目级差异；核心事实和审核逻辑不得依赖具体项目号。"""

    project_id: str
    name: str
    root_dir: str
    budget_report_dir: str
    budget_report_patterns: tuple[str, ...] = ()
    template_path: str = ""
    po_usage_template_path: str = ""
    acceptance_template_path: str = ""
    acceptance_mapping: dict[str, Any] = field(default_factory=dict)
    source_dirs: dict[str, str] = field(default_factory=dict)
    fiscal_year_start_month: int = 1
    currency: str = "CNY"
    hours_per_day: float = 8.0
    display_unit: str = "hour"
    allocation_mode: str = "usage"
    cost_centers: tuple[str, ...] = ()
    required_sources: tuple[str, ...] = DEFAULT_REQUIRED_SOURCES
    deliverables: dict[str, bool] = field(default_factory=dict)
    macro_policy: str = "xlsx_only"

    def to_dict(self) -> dict[str, Any]:
        result = asdict(self)
        result["budget_report_patterns"] = list(self.budget_report_patterns)
        result["cost_centers"] = list(self.cost_centers)
        result["required_sources"] = list(self.required_sources)
        return result


def _merge_dict(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            nested = dict(merged[key])
            nested.update(value)
            merged[key] = nested
        else:
            merged[key] = value
    return merged


def _infer_root_dir(project_id: str, project_config: dict[str, Any]) -> str:
    if project_config.get("root_dir"):
        return project_config["root_dir"]
    budget_dir = project_config.get("budget_report_dir", "")
    if project_id in {"CN04-22028", "CN71-21043"}:
        return os.path.join(config.ROOT, "CN04-22028 & CN71-21043")
    return str(Path(budget_dir).parent if budget_dir else Path(config.ROOT))


def build_profile(project_id: str) -> ProjectProfile:
    """将旧 PROJECTS 基础配置与新 Profile 覆盖项合并。"""
    if project_id not in config.PROJECTS:
        raise KeyError(f"Unknown project: {project_id}")

    base = dict(config.PROJECTS[project_id])
    override = dict(getattr(config, "PROJECT_PROFILE_OVERRIDES", {}).get(project_id, {}))
    merged = _merge_dict(base, override)
    template_path = merged.get("template_path", "")
    if not template_path:
        template_path = getattr(config, "BUDGET_REPORT_TEMPLATE_PATHS", {}).get(project_id, "")

    patterns = merged.get("budget_report_patterns") or (
        "*Budget*Report*.xlsx",
        "*Budget*Report*.xlsm",
    )
    required = merged.get("required_sources") or DEFAULT_REQUIRED_SOURCES

    return ProjectProfile(
        project_id=project_id,
        name=merged.get("name", project_id),
        root_dir=_infer_root_dir(project_id, merged),
        budget_report_dir=merged.get("budget_report_dir", ""),
        budget_report_patterns=tuple(patterns),
        template_path=template_path,
        po_usage_template_path=str(merged.get("po_usage_template_path", "")),
        acceptance_template_path=str(merged.get("acceptance_template_path", "")),
        acceptance_mapping=dict(merged.get("acceptance_mapping", {})),
        source_dirs=dict(merged.get("source_dirs", {})),
        fiscal_year_start_month=int(merged.get("fiscal_year_start_month", 1)),
        currency=str(merged.get("currency", "CNY")),
        hours_per_day=float(merged.get("hours_per_day", config.HOURS_PER_DAY)),
        display_unit=str(merged.get("display_unit", "hour")),
        allocation_mode=str(merged.get("allocation_mode", "usage")),
        cost_centers=tuple(str(x) for x in merged.get("cost_centers", ())),
        required_sources=tuple(str(x) for x in required),
        deliverables=dict(merged.get("deliverables", {})),
        macro_policy=str(merged.get("macro_policy", "xlsx_only")),
    )


def list_profiles() -> list[ProjectProfile]:
    return [build_profile(project_id) for project_id in sorted(config.PROJECTS)]


def _is_excluded_budget_file(path: Path) -> bool:
    name = path.name.lower()
    excluded_tokens = ("template", "backup", "_diff", "~$", "po_usage_status")
    return any(token in name for token in excluded_tokens)


def discover_budget_reports(profile: ProjectProfile) -> list[Path]:
    """按 Profile 的文件名模式发现报表，不打开或修改文件。"""
    directory = Path(profile.budget_report_dir)
    if not directory.exists():
        return []
    found: set[Path] = set()
    for pattern in profile.budget_report_patterns:
        found.update(path for path in directory.rglob(pattern) if not _is_excluded_budget_file(path))
    return sorted(found, key=lambda item: (item.stat().st_mtime_ns, str(item).lower()), reverse=True)


def latest_budget_report(profile: ProjectProfile) -> Path | None:
    reports = discover_budget_reports(profile)
    return reports[0] if reports else None


def profile_health(profile: ProjectProfile) -> dict[str, Any]:
    latest = latest_budget_report(profile)
    source_status = {
        name: {
            "path": path,
            "exists": Path(path).exists(),
        }
        for name, path in profile.source_dirs.items()
    }
    issues: list[str] = []
    if not Path(profile.root_dir).exists():
        issues.append("project_root_missing")
    if not Path(profile.budget_report_dir).exists():
        issues.append("budget_report_dir_missing")
    if latest is None:
        issues.append("budget_report_not_found")
    for source_name in profile.required_sources:
        status = source_status.get(source_name)
        if not status or not status["exists"]:
            issues.append(f"required_source_missing:{source_name}")
    if profile.template_path and not Path(profile.template_path).exists():
        issues.append("configured_template_missing")
    if profile.po_usage_template_path and not Path(profile.po_usage_template_path).exists():
        issues.append("configured_po_usage_template_missing")
    if not 1 <= profile.fiscal_year_start_month <= 12:
        issues.append("invalid_fiscal_year_start_month")
    if profile.hours_per_day <= 0:
        issues.append("invalid_hours_per_day")
    return {
        "project_id": profile.project_id,
        "name": profile.name,
        "healthy": not issues,
        "issues": issues,
        "latest_budget_report": str(latest) if latest else None,
        "budget_report_count": len(discover_budget_reports(profile)),
        "source_status": source_status,
        "profile": profile.to_dict(),
        "checked_at": datetime.now().isoformat(timespec="seconds"),
    }


def registry_health(project_id: str | None = None) -> dict[str, Any]:
    profiles = [build_profile(project_id)] if project_id else list_profiles()
    reports = [profile_health(profile) for profile in profiles]
    return {
        "healthy": all(item["healthy"] for item in reports),
        "projects": reports,
        "checked_at": datetime.now().isoformat(timespec="seconds"),
    }


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="IAV cross-project registry health check")
    sub = parser.add_subparsers(dest="command", required=True)
    list_cmd = sub.add_parser("list", help="list project profiles")
    list_cmd.add_argument("--json", action="store_true")
    health_cmd = sub.add_parser("health", help="check project paths and inputs")
    health_cmd.add_argument("--project", default=None)
    health_cmd.add_argument("--json", action="store_true")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    if args.command == "list":
        payload = [profile.to_dict() for profile in list_profiles()]
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            for profile in list_profiles():
                print(f"{profile.project_id}\t{profile.name}\t{profile.display_unit}\t{profile.allocation_mode}")
        return 0

    payload = registry_health(args.project)
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        for item in payload["projects"]:
            status = "OK" if item["healthy"] else "CHECK"
            print(f"[{status}] {item['project_id']} {item['name']}")
            print(f"  Budget: {item['latest_budget_report'] or 'NOT FOUND'}")
            if item["issues"]:
                print(f"  Issues: {', '.join(item['issues'])}")
    return 0 if payload["healthy"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
