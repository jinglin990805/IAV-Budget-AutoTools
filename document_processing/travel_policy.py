"""Deterministic Travel policy rules and auditable calculation traces.

This module deliberately does not read documents or make OCR decisions.  It
accepts the already configured policy and explicit business inputs, then
returns a serialisable decision with its calculation steps and exceptions.
"""

from __future__ import annotations

import math
from typing import Any, Iterable, Optional


LEVEL_ORDER = ("1", "1.5", "2", "international")
STATUS_ORDER = {"PASS": 0, "WARNING": 1, "BLOCKED": 2}


def _rounded(value: float) -> float:
    return round(float(value), 2)


class TravelPolicyEngine:
    """Evaluate configured travel standards without hidden fallbacks.

    ``config`` may be the full ``travel_policies.json`` envelope or one policy
    object.  The result shape is JSON-safe so it can be returned by the API
    and displayed in the review workbench without another transformation.
    """

    def __init__(self, config: dict[str, Any], policy_name: Optional[str] = None):
        if not isinstance(config, dict):
            raise ValueError("差标配置必须是对象")
        if isinstance(config.get("policies"), dict):
            policies = config["policies"]
            selected = policy_name or config.get("active_policy")
            if not selected:
                selected = next(iter(policies), None)
            policy = policies.get(selected) if selected else None
        else:
            policies = {policy_name or "VGC": config}
            selected = policy_name or "VGC"
            policy = config
        if not isinstance(policy, dict):
            raise ValueError(f"差标不存在: {selected}")
        self.policy_name = str(selected)
        self.policy = policy
        self.levels = policy.get("levels") if isinstance(policy.get("levels"), dict) else {}
        if not self.levels:
            raise ValueError(f"差标 {self.policy_name} 缺少城市等级")

    def _issue(self, code: str, severity: str, message: str, field: str) -> dict[str, str]:
        return {"code": code, "severity": severity, "message": message, "field": field}

    def _status(self, issues: Iterable[dict[str, str]]) -> str:
        status = "PASS"
        for issue in issues:
            if STATUS_ORDER.get(issue.get("severity", "WARNING"), 1) > STATUS_ORDER[status]:
                status = issue.get("severity", "WARNING")
        return status

    def resolve_city(self, city: str) -> dict[str, Any]:
        """Resolve a city to one configured level, recording fallback use."""
        raw_city = "" if city is None else str(city).strip()
        city_key = raw_city.casefold()
        for level_key in (*LEVEL_ORDER, *(key for key in self.levels if key not in LEVEL_ORDER)):
            definition = self.levels.get(level_key)
            if not isinstance(definition, dict):
                continue
            cities = definition.get("cities", [])
            if any(str(item).strip().casefold() == city_key for item in cities):
                return {
                    "city": raw_city,
                    "level": str(level_key),
                    "label": str(definition.get("label") or f"Layer {level_key}"),
                    "known_city": True,
                    "accommodation_per_night": _rounded(float(definition.get("accommodation", 0))),
                    "allowance_per_day": _rounded(float(definition.get("allowance", 0))),
                    "issues": [],
                }

        fallback = str(self.policy.get("unknown_level", "2"))
        if fallback not in self.levels:
            fallback = "2" if "2" in self.levels else next(iter(self.levels), "")
        definition = self.levels.get(fallback, {})
        issue = self._issue(
            "UNKNOWN_CITY_POLICY_FALLBACK",
            "WARNING",
            f"城市 {raw_city or '（空）'} 未在差标城市清单中，已采用 {fallback} 层默认标准。",
            "city",
        )
        return {
            "city": raw_city,
            "level": fallback,
            "label": str(definition.get("label") or f"Layer {fallback}"),
            "known_city": False,
            "accommodation_per_night": _rounded(float(definition.get("accommodation", 0))),
            "allowance_per_day": _rounded(float(definition.get("allowance", 0))),
            "issues": [issue],
        }

    def _invalid_result(self, rule_id: str, values: dict[str, Any], issues: list[dict[str, str]]) -> dict[str, Any]:
        return {
            "rule_id": rule_id,
            "policy_name": self.policy_name,
            "status": self._status(issues),
            "input": values,
            "output": {},
            "calculation_trace": [{
                "step": "validate_input",
                "rule_id": f"{rule_id}.INPUT_VALIDATION",
                "input": values,
                "output": {"valid": False},
            }],
            "issues": issues,
        }

    def calculate_accommodation(self, city: str, amount: Any, nights: Any) -> dict[str, Any]:
        """Apply the configured per-night cap to one accommodation claim."""
        values = {"city": city, "claimed_amount": amount, "nights": nights}
        issues: list[dict[str, str]] = []
        try:
            claimed = float(amount)
            if not math.isfinite(claimed) or claimed < 0:
                raise ValueError
        except (TypeError, ValueError):
            issues.append(self._issue("INVALID_ACCOMMODATION_AMOUNT", "BLOCKED", "住宿金额必须是非负数字。", "amount"))
            return self._invalid_result("TRAVEL.HOTEL_CAP", values, issues)
        if isinstance(nights, bool) or not isinstance(nights, (int, float)) or not math.isfinite(float(nights)):
            issues.append(self._issue("INVALID_ACCOMMODATION_NIGHTS", "BLOCKED", "住宿晚数必须是非负整数。", "nights"))
            return self._invalid_result("TRAVEL.HOTEL_CAP", values, issues)
        if int(nights) != float(nights) or int(nights) < 1:
            issues.append(self._issue("INVALID_ACCOMMODATION_NIGHTS", "BLOCKED", "住宿晚数必须是正整数。", "nights"))
            return self._invalid_result("TRAVEL.HOTEL_CAP", values, issues)

        nights_value = int(nights)
        resolution = self.resolve_city(city)
        issues.extend(resolution["issues"])
        per_night = resolution["accommodation_per_night"]
        maximum = _rounded(per_night * nights_value)
        allowed = _rounded(min(claimed, maximum))
        over_limit = _rounded(max(claimed - allowed, 0.0))
        if over_limit > 0:
            issues.append(self._issue(
                "ACCOMMODATION_OVER_LIMIT",
                "WARNING",
                f"住宿申报 {claimed:.2f} 元超过 {resolution['level']} 层标准 {per_night:.2f} 元/晚 × {nights_value} 晚，超标 {over_limit:.2f} 元。",
                "amount",
            ))
        return {
            "rule_id": "TRAVEL.HOTEL_CAP",
            "policy_name": self.policy_name,
            "status": self._status(issues),
            "input": {**values, "claimed_amount": _rounded(claimed), "nights": nights_value},
            "output": {
                "city": resolution["city"],
                "level": resolution["level"],
                "known_city": resolution["known_city"],
                "limit_per_night": per_night,
                "maximum_allowed": maximum,
                "allowed_amount": allowed,
                "over_limit": over_limit,
            },
            "calculation_trace": [
                {
                    "step": "resolve_city",
                    "rule_id": "TRAVEL.CITY_LEVEL",
                    "input": {"city": city, "policy_name": self.policy_name},
                    "output": {
                        "level": resolution["level"],
                        "known_city": resolution["known_city"],
                        "limit_per_night": per_night,
                    },
                },
                {
                    "step": "calculate_nightly_limit",
                    "rule_id": "TRAVEL.HOTEL_NIGHTLY_LIMIT",
                    "input": {"limit_per_night": per_night, "nights": nights_value},
                    "output": {"maximum_allowed": maximum},
                },
                {
                    "step": "apply_claim_cap",
                    "rule_id": "TRAVEL.HOTEL_CAP",
                    "input": {"claimed_amount": _rounded(claimed), "maximum_allowed": maximum},
                    "output": {"allowed_amount": allowed, "over_limit": over_limit},
                },
            ],
            "issues": issues,
        }

    def calculate_allowance(
        self,
        city: str,
        trip_days: Optional[Any] = None,
        allowance_days: Optional[Any] = None,
    ) -> dict[str, Any]:
        """Calculate daily allowance; the first trip day is not payable by default."""
        values = {"city": city, "trip_days": trip_days, "allowance_days": allowance_days}
        issues: list[dict[str, str]] = []
        if trip_days is None:
            trip_days_value = 0
        elif isinstance(trip_days, bool) or not isinstance(trip_days, (int, float)) or int(trip_days) != float(trip_days):
            issues.append(self._issue("INVALID_TRIP_DAYS", "BLOCKED", "行程天数必须是非负整数。", "trip_days"))
            return self._invalid_result("TRAVEL.DAILY_ALLOWANCE", values, issues)
        else:
            trip_days_value = int(trip_days)
        if trip_days_value < 0:
            issues.append(self._issue("INVALID_TRIP_DAYS", "BLOCKED", "行程天数必须是非负整数。", "trip_days"))
            return self._invalid_result("TRAVEL.DAILY_ALLOWANCE", values, issues)

        if allowance_days is None:
            allowance_days_value = max(trip_days_value - 1, 0)
        elif (
            isinstance(allowance_days, bool)
            or not isinstance(allowance_days, (int, float))
            or int(allowance_days) != float(allowance_days)
            or int(allowance_days) < 0
        ):
            issues.append(self._issue("INVALID_ALLOWANCE_DAYS", "BLOCKED", "补贴天数必须是非负整数。", "allowance_days"))
            return self._invalid_result("TRAVEL.DAILY_ALLOWANCE", values, issues)
        else:
            allowance_days_value = int(allowance_days)

        resolution = self.resolve_city(city)
        issues.extend(resolution["issues"])
        per_day = resolution["allowance_per_day"]
        total = _rounded(per_day * allowance_days_value)
        return {
            "rule_id": "TRAVEL.DAILY_ALLOWANCE",
            "policy_name": self.policy_name,
            "status": self._status(issues),
            "input": {**values, "trip_days": trip_days_value, "allowance_days": allowance_days_value},
            "output": {
                "city": resolution["city"],
                "level": resolution["level"],
                "known_city": resolution["known_city"],
                "allowance_per_day": per_day,
                "eligible_days": allowance_days_value,
                "allowed_amount": total,
            },
            "calculation_trace": [
                {
                    "step": "resolve_city",
                    "rule_id": "TRAVEL.CITY_LEVEL",
                    "input": {"city": city, "policy_name": self.policy_name},
                    "output": {
                        "level": resolution["level"],
                        "known_city": resolution["known_city"],
                        "allowance_per_day": per_day,
                    },
                },
                {
                    "step": "calculate_allowance_days",
                    "rule_id": "TRAVEL.ALLOWANCE_ELIGIBILITY",
                    "input": {"trip_days": trip_days_value, "first_day_excluded": True},
                    "output": {"eligible_days": allowance_days_value},
                },
                {
                    "step": "calculate_allowance_total",
                    "rule_id": "TRAVEL.DAILY_ALLOWANCE",
                    "input": {"allowance_per_day": per_day, "eligible_days": allowance_days_value},
                    "output": {"allowed_amount": total},
                },
            ],
            "issues": issues,
        }

