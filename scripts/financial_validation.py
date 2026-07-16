#!/usr/bin/env python3
"""Structured metric records and deterministic financial validation rules."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict, dataclass
from math import isfinite


@dataclass(frozen=True)
class MetricRecord:
    company: str
    period: str
    indicator: str
    raw_value: str
    raw_unit: str
    normalized_value: float | None
    normalized_unit: str
    source_file: str
    source_location: str
    status: str

    def to_dict(self) -> dict:
        return asdict(self)


def _issue(severity: str, category: str, record: MetricRecord | None, message: str, **details) -> dict:
    payload = {
        "severity": severity,
        "category": category,
        "company": record.company if record else "",
        "period": record.period if record else "",
        "indicator": record.indicator if record else "",
        "message": message,
    }
    payload.update(details)
    return payload


def validate_records(
    records: list[MetricRecord],
    margin_tolerance: float = 0.5,
    net_margin_tolerance: float = 0.2,
    balance_tolerance_pct: float = 1.0,
    yoy_warning_pct: float = 100.0,
) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []
    missing: list[dict] = []
    conversions: list[dict] = []
    by_entity: dict[tuple[str, str], dict[str, MetricRecord]] = defaultdict(dict)

    for record in records:
        by_entity[(record.company, record.period)][record.indicator] = record
        if record.normalized_value is None:
            missing.append(_issue("info", "missing", record, "指标未披露或原始值为空"))
        elif not isfinite(record.normalized_value):
            errors.append(_issue("error", "non_finite", record, "标准化数值不是有限数"))
        elif record.status == "converted":
            conversions.append(_issue(
                "info", "unit_conversion", record,
                f"{record.raw_unit} → {record.normalized_unit}",
                raw_value=record.raw_value,
                normalized_value=record.normalized_value,
            ))
        if record.normalized_unit == "%" and record.normalized_value is not None and abs(record.normalized_value) > 1000:
            errors.append(_issue("error", "ratio_range", record, "比例绝对值超过1000%，疑似单位或数量级错误"))

    def check_ratio(
        metrics: dict[str, MetricRecord], numerator: str, denominator: str,
        reported: str, tolerance: float,
    ) -> None:
        parts = [metrics.get(numerator), metrics.get(denominator), metrics.get(reported)]
        if not all(parts) or any(item.normalized_value is None for item in parts):
            return
        num, den, rep = parts
        if den.normalized_value == 0:
            return
        expected = num.normalized_value / den.normalized_value * 100
        difference = abs(expected - rep.normalized_value)
        if difference > tolerance:
            warnings.append(_issue(
                "warning", "recalculation", rep,
                f"重算值 {expected:.2f}% 与披露值 {rep.normalized_value:.2f}% 相差 {difference:.2f}pct",
                expected=expected,
                reported=rep.normalized_value,
                tolerance=tolerance,
                evidence=[numerator, denominator, reported],
            ))

    for (_, _), metrics in by_entity.items():
        check_ratio(metrics, "毛利", "营业收入", "毛利率", margin_tolerance)
        check_ratio(metrics, "归母净利润", "营业收入", "归母净利率", net_margin_tolerance)
        check_ratio(metrics, "核心净利润", "营业收入", "核心净利率", net_margin_tolerance)

        assets = metrics.get("总资产")
        liabilities = metrics.get("总负债")
        equity = metrics.get("所有者权益") or metrics.get("总权益")
        if all((assets, liabilities, equity)) and all(
            item.normalized_value is not None for item in (assets, liabilities, equity)
        ) and assets.normalized_value:
            difference = abs(assets.normalized_value - liabilities.normalized_value - equity.normalized_value)
            difference_pct = difference / abs(assets.normalized_value) * 100
            if difference_pct > balance_tolerance_pct:
                warnings.append(_issue(
                    "warning", "balance_equation", assets,
                    f"资产与负债加权益相差 {difference_pct:.2f}%",
                    difference=difference,
                    difference_pct=difference_pct,
                    tolerance=balance_tolerance_pct,
                ))

    by_series: dict[tuple[str, str], list[MetricRecord]] = defaultdict(list)
    for record in records:
        if record.normalized_value is not None:
            by_series[(record.company, record.indicator)].append(record)
    for (_, _), series in by_series.items():
        ordered = sorted(series, key=lambda item: item.period)
        for previous, current in zip(ordered, ordered[1:]):
            if previous.normalized_value == 0 or previous.normalized_unit != current.normalized_unit:
                continue
            change_pct = (current.normalized_value - previous.normalized_value) / abs(previous.normalized_value) * 100
            if abs(change_pct) > yoy_warning_pct:
                warnings.append(_issue(
                    "warning", "large_period_change", current,
                    f"较 {previous.period} 变化 {change_pct:.1f}%",
                    previous_period=previous.period,
                    previous_value=previous.normalized_value,
                    current_value=current.normalized_value,
                    change_pct=change_pct,
                    threshold=yoy_warning_pct,
                ))

    return {
        "schema_version": 1,
        "status": "failed" if errors else "passed_with_warnings" if warnings else "passed",
        "stats": {
            "records": len(records),
            "errors": len(errors),
            "warnings": len(warnings),
            "missing": len(missing),
            "conversions": len(conversions),
        },
        "errors": errors,
        "warnings": warnings,
        "missing": missing,
        "conversions": conversions,
    }
