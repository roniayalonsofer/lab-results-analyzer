"""
core/file_limits.py
-------------------
The LOD and LOQ values read from a lab file, for the app's limits section.
"""

# Analysis types behind each group of fields in the limits section
_METALS = ("SOIL_METALS",)
_VOC = ("SOIL_VOC", "SOIL_SVOC")


def _single_value(records: list[dict], limit: str) -> float | None:
    """The limit shared by all these records, or None if there are several or none."""
    values = {r[limit] for r in records if r.get(limit)}
    return values.pop() if len(values) == 1 else None


def field_values(records: list[dict]) -> dict[str, float]:
    """Field key → value, for each field whose records share one value in the file."""
    by_field = {
        "metals": [r for r in records if r.get("analysis_type") in _METALS],
        "voc":    [r for r in records if r.get("analysis_type") in _VOC],
    }
    tph = [r for r in records if r.get("analysis_type") == "SOIL_TPH"]
    by_field["dro"] = [r for r in tph if r.get("compound") == "DRO"]
    by_field["oro"] = [r for r in tph if r.get("compound") == "ORO"]
    by_field["tph"] = [r for r in tph if r.get("compound") in ("TPH", "Total TPH")]

    values: dict[str, float] = {}
    for field, group in by_field.items():
        for limit in ("lod", "loq"):
            value = _single_value(group, limit)
            if value is not None:
                values[f"ov_{limit}_{field}"] = value
    return values


def compound_limits(records: list[dict]) -> list[dict]:
    """One row per compound with the LOD and LOQ read from the file."""
    rows: dict[tuple, dict] = {}
    for r in records:
        if r.get("lod") is None and r.get("loq") is None:
            continue
        rows.setdefault((r.get("analysis_type"), r.get("compound")), {
            "תרכובת": r.get("compound"),
            "CAS": r.get("cas") or "",
            "LOD": r.get("lod"),
            "LOQ": r.get("loq"),
            "יחידות": r.get("unit") or "",
        })
    return list(rows.values())
