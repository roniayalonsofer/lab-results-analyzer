"""
core/compound_choices.py
------------------------
Compounds whose lab name doesn't clearly match the threshold table, and the
user's answers about them. The app asks before building a report, so the
Excel, Word and preview outputs all come out with one definite threshold.
"""

from core.cas_lookup import fuzzy_name_to_cas, name_matches_cas

# Answer meaning "none of the candidates": the compound has no threshold.
NO_THRESHOLD = "__no_threshold__"

# Results with no threshold comparison, or matched by their own TPH fraction rules.
_NOT_ASKED_TYPES = {"LOWFLOW", "GW_FIELD_PARAMS", "SOIL_TPH"}


def _name_confirms(name: str, cas: str, tm) -> bool:
    """True when a CAS looked up from the name is certainly this compound."""
    if not cas:
        return False
    table_name = tm._thresh_names().get(cas)
    return name_matches_cas(name, cas) or bool(
        table_name and tm.name_match_confidence(name, table_name) == "exact")


def find_ambiguous_compounds(records: list[dict], tm, threshold_keys: list[str]) -> list[dict]:
    """
    Return {'compound', 'cas', 'candidates'} per compound that needs an answer.

    The CAS decides: with a CAS in the threshold table the match is settled, and a
    CAS the table doesn't have means no threshold. A compound with no CAS is asked
    about when the table has similar names. Always asked: a summed result, and a
    CAS the table lacks when the name is a compound we know under another CAS.
    """
    ambiguous: list[dict] = []
    if not threshold_keys:
        return ambiguous
    seen: set[str] = set()
    for r in records:
        name = r["compound"]
        if (name in seen or r.get("analysis_type") in _NOT_ASKED_TYPES
                or (r.get("value") is None and r.get("loq") is None)
                or name.strip().lower() in tm._TPH_ALIASES):
            continue
        seen.add(name)
        own_cas = str(r.get("cas") or "").strip()
        # Same CAS enrichment as the Excel builder
        cas = (own_cas or tm.get_cas_by_name(name)
               or fuzzy_name_to_cas(name.strip().lower()) or "")
        confs = [tm.get_threshold_with_confidence(cas, k, compound_name=name)[1]
                 for k in threshold_keys]
        uncertain = "uncertain" in confs
        confirmed = own_cas or _name_confirms(name, cas, tm) or (not cas and "exact" in confs)
        if confirmed and not uncertain:
            continue
        candidates = tm.threshold_candidates(name, cas, threshold_keys)
        if candidates or uncertain:
            ambiguous.append({"compound": name, "cas": cas, "candidates": candidates})
    return ambiguous


def apply_compound_choices(records: list[dict], choices: dict[str, str]) -> list[dict]:
    """Copy *records* with the chosen CAS, so the report shows what was chosen."""
    return [
        {**r, "cas": choices[r["compound"]]}
        if choices.get(r["compound"], NO_THRESHOLD) != NO_THRESHOLD else r
        for r in records
    ]


class ResolvedThresholds:
    """A ThresholdManager that answers from *choices* for the compounds in it."""

    def __init__(self, tm, choices: dict[str, str]):
        self._tm = tm
        self._choices = choices

    def __getattr__(self, name):
        return getattr(self._tm, name)

    def _chosen_value(self, compound_name: str, threshold_key: str) -> float | None:
        cas = self._choices[compound_name]
        return None if cas == NO_THRESHOLD else self._tm.get_threshold(cas, threshold_key)

    def get_threshold_with_name(self, cas: str, threshold_key: str,
                                compound_name: str = "") -> float | None:
        if compound_name in self._choices:
            return self._chosen_value(compound_name, threshold_key)
        return self._tm.get_threshold_with_name(cas, threshold_key, compound_name)

    def get_threshold_with_confidence(self, cas: str, threshold_key: str,
                                      compound_name: str = "") -> tuple[float | None, str]:
        if compound_name in self._choices:
            val = self._chosen_value(compound_name, threshold_key)
            return (val, "exact" if val is not None else "none")
        return self._tm.get_threshold_with_confidence(cas, threshold_key, compound_name)
