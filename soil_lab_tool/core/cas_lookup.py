"""
core/cas_lookup.py
------------------
Maps chemical names (Hebrew / English / symbols) to CAS numbers.
"""

CHEMICAL_MAP: dict[str, str] = {
    # BTEX
    "benzene":              "71-43-2",
    "בנזן":                 "71-43-2",
    "toluene":              "108-88-3",
    "טולואן":               "108-88-3",
    "ethylbenzene":         "100-41-4",
    "ethyl benzene":        "100-41-4",
    "אתיל-בנזן":            "100-41-4",
    "אתילבנזן":             "100-41-4",
    "xylene":               "1330-20-7",
    "xylenes":              "1330-20-7",
    "קסילן":                "1330-20-7",
    "קסילנים":              "1330-20-7",
    "o-xylene":             "95-47-6",
    "m-xylene":             "108-38-3",
    "p-xylene":             "106-42-3",
    # MTBE / oxygenates
    "mtbe":                 "1634-04-4",
    "methyl tert-butyl ether": "1634-04-4",
    "etbe":                 "637-92-3",
    "tame":                 "994-05-8",
    # Chlorinated solvents
    "tce":                  "79-01-6",
    "trichloroethylene":    "79-01-6",
    "טריכלורואתילן":        "79-01-6",
    "pce":                  "127-18-4",
    "tetrachloroethylene":  "127-18-4",
    "perchloroethylene":    "127-18-4",
    "פרכלורואתילן":         "127-18-4",
    "vinyl chloride":       "75-01-4",
    "vc":                   "75-01-4",
    "ויניל כלוריד":         "75-01-4",
    "1,1-dichloroethylene": "75-35-4",
    "cis-1,2-dce":          "156-59-2",
    "trans-1,2-dce":        "156-60-5",
    "chloroform":           "67-66-3",
    "carbon tetrachloride": "56-23-5",
    "1,2-dichloroethane":   "107-06-2",
    "1,1,1-trichloroethane":"71-55-6",
    # Aromatics / PAHs
    "naphthalene":          "91-20-3",
    "נפטלן":                "91-20-3",
    "styrene":              "100-42-5",
    "acetone":              "67-64-1",
    "1,2,3-trimethylbenzene": "526-73-8",
    "1,2,4-trimethylbenzene": "95-63-6",
    "1,3,5-trimethylbenzene": "108-67-8",
    "isopropylbenzene":     "98-82-8",
    "cumene":               "98-82-8",
    "n-propylbenzene":      "103-65-1",
    "n-butylbenzene":       "104-51-8",
    # Metals — by symbol
    "pb":   "7439-92-1",  "lead":      "7439-92-1",  "עופרת":    "7439-92-1",
    "zn":   "7440-66-6",  "zinc":      "7440-66-6",  "אבץ":      "7440-66-6",
    "cu":   "7440-50-8",  "copper":    "7440-50-8",  "נחושת":    "7440-50-8",
    "as":   "7440-38-2",  "arsenic":   "7440-38-2",  "ארסן":     "7440-38-2",
    "cd":   "7440-43-9",  "cadmium":   "7440-43-9",  "קדמיום":   "7440-43-9",
    "cr":   "7440-47-3",  "chromium":  "7440-47-3",  "כרום":     "7440-47-3",
    "ni":   "7440-02-0",  "nickel":    "7440-02-0",  "ניקל":     "7440-02-0",
    "hg":   "7439-97-6",  "mercury":   "7439-97-6",  "כספית":    "7439-97-6",
    "al":   "7429-90-5",  "aluminium": "7429-90-5",  "אלומיניום":"7429-90-5",
    "fe":   "7439-89-6",  "iron":      "7439-89-6",  "ברזל":     "7439-89-6",
    "mn":   "7439-96-5",  "manganese": "7439-96-5",  "מנגן":     "7439-96-5",
    "ba":   "7440-39-3",  "barium":    "7440-39-3",  "בריום":    "7440-39-3",
    "be":   "7440-41-7",  "beryllium": "7440-41-7",  "בריליום":  "7440-41-7",
    "ag":   "7440-22-4",  "silver":    "7440-22-4",  "כסף":      "7440-22-4",
    "co":   "7440-48-4",  "cobalt":    "7440-48-4",  "קובלט":    "7440-48-4",
    "b":    "7440-42-8",  "boron":     "7440-42-8",  "בורון":    "7440-42-8",
    "mo":   "7439-98-7",  "molybdenum":"7439-98-7",  "מוליבדן":  "7439-98-7",
    "se":   "7782-49-2",  "selenium":  "7782-49-2",  "סלניום":   "7782-49-2",
    "sb":   "7440-36-0",  "antimony":  "7440-36-0",  "אנטימון":  "7440-36-0",
    "tl":   "7440-28-0",  "thallium":  "7440-28-0",  "תלאיום":   "7440-28-0",
    "v":    "7440-62-2",  "vanadium":  "7440-62-2",  "ונדיום":   "7440-62-2",
    "sn":   "7440-31-5",  "tin":       "7440-31-5",  "בדיל":     "7440-31-5",
}


def name_to_cas(chemical_name: str) -> str | None:
    key = chemical_name.strip().lower()
    return CHEMICAL_MAP.get(key)


def cas_to_name(cas: str, lang: str = "en") -> str | None:
    cas = cas.strip()
    for name, c in CHEMICAL_MAP.items():
        if c == cas:
            has_hebrew = any('\u0590' <= ch <= '\u05ff' for ch in name)
            if lang == "he" and has_hebrew:
                return name
            if lang == "en" and not has_hebrew:
                return name
    return None
