"""
core/drinking_water.py
----------------------
Maximum concentrations from the Israeli drinking water regulations:
Public Health Regulations (Sanitary Quality of Drinking Water and Drinking Water
Facilities), 2013 (תקנות בריאות העם, התשע"ג-2013), version of 1.7.2016.

Values are in mg/L, the unit of the groundwater sheets. The regulations give the
First Addendum in µg/L and the Second in mg/L. Rows marked "sum" are a standard for
the sum of a group: a member gets the group's limit too, since a member above it
puts the sum above it.

Not here: limits with no single-sample number (pH range, taste, odor), radioactive
materials (Bq/L), the relative-sum rule (1.5), disinfection by-products, and
"oils" (OG), which isn't the labs' Total Petroleum Hydrocarbons.
"""

import re

# The threshold key of this standard
DRINKING_WATER = "DRINKING_WATER"

_UG = 0.001  # µg/L → mg/L

# (CAS, limit in mg/L, names)
_LIMITS: list[tuple[str, float, tuple[str, ...]]] = [
    # ── First Addendum, Table A: inorganic substances (µg/L) ──
    ("7440-36-0", 6 * _UG,      ("Antimony",)),
    ("7440-61-1", 15 * _UG,     ("Uranium",)),
    ("7440-38-2", 10 * _UG,     ("Arsenic",)),
    ("7440-42-8", 1000 * _UG,   ("Boron",)),
    ("7440-39-3", 1000 * _UG,   ("Barium",)),
    ("7440-41-7", 4 * _UG,      ("Beryllium",)),
    ("14797-55-8", 70000 * _UG, ("Nitrates", "Nitrate", "Nitrate as NO3")),  # as NO3
    ("7440-22-4", 100 * _UG,    ("Silver",)),
    ("7439-97-6", 1 * _UG,      ("Mercury",)),
    ("7440-47-3", 50 * _UG,     ("Chromium",)),
    ("7439-98-7", 70 * _UG,     ("Molybdenum",)),
    ("7440-02-0", 20 * _UG,     ("Nickel",)),
    ("7782-49-2", 10 * _UG,     ("Selenium",)),
    ("7439-92-1", 10 * _UG,     ("Lead",)),
    ("16984-48-8", 1700 * _UG,  ("Fluoride",)),
    ("57-12-5", 50 * _UG,       ("Total Cyanide", "Cyanide")),
    ("7440-43-9", 5 * _UG,      ("Cadmium",)),
    ("7440-28-0", 2 * _UG,      ("Thallium",)),

    # ── First Addendum, Table B: pesticides (µg/L) ──
    ("23135-22-0", 200 * _UG,   ("Oxamyl",)),
    ("15972-60-8", 4 * _UG,     ("Alachlor",)),
    ("116-06-3", 10 * _UG,      ("Aldicarb",)),            # sum
    ("1646-88-4", 10 * _UG,     ("Aldicarb sulfone",)),    # sum
    ("1646-87-3", 10 * _UG,     ("Aldicarb sulfoxide",)),  # sum
    ("1912-24-9", 2 * _UG,      ("Atrazine",)),
    ("309-00-2", 0.05 * _UG,    ("Aldrin",)),              # sum
    ("60-57-1", 0.05 * _UG,     ("Dieldrin",)),            # sum
    ("106-93-4", 0.05 * _UG,    ("Ethylene Dibromide", "1,2-Dibromoethane (EDB)",
                                 "1,2-Dibromoethane")),
    ("94-75-7", 30 * _UG,       ("2,4-D",)),
    ("96-12-8", 0.3 * _UG,      ("1,2-Dibromo-3-chloropropane",)),
    ("107917-42-0", 1 * _UG,    ("DDT", "Sum of 4 isomers DDT", "Sum of 6 isomers DDT",
                                 "2,4-DDT", "4,4-DDT")),   # sum, with DDD and DDE
    ("72-54-8", 1 * _UG,        ("DDD", "4,4-DDD", "2,4-DDD")),  # sum
    ("72-55-9", 1 * _UG,        ("DDE", "4,4-DDE", "2,4-DDE")),  # sum
    ("78-87-5", 5 * _UG,        ("1,2-Dichloropropane",)),
    ("60-51-5", 6 * _UG,        ("Dimethoate",)),
    ("88-85-7", 7 * _UG,        ("Dinoseb",)),
    ("85-00-7", 20 * _UG,       ("Diquat",)),
    ("76-44-8", 0.4 * _UG,      ("Heptachlor",)),
    ("1024-57-3", 0.2 * _UG,    ("Heptachlor epoxide", "Heptachloroepoxide-cis",
                                 "Heptachloroepoxide-trans",
                                 "Sum of heptachloroepoxide-cis & trans")),
    ("93-76-5", 9 * _UG,        ("2,4,5-T",)),
    ("1582-09-8", 20 * _UG,     ("Trifluralin",)),
    ("57-74-9", 1 * _UG,        ("Chlordane", "alpha-Chlordane", "gamma-Chlordane",
                                 "cis-Chlordane", "trans-Chlordane", "Oxychlordane")),
    ("2921-88-2", 30 * _UG,     ("Chlorpyrifos",)),
    ("58-89-9", 1 * _UG,        ("Lindane", "gamma-HCH", "gamma-Hexachlorocyclohexane")),
    ("51218-45-2", 10 * _UG,    ("Metolachlor", "Metolachlor (isomers)")),
    ("94-74-6", 2 * _UG,        ("MCPA",)),
    ("122-34-9", 2 * _UG,       ("Simazine",)),
    ("93-72-1", 10 * _UG,       ("2,4,5-TP", "Silvex")),
    ("87-86-5", 3 * _UG,        ("Pentachlorophenol", "Pentachlorophenol (PCP)")),
    ("1563-66-2", 20 * _UG,     ("Carbofuran",)),

    # ── First Addendum, Table C: organic substances of industrial origin (µg/L) ──
    ("100-41-4", 300 * _UG,     ("Ethylbenzene", "Ethyl Benzene")),
    ("1336-36-3", 0.5 * _UG,    ("Polychlorinated Biphenyls", "PCBs", "PCB")),  # sum
    ("71-43-2", 5 * _UG,        ("Benzene",)),
    ("50-32-8", 0.5 * _UG,      ("Benzo[a]pyrene",)),
    ("117-81-7", 8 * _UG,       ("Di(2-ethylhexyl) phthalate", "DEHP")),
    ("75-35-4", 10 * _UG,       ("1,1-Dichloroethylene", "1,1-Dichloroethene")),
    ("156-59-2", 50 * _UG,      ("cis-1,2-Dichloroethylene", "cis-1,2-Dichloroethene")),
    ("156-60-5", 50 * _UG,      ("trans-1,2-Dichloroethylene", "trans-1,2-Dichloroethene")),
    ("107-06-2", 4 * _UG,       ("1,2-Dichloroethane",)),
    ("95-50-1", 600 * _UG,      ("1,2-Dichlorobenzene",)),
    ("106-46-7", 75 * _UG,      ("1,4-Dichlorobenzene",)),
    ("75-09-2", 5 * _UG,        ("Dichloromethane",)),
    ("75-01-4", 0.5 * _UG,      ("Vinyl chloride",)),
    ("108-88-3", 700 * _UG,     ("Toluene",)),             # the regulations print 105-88-3
    ("127-18-4", 10 * _UG,      ("Tetrachloroethylene", "Tetrachloroethene")),
    ("71-55-6", 200 * _UG,      ("1,1,1-Trichloroethane",)),
    ("79-00-5", 5 * _UG,        ("1,1,2-Trichloroethane",)),
    ("79-01-6", 20 * _UG,       ("Trichloroethylene", "Trichloroethene")),
    ("120-82-1", 70 * _UG,      ("1,2,4-Trichlorobenzene",)),
    ("67-66-3", 80 * _UG,       ("Chloroform",)),
    ("108-90-7", 100 * _UG,     ("Monochlorobenzene", "Chlorobenzene")),
    ("100-42-5", 50 * _UG,      ("Styrene",)),
    ("50-00-0", 900 * _UG,      ("Formaldehyde",)),        # the regulations print 50-00-00
    ("56-23-5", 4 * _UG,        ("Carbon tetrachloride", "Tetrachloromethane")),
    ("1330-20-7", 500 * _UG,    ("Xylene", "Xylenes", "Sum of xylenes", "Total Xylenes",
                                 "meta- & para-Xylene", "ortho-Xylene")),  # sum of all isomers

    # ── Second Addendum, Table A: substances with an organoleptic effect (mg/L) ──
    ("7440-66-6", 5.0,          ("Zinc",)),
    ("7429-90-5", 0.2,          ("Aluminium", "Aluminum")),
    ("7439-89-6", 1.0,          ("Iron",)),
    ("14808-79-8", 250.0,       ("Sulphate as SO4 2-", "Sulphate", "Sulfate", "Sulfate as SO4")),
    ("16887-00-6", 400.0,       ("Chloride",)),
    ("7439-96-5", 0.2,          ("Manganese",)),
    ("", 0.5,                   ("MBAS", "Anionic surfactants", "Anionic detergents")),
    ("1634-04-4", 0.04,         ("MTBE", "Methyl tert-Butyl Ether (MTBE)",
                                 "Methyl tert-Butyl Ether")),
    ("7440-50-8", 1.4,          ("Copper",)),
]


def _norm(name: str) -> str:
    """Compare names without punctuation: 1.2-Dichloroethane = 1,2-dichloroethane."""
    name = re.sub(r"\(m\d+\)\s*$", "", name.strip().lower())  # the labs' "(M1)" variant of a sum
    return re.sub(r"[^a-z0-9]", "", name)


_BY_NAME = {_norm(n): round(limit, 9) for _, limit, names in _LIMITS for n in names}
_BY_CAS = {cas: round(limit, 9) for cas, limit, _ in _LIMITS if cas}


def drinking_water_limit(cas: str = "", compound_name: str = "") -> float | None:
    """The compound's maximum in mg/L, or None when the regulations don't list it.

    The name decides first, then the CAS.
    """
    if compound_name:
        by_name = _BY_NAME.get(_norm(compound_name))
        if by_name is not None:
            return by_name
    return _BY_CAS.get(str(cas or "").strip())
