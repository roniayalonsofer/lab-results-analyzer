"""
thresholds/pfas_thresholds.py
------------------------------
PFAS threshold values (March 2026) in ng/g (= µg/kg = ng/g dry weight).
CAS numbers are keys.
"""

# Soil VSL
PFAS_VSL: dict[str, float] = {
    "355-46-4":  5.893,      # PFHxS
    "307-24-4":  2.166,      # PFHxA
    "1763-23-1": 115.031,    # PFOS
    "335-67-1":  610.424,    # PFOA
}

# Tier 1 Residential — high sensitivity (same as VSL)
PFAS_TIER1_RES_HIGH_SENSITIVITY: dict[str, float] = {
    "355-46-4":  5.893,
    "307-24-4":  2.166,
    "1763-23-1": 115.031,
    "335-67-1":  610.424,
}

# Tier 1 Residential — depth 0–6 m bgs (same as VSL)
PFAS_TIER1_RES_0_6: dict[str, float] = {
    "355-46-4":  5.893,
    "307-24-4":  2.166,
    "1763-23-1": 115.031,
    "335-67-1":  610.424,
}

# Tier 1 Residential — depth 6+ m bgs
PFAS_TIER1_RES_6PLUS: dict[str, float] = {
    "355-46-4":  15.715,
    "307-24-4":  5.776,
    "1763-23-1": 153.609,
    "335-67-1":  1627.797,
}

# Tier 1 Residential — no groundwater pathway
PFAS_TIER1_RES_NO_GW: dict[str, float] = {
    "355-46-4":  613.173,
    "307-24-4":  56892.312,
    "1763-23-1": 153.609,
    "335-67-1":  4108.889,
}

# Tier 1 Industrial — high sensitivity (same as VSL)
PFAS_TIER1_IND_HIGH_SENSITIVITY: dict[str, float] = {
    "355-46-4":  5.893,
    "307-24-4":  2.166,
    "1763-23-1": 115.031,
    "335-67-1":  610.424,
}

# Tier 1 Industrial — depth 0–6 m bgs (same as VSL)
PFAS_TIER1_IND_0_6: dict[str, float] = {
    "355-46-4":  5.893,
    "307-24-4":  2.166,
    "1763-23-1": 115.031,
    "335-67-1":  610.424,
}

# Tier 1 Industrial — depth 6+ m bgs
PFAS_TIER1_IND_6PLUS: dict[str, float] = {
    "355-46-4":  15.715,
    "307-24-4":  5.776,
    "1763-23-1": 306.750,
    "335-67-1":  1627.797,
}

# Tier 1 Industrial — no groundwater pathway
PFAS_TIER1_IND_NO_GW: dict[str, float] = {
    "355-46-4":  6965.375,
    "307-24-4":  646271.887,
    "1763-23-1": 1744.934,
    "335-67-1":  46675.192,
}
