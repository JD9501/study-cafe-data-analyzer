# Study Cafe Survey Analyzer: browser edition (runs on Pyodide), built from v57
# ============================================================
#  STUDY CAFÉ SURVEY — DATA CLEANING & RELIABILITY PIPELINE
#  Google Colab Script  |  v57 (v56 + plain-language Surface and Deep reports: key findings, section guides, glossary)
#   v56 (v55 + pricing: demand curves, elasticity, price headroom, willingness to pay by group, profit sensitivity table)
#   v55 (v54 + design upgrade: brand chart palette, consistent category colours, grouped what-if settings)
#   v54 (v53 + bundled dataset: drag the Excel icon into the drop box to start)
#   v53 (v52 + Surface / Deep view switch, rules worth acting on)
# ============================================================
#  HOW TO USE
#  1. Upload this .py to Google Colab → paste contents into
#     a single code cell (or File ▸ Upload notebook).
#  2. Run the cell.
#  3. Click the "Open Data Cleaner" link that appears.
#  4. Drag & drop your Excel survey file.
#  5. Watch the progress bar, then explore the dashboard.
# ============================================================

import sys

# ── Imports ──────────────────────────────────────────────────
import io, json, os, threading, time
import pandas as pd
import numpy as np
from flask import Flask, request, jsonify, send_file
from scipy import stats
from collections import Counter
try:
    from sklearn.linear_model import LogisticRegression, Ridge
    from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier, RandomForestRegressor
    from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor
    from sklearn.neighbors import KNeighborsClassifier, KNeighborsRegressor
    from sklearn.naive_bayes import GaussianNB
    from sklearn.svm import SVC, SVR
    from sklearn.cluster import KMeans
    from sklearn.decomposition import PCA
    from sklearn.preprocessing import StandardScaler, LabelEncoder
    from sklearn.model_selection import cross_val_score, train_test_split
    from sklearn.metrics import confusion_matrix, classification_report
    SKLEARN_OK = True
except ImportError:
    SKLEARN_OK = False

try:
    from google.colab import output as _colab_output
    IN_COLAB = True
except ImportError:
    IN_COLAB = False

app  = Flask(__name__)
# v41: never re-sort result keys alphabetically — chart order is decided by the analysis.
try:
    app.json.sort_keys = False
except Exception:
    app.config["JSON_SORT_KEYS"] = False

# ════════════════════════════════════════════════════════════
#  DOMAIN KNOWLEDGE  —  valid category sets
# ════════════════════════════════════════════════════════════
VALID = {
    "Avg_Spend_per_visit":       {"100-200","200-400","400-600",">600"},
    "Budget_Per_Visit":          {"<=200","200-400","400-600",">600"},
    "Preferred_Time":            {"Morning","Afternoon","Evening","Late Night"},
    "Noise_Tolerance":           {"Silent","Very Quiet","Conversation OK","Moderate","Lively OK"},
    "Preferred_Seating_Type":    {"Individual Desk","Group Table","Private Booth",
                                  "Sofa Lounge","Lounge/Sofa","Standing Desk"},
    "Food_Preference":           {"Snacks","Quick Meals","Desserts","Full Meals","No Food"},
    "Beverage_Preference":       {"Coffee","Tea","Juices","Smoothie","Water",
                                  "Energy Drinks","Energy Drink"},
    "Distance_from_Campus":      {"<1 km","1-2 km","2-3 km",">3 km","3-5 km",">5 km"},
    "Study_Hours":               {"<1","1-2","2-3","3-4",">4","5-6"},
    "Group_Study_Frequency":     {"Never","Rarely","Sometimes","Occasionally","Frequently","Often","Always"},
    "Visit_Frequency_per_week":  {"Daily","1","2-3","4-5","Rarely"},
    "Membership_Price_Tolerance":{"500-1000","1000-1500","1500-2000",">2000"},
    "Late_Night_Study_Preference":{"Yes","No"},
    "Prefer_Study_Cafe":         {"Yes","No"},
    "Willingness_for_Membership":{"Yes","No"},
}
LIKERT_COLS = ["Power_Outlet_Requirement","WiFi_Importance","Productivity_Tools_Interest"]
VALID_AGE   = (17, 35)
SPEND_MID   = {"100-200":150,"200-400":300,"400-600":500,">600":700}
BUDGET_MID  = {"<=200":200,"200-400":400,"400-600":600,">600":800}

# ════════════════════════════════════════════════════════════
#  v41 — GENERIC LABEL / ENCODING HELPERS
#  Every chart order, numeric encoding and category list is now
#  derived from the data that was uploaded. The dictionaries below
#  are only *preferred* values for labels the tool already knows;
#  any label it has never seen is parsed or ranked automatically,
#  so new category labels can never be silently dropped.
# ════════════════════════════════════════════════════════════
import re as _re, math as _math

COL_ORDER = {
    "Visit_Frequency_per_week":   ["Rarely","1","2-3","4-5","Daily"],
    "Group_Study_Frequency":      ["Never","Rarely","Sometimes","Occasionally","Frequently","Often","Always"],
    "Preferred_Time":             ["Morning","Afternoon","Evening","Late Night"],
    "Study_Hours":                ["<1","1-2","2-3","3-4",">4","5-6"],
    "Distance_from_Campus":       ["<1 km","1-2 km","2-3 km",">3 km","3-5 km",">5 km","5+ km"],
    "Avg_Spend_per_visit":        ["100-200","200-400","400-600",">600"],
    "Budget_Per_Visit":           ["<=200","200-400","400-600",">600"],
    "Membership_Price_Tolerance": ["500-1000","1000-1500","1500-2000",">2000"],
    "Willingness_for_Membership": ["Yes","No"],
    "Late_Night_Study_Preference":["Yes","No"],
    "Noise_Tolerance":            ["Silent","Very Quiet","Moderate","Conversation OK","Lively OK"],
}
ORD_MAPS = {
    "Avg_Spend_per_visit":      {"100-200":1,"200-400":2,"400-600":3,">600":4},
    "Budget_Per_Visit":         {"<=200":1,"200-400":2,"400-600":3,">600":4},
    "Visit_Frequency_per_week": {"Rarely":1,"1":2,"2-3":3,"4-5":4,"Daily":5},
    "Study_Hours":              {"<1":1,"1-2":2,"2-3":3,"3-4":4,">4":5,"5-6":5},
    "Distance_from_Campus":     {"<1 km":1,"1-2 km":2,"2-3 km":3,">3 km":4,"3-5 km":4,">5 km":5,"5+ km":5},
    "Group_Study_Frequency":    {"Never":1,"Rarely":2,"Sometimes":3,"Occasionally":3,"Frequently":4,"Often":4,"Always":5},
}
NUM_MAPS = {
    "Visit_Frequency_per_week":   {"Rarely":0.5,"1":1,"2-3":2.5,"4-5":4.5,"Daily":7},
    "Avg_Spend_per_visit":        SPEND_MID,
    "Budget_Per_Visit":           BUDGET_MID,
    "Study_Hours":                {"<1":0.5,"1-2":1.5,"2-3":2.5,"3-4":3.5,">4":4.5,"5-6":5.5},
    "Distance_from_Campus":       {"<1 km":0.5,"1-2 km":1.5,"2-3 km":2.5,">3 km":4,"3-5 km":4,">5 km":6,"5+ km":6},
    "Group_Study_Frequency":      {"Never":0,"Rarely":1,"Sometimes":2,"Occasionally":3,"Frequently":4,"Often":5,"Always":6},
    "Preferred_Time":             {"Morning":0,"Afternoon":1,"Evening":2,"Late Night":3},
    "Membership_Price_Tolerance": {"500-1000":750,"1000-1500":1250,"1500-2000":1750,">2000":2250},
}

def _label_nums(label):
    return [float(x) for x in _re.findall(r"\d+(?:\.\d+)?", str(label))]

def _is_upper_open(s):
    s = str(s).strip()
    return s.startswith((">", "≥")) or s.endswith("+")

def _is_lower_open(s):
    return str(s).strip().startswith(("<", "≤"))

def _label_sort_key(label):
    s = str(label).strip(); nums = _label_nums(s)
    if not nums:
        return (1, 0.0, 0, s.lower())
    bump = -1 if _is_lower_open(s) else (1 if _is_upper_open(s) else 0)
    return (0, nums[0], bump, s.lower())

def ordered_cats(series, col=None):
    """Category labels actually present in the data, in a sensible order:
    1) the known natural order for that question, 2) numeric order parsed
    from range labels such as '200-400' or '>600', 3) most frequent first."""
    s = series.dropna().astype(str).str.strip()
    present = list(s.value_counts().index)
    known = COL_ORDER.get(col, [])
    if present and all(p in known for p in present):
        return [k for k in known if k in present]
    if present and all(_label_nums(p) for p in present):
        return sorted(present, key=_label_sort_key)
    if known and any(p in known for p in present):
        return [k for k in known if k in present] + [p for p in present if p not in known]
    return present

def _parse_mid(label):
    """Numeric midpoint of a range label ('200-400' -> 300, '<1' -> 0.5, '>600' -> 690)."""
    nums = _label_nums(label)
    if not nums: return None
    if len(nums) >= 2: return (nums[0] + nums[1]) / 2.0
    if _is_lower_open(label): return nums[0] / 2.0
    if _is_upper_open(label): return nums[0] * 1.15
    return nums[0]

def num_series(df, col):
    """Numeric value for an ordinal/range question. Known labels use the documented
    midpoints; unseen labels are parsed; purely verbal unseen labels use their rank."""
    known = NUM_MAPS.get(col, {})
    s = df[col].astype(object).where(df[col].notna(), None)
    cats = ordered_cats(df[col], col)
    mapping = {}
    for i, c in enumerate(cats):
        if c in known: mapping[c] = float(known[c])
        else:
            m = _parse_mid(c)
            mapping[c] = float(m) if m is not None else float(i)
    return s.map(lambda v: mapping.get(str(v).strip()) if v is not None else np.nan).astype(float)

def ord_series(df, col):
    """Ordinal level (1..k). Known labels keep their documented level; if any
    unseen label is present the levels are the rank in ordered_cats()."""
    known = ORD_MAPS.get(col, {})
    cats = ordered_cats(df[col], col)
    if cats and all(c in known for c in cats):
        mapping = {c: float(known[c]) for c in cats}
    else:
        mapping = {c: float(i + 1) for i, c in enumerate(cats)}
    s = df[col].astype(object).where(df[col].notna(), None)
    return s.map(lambda v: mapping.get(str(v).strip()) if v is not None else np.nan).astype(float)

def yes_series(series):
    return series.astype(str).str.strip().str.lower().eq("yes").astype(int)

def _budget_band(label, prev_width=None):
    """(low, high) willingness-to-pay band implied by a budget label."""
    nums = _label_nums(label)
    if not nums: return None
    if len(nums) >= 2: return (min(nums[:2]), max(nums[:2]))
    if _is_lower_open(label): return (0.0, nums[0])
    if _is_upper_open(label):
        w = prev_width if prev_width else nums[0] * 0.67
        return (nums[0], nums[0] + 2 * w)
    return (nums[0], nums[0])

def budget_bands(df, col="Budget_Per_Visit"):
    cats = ordered_cats(df[col], col); out = {}; prev_w = None
    for c in cats:
        b = _budget_band(c, prev_w)
        if b is None: continue
        out[c] = b
        if b[1] > b[0] and not _is_upper_open(c): prev_w = b[1] - b[0]
    return out

def majority_share(y):
    vc = pd.Series(y).value_counts(normalize=True)
    return float(vc.iloc[0]) if len(vc) else 0.0

def _clamp(v, lo, hi):
    try: v = float(v)
    except Exception: v = lo
    if _math.isnan(v): v = lo
    return max(lo, min(hi, v))

FEATURE_LABELS = {
    "Age":"Age","_spend":"Spend","_budget":"Budget","_sh":"Study Hours","_vf":"Visit Frequency",
    "WiFi_Importance":"WiFi Need","Power_Outlet_Requirement":"Power Need",
    "Productivity_Tools_Interest":"Tools Interest",
}

def age_bucket_edges(ages):
    a = pd.Series(ages).dropna().astype(int)
    if a.empty: return [0], 1
    lo = int(a.min()); hi = int(np.percentile(a, 97))
    width = 2 if (hi - lo) <= 14 else (5 if (hi - lo) <= 40 else 10)
    edges = list(range(lo, max(hi, lo) + 1, width))
    return edges, width

def age_bucket_labels(edges, width):
    labs = [f"{e}-{e+width-1}" if width > 1 else f"{e}" for e in edges[:-1]]
    return labs + [f"{edges[-1]}+"]

def age_bucket_index(age, edges, width):
    i = int((int(age) - edges[0]) // width)
    return max(0, min(len(edges) - 1, i))

def age_valid_range(ages):
    a = pd.Series(ages).dropna().astype(float)
    if a.empty or a.between(*VALID_AGE).mean() >= 0.90:
        return VALID_AGE
    q1, q3 = a.quantile(0.25), a.quantile(0.75); iqr = max(float(q3 - q1), 1.0)
    return (max(10, int(_math.floor(q1 - 2.5 * iqr))), min(100, int(_math.ceil(q3 + 2.5 * iqr))))

def spend_and_ceiling(df):
    """Per-row spend midpoint and budget ceiling (top of the stated budget band)."""
    sp = num_series(df, "Avg_Spend_per_visit")
    bands = budget_bands(df)
    lab = df["Budget_Per_Visit"].astype(object).where(df["Budget_Per_Visit"].notna(), None)
    bu = lab.map(lambda v: (BUDGET_MID.get(str(v).strip(), bands.get(str(v).strip(), (np.nan, np.nan))[1])
                            if v is not None else np.nan)).astype(float)
    return sp, bu

def _mode(series):
    c = Counter(series.dropna())
    return c.most_common(1)[0][0] if c else None

# ════════════════════════════════════════════════════════════
#  RELIABILITY TESTS
# ════════════════════════════════════════════════════════════

def cronbach_alpha(df_items):
    """Cronbach's α for internal consistency of Likert items."""
    k   = df_items.shape[1]
    iv  = df_items.var(axis=0, ddof=1).sum()
    tv  = df_items.sum(axis=1).var(ddof=1)
    if tv == 0:
        return 0.0
    return round((k/(k-1)) * (1 - iv/tv), 4)

def interpret_alpha(a):
    if a >= 0.9:  return "Excellent"
    if a >= 0.8:  return "Good"
    if a >= 0.7:  return "Acceptable"
    if a >= 0.6:  return "Questionable"
    if a >= 0.5:  return "Poor"
    return "Unacceptable"

def run_reliability_tests(df: pd.DataFrame) -> dict:
    results = {}

    # 1. Cronbach's Alpha
    lik = df[LIKERT_COLS].dropna().astype(float)
    alpha = cronbach_alpha(lik) if len(lik) > 5 else None
    results["cronbach_alpha"] = {
        "value": alpha,
        "n":     int(len(lik)),
        "interpretation": interpret_alpha(alpha) if alpha is not None else "N/A",
        "description": (
            "Measures internal consistency of the three Likert-scale items "
            "(Power Outlet Requirement, WiFi Importance, Productivity Tools Interest). "
            "Values ≥ 0.7 indicate acceptable reliability."
        )
    }

    # 2. Inter-item correlations
    corr = lik.corr()
    _c = corr.to_numpy(dtype=float, copy=True)   # v41: writable copy (pandas 3 arrays are read-only)
    np.fill_diagonal(_c, np.nan)
    avg_corr = float(np.nanmean(_c)) if np.isfinite(_c).any() else 0.0
    corr_dict = {}
    for c1 in corr.columns:
        for c2 in corr.columns:
            if c1 < c2:
                corr_dict[f"{c1} ↔ {c2}"] = round(float(corr.loc[c1,c2]), 4)
    results["inter_item_correlation"] = {
        "average": round(avg_corr, 4),
        "pairs":   corr_dict,
        "description": (
            "Average pairwise Pearson correlation among Likert items. "
            "Values between 0.15–0.50 suggest items measure related but distinct constructs."
        )
    }

    # 3. Normality test (Age — Shapiro-Wilk on clean age)
    age_clean = df["Age"].dropna()
    _ar = age_valid_range(age_clean)
    age_clean = age_clean[(age_clean >= _ar[0]) & (age_clean <= _ar[1])]
    sample    = age_clean.sample(min(500, len(age_clean)), random_state=42) if len(age_clean) > 5 else age_clean
    if len(sample) >= 3:
        sw_stat, sw_p = stats.shapiro(sample)
        normal = sw_p > 0.05
    else:
        sw_stat, sw_p, normal = None, None, None
    results["normality_age"] = {
        "shapiro_stat": round(float(sw_stat), 4) if sw_stat else None,
        "p_value":      round(float(sw_p),    4) if sw_p    else None,
        "is_normal":    bool(normal) if normal is not None else None,
        "n":            int(len(sample)),
        "description": (
            "Shapiro-Wilk test on respondent Age (within 17–35). "
            "p > 0.05 suggests the age distribution does not significantly deviate from normal."
        )
    }

    # 4. Response completeness rate
    total_cells  = df.shape[0] * df.shape[1]
    missing_cells = int(df.isnull().sum().sum())
    completeness  = round((1 - missing_cells/total_cells)*100, 2)
    results["completeness"] = {
        "total_cells":    total_cells,
        "missing_cells":  missing_cells,
        "completeness_pct": completeness,
        "complete_rows":  int(df.dropna().shape[0]),
        "total_rows":     int(df.shape[0]),
        "description": (
            "Percentage of non-missing values across the entire dataset. "
            "Higher is better; ≥ 85% is generally considered acceptable."
        )
    }

    # 5. Straight-line / acquiescence bias detection
    pl = [c for c in LIKERT_COLS if c in df.columns]
    if len(pl) >= 2:
        filled = df[pl].dropna()
        sl_count  = int((filled.nunique(axis=1) == 1).sum())
        acq_count = int((filled.mean(axis=1) >= 4.5).sum())   # always high
        dis_count = int((filled.mean(axis=1) <= 1.5).sum())   # always low
    else:
        sl_count = acq_count = dis_count = 0
    results["response_bias"] = {
        "straight_line_count": sl_count,
        "acquiescence_bias_count": acq_count,
        "disacquiescence_bias_count": dis_count,
        "description": (
            "Straight-lining = all Likert scores identical. "
            "Acquiescence bias = respondent always chose 4–5. "
            "Disacquiescence = respondent always chose 1–2. "
            "High counts reduce data quality."
        )
    }

    # 6. Late-night contradiction
    ln_contra = int(len(df[
        (df["Late_Night_Study_Preference"] == "No") &
        (df["Preferred_Time"] == "Late Night")
    ]))
    results["late_night_contradiction"] = {
        "count": ln_contra,
        "description": (
            "Respondents who said No to Late_Night_Study_Preference but "
            "listed 'Late Night' as their Preferred_Time — a logical contradiction."
        )
    }

    # 7. Spend > Budget contradiction
    df2 = df.copy()
    df2["_sp"], df2["_bu"] = spend_and_ceiling(df2)
    sp_contra = int(len(df2[df2["_sp"].notna() & df2["_bu"].notna() & (df2["_sp"] > df2["_bu"])]))
    results["spend_budget_contradiction"] = {
        "count": sp_contra,
        "description": (
            "Rows where Avg_Spend_per_visit midpoint exceeds Budget_Per_Visit midpoint — "
            "the respondent claims to spend more than their stated budget allows."
        )
    }

    # 8. Missing pattern analysis — columns with > 30% missing
    miss_pct = (df.isnull().sum() / len(df) * 100).round(1)
    high_miss = {col: float(pct) for col, pct in miss_pct.items() if pct > 30}
    results["high_missing_columns"] = {
        "columns": high_miss,
        "description": (
            "Columns where more than 30% of values are missing. "
            "These may require extra caution during analysis."
        )
    }

    return results


# ════════════════════════════════════════════════════════════
#  DATA CLEANING  —  all techniques
# ════════════════════════════════════════════════════════════

def clean_survey(df_raw: pd.DataFrame) -> tuple:
    log = {}
    df  = df_raw.copy()
    n_start = len(df)

    # ── 1. Duplicate Removal ─────────────────────────────────
    b = len(df)
    df = df.drop_duplicates(subset=["Gmail_Account"], keep="first")
    df = df.drop_duplicates(subset=["Student_ID"],    keep="first")
    df = df.reset_index(drop=True)
    log["01_duplicates_removed"] = {
        "label": "Duplicate Responses Removed",
        "count": b - len(df),
        "method": "Exact match on Gmail_Account then Student_ID; kept first occurrence."
    }

    # ── 2. Outlier Detection (Age + Likert) ──────────────────
    out_idx = set()
    _age_rng = age_valid_range(df["Age"])
    age_mask = df["Age"].notna() & ~df["Age"].between(*_age_rng)
    out_idx.update(df[age_mask].index.tolist())
    for col in LIKERT_COLS:
        if col in df.columns:
            bad = df[col].notna() & ~df[col].between(1, 5)
            out_idx.update(df[bad].index.tolist())
    df = df.drop(index=list(out_idx)).reset_index(drop=True)
    log["02_outliers_removed"] = {
        "label": "Statistical Outliers Removed",
        "count": len(out_idx),
        "method": f"Age outside [{_age_rng[0]}–{_age_rng[1]}]; Likert values outside [1–5] removed entirely."
    }

    # ── 3. Range / Category Validation ──────────────────────
    # v41: a label outside the known list is treated as a data-entry error only
    # when it is rare (< 0.5 % of rows and fewer than 5 rows would be too strict
    # for tiny files, so the floor is 5). A label that many respondents chose is
    # a genuine category of THIS dataset and is kept and analysed.
    inv_idx = set(); accepted_new = {}
    rare_cut = max(5, int(0.005 * len(df)))
    for col, vset in VALID.items():
        if col in df.columns:
            unknown = df.loc[df[col].notna() & ~df[col].isin(vset), col]
            for label, cnt in unknown.value_counts().items():
                if cnt < rare_cut:
                    inv_idx.update(unknown[unknown == label].index.tolist())
                else:
                    accepted_new.setdefault(col, []).append(str(label))
    df = df.drop(index=list(inv_idx)).reset_index(drop=True)
    log["03_range_invalid_removed"] = {
        "label": "Out-of-Range Category Values Removed",
        "count": len(inv_idx),
        "method": "Each categorical column validated against its allowed value set; "
                  "rows with rare unrecognised values (e.g. ₹5000, ₹10000 in Avg_Spend) removed."
                  + ((" New categories kept: " + "; ".join(f"{c}: {', '.join(v)}" for c, v in accepted_new.items()) + ".")
                     if accepted_new else "")
    }

    # ── 4. Straight-Line Response Removal ───────────────────
    pl = [c for c in LIKERT_COLS if c in df.columns]
    if len(pl) >= 2:
        filled    = df[pl].notna().all(axis=1)
        sl_mask   = df[pl].nunique(axis=1) == 1
        sl_idx    = df[sl_mask & filled].index.tolist()
        df        = df.drop(index=sl_idx).reset_index(drop=True)
        sl_count  = len(sl_idx)
    else:
        sl_count = 0
    log["04_straight_line_removed"] = {
        "label": "Straight-Line Responses Removed",
        "count": sl_count,
        "method": "Rows where all three Likert items carry the identical value flagged and removed."
    }

    # ── 5. Logical Contradiction: Late Night ─────────────────
    ln_idx = df[
        (df["Late_Night_Study_Preference"] == "No") &
        (df["Preferred_Time"] == "Late Night")
    ].index.tolist()
    df = df.drop(index=ln_idx).reset_index(drop=True)
    log["05_late_night_contradiction_removed"] = {
        "label": "Late-Night Logical Contradictions Removed",
        "count": len(ln_idx),
        "method": "Rows stating No to Late_Night_Study_Preference but choosing "
                  "'Late Night' as Preferred_Time are internally inconsistent."
    }

    # ── 6. Logical Contradiction: Spend > Budget ─────────────
    df["_sp"], df["_bu"] = spend_and_ceiling(df)
    sb_idx = df[df["_sp"].notna() & df["_bu"].notna() & (df["_sp"] > df["_bu"])].index.tolist()
    df = df.drop(index=sb_idx).reset_index(drop=True)
    df = df.drop(columns=["_sp","_bu"], errors="ignore")
    log["06_spend_budget_contradiction_removed"] = {
        "label": "Spend > Budget Contradictions Removed",
        "count": len(sb_idx),
        "method": "Rows where midpoint of Avg_Spend_per_visit exceeds midpoint of "
                  "Budget_Per_Visit — the stated spend exceeds stated budget."
    }

    # ── 7. Excessive Missing Value Removal ──────────────────
    thresh  = int(len(df.columns) * 0.60)
    b       = len(df)
    df      = df.dropna(thresh=thresh).reset_index(drop=True)
    log["07_excessive_missing_removed"] = {
        "label": "Rows with Excessive Missing Values Removed",
        "count": b - len(df),
        "method": "Rows missing more than 40% of fields removed entirely."
    }

    # ── 8. Mode Imputation ──────────────────────────────────
    filled_count = 0
    for col in df.columns:
        if df[col].isna().any():
            fv = _mode(df[col])
            if fv is not None:
                filled_count += int(df[col].isna().sum())
                df[col] = df[col].fillna(fv)
    log["08_missing_filled_by_mode"] = {
        "label": "Missing Values Filled via Mode Imputation",
        "count": filled_count,
        "method": "Remaining NaN cells filled with the most frequent (mode) value of that column."
    }

    # ── 9. Filter: Target Customers Only ────────────────────
    b  = len(df)
    df = df[df["Prefer_Study_Cafe"].str.strip().str.title() == "Yes"].reset_index(drop=True)
    log["09_non_target_removed"] = {
        "label": "Non-Target Respondents Removed",
        "count": b - len(df),
        "method": "Retained only rows where Prefer_Study_Cafe = 'Yes' — "
                  "the core potential customer base."
    }

    # ── 10. Membership Consistency Fix ───────────────────────
    fixes = 0
    if "Membership_Price_Tolerance" in df.columns and "Willingness_for_Membership" in df.columns:
        not_member = df["Willingness_for_Membership"].str.strip().str.title() != "Yes"
        bad_price  = df["Membership_Price_Tolerance"].notna()
        mask       = not_member & bad_price
        df.loc[mask, "Membership_Price_Tolerance"] = np.nan
        fixes = int(mask.sum())
        # Re-impute the now-blank cells for Yes-members with no price
        yes_member  = df["Willingness_for_Membership"].str.strip().str.title() == "Yes"
        no_price    = df["Membership_Price_Tolerance"].isna()
        fill_mp     = _mode(df.loc[yes_member, "Membership_Price_Tolerance"])
        if fill_mp and (yes_member & no_price).any():
            df.loc[yes_member & no_price, "Membership_Price_Tolerance"] = fill_mp
    log["10_membership_consistency_fixed"] = {
        "label": "Membership Consistency Corrections",
        "count": fixes,
        "method": "Membership_Price_Tolerance nulled for respondents who said No to "
                  "Willingness_for_Membership; remaining blanks filled with mode price."
    }

    # ── Summary totals ───────────────────────────────────────
    total_removed = n_start - len(df)
    return df, log, n_start, total_removed


# ════════════════════════════════════════════════════════════
#  IN-MEMORY STATE
# ════════════════════════════════════════════════════════════
_store: dict = {}

def _sanitise(obj):
    """Recursively replace NaN/Inf with None so jsonify produces valid JSON."""
    import math
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, dict):
        return {k: _sanitise(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitise(v) for v in obj]
    return obj


# ════════════════════════════════════════════════════════════
#  FLASK ROUTES
# ════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return HTML_PAGE


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Only .xlsx / .xls files are supported"}), 400
    raw = f.read()
    try:
        df = pd.read_excel(io.BytesIO(raw))
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    # ── Column validation ─────────────────────────────────────────────
    # These are the exact 22 columns required by the Study Café survey pipeline.
    REQUIRED_COLUMNS = [
        "Gmail_Account", "Student_ID", "University", "Age",
        "Distance_from_Campus", "Study_Hours", "Prefer_Study_Cafe",
        "Visit_Frequency_per_week", "Preferred_Time",
        "Late_Night_Study_Preference", "Group_Study_Frequency",
        "Noise_Tolerance", "Preferred_Seating_Type",
        "Power_Outlet_Requirement", "WiFi_Importance",
        "Food_Preference", "Beverage_Preference", "Avg_Spend_per_visit",
        "Budget_Per_Visit", "Productivity_Tools_Interest",
        "Willingness_for_Membership", "Membership_Price_Tolerance",
    ]
    uploaded_cols  = list(df.columns)
    required_set   = set(REQUIRED_COLUMNS)
    uploaded_set   = set(uploaded_cols)
    missing_cols   = sorted(required_set - uploaded_set)
    extra_cols     = sorted(uploaded_set - required_set)

    if missing_cols or len(uploaded_cols) != len(REQUIRED_COLUMNS):
        parts = [
            f"Wrong dataset uploaded ({len(uploaded_cols)} columns found, "
            f"{len(REQUIRED_COLUMNS)} expected)."
        ]
        if missing_cols:
            parts.append(f"Missing columns: {', '.join(missing_cols)}.")
        if extra_cols:
            parts.append(f"Unexpected columns: {', '.join(extra_cols)}.")
        parts.append(
            "Please check that you are uploading the Study Café survey dataset."
        )
        return jsonify({"error": " ".join(parts)}), 400
    # ── End column validation ─────────────────────────────────────────

    _store["raw_bytes"] = raw
    _store["raw_shape"] = list(df.shape)
    return jsonify({
        "shape":   list(df.shape),
        "columns": list(df.columns),
        "preview": df.head(5).fillna("").to_dict(orient="records")
    })


@app.route("/clean", methods=["POST"])
def clean():
    if "raw_bytes" not in _store:
        return jsonify({"error": "Upload a file first"}), 400
    df_raw = pd.read_excel(io.BytesIO(_store["raw_bytes"]))

    # Run reliability BEFORE cleaning (on raw data)
    reliability = run_reliability_tests(df_raw)

    # Run cleaning
    df_clean, clean_log, n_start, total_removed = clean_survey(df_raw)

    # Run reliability AFTER cleaning (on target subset)
    reliability_post = run_reliability_tests(df_clean)

    # Column stats for summary
    col_stats = {}
    for col in df_clean.columns:
        s = df_clean[col]
        stat = {
            "missing":     int(s.isna().sum()),
            "unique":      int(s.nunique()),
            "dtype":       str(s.dtype),
        }
        if pd.api.types.is_numeric_dtype(s):
            stat["mean"] = round(float(s.mean()), 2) if not s.isna().all() else None
            stat["min"]  = round(float(s.min()),  2) if not s.isna().all() else None
            stat["max"]  = round(float(s.max()),  2) if not s.isna().all() else None
        else:
            top = s.value_counts().head(3).to_dict()
            stat["top_values"] = {str(k): int(v) for k, v in top.items()}
        col_stats[col] = stat

    # Save cleaned file
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_clean.to_excel(w, index=False, sheet_name="Cleaned_Data")
        summary_rows = [
            {"Step": v["label"], "Rows Affected": v["count"], "Method": v["method"]}
            for v in clean_log.values()
        ]
        pd.DataFrame(summary_rows).to_excel(w, index=False, sheet_name="Cleaning_Summary")
    _store["clean_bytes"]      = buf.getvalue()
    _store["clean_log"]        = clean_log
    _store["reliability_pre"]  = reliability
    _store["reliability_post"] = reliability_post
    _store["n_start"]          = n_start
    _store["total_removed"]    = total_removed
    _store["n_final"]          = len(df_clean)
    _store["col_stats"]        = col_stats
    _store["clean_columns"]    = list(df_clean.columns)
    _store["clean_preview"]    = df_clean.head(5).fillna("").to_dict(orient="records")
    _store.pop("analytics_cache", None)
    _store.pop("predictive_cache", None)
    _store.pop("predictive_cache_map", None)
    _store.pop("lever_progress", None)
    _store["lever_gen"] = None
    _store.pop("mem_model", None)

    return jsonify({"ok": True})


@app.route("/summary_data")
def summary_data():
    if "clean_log" not in _store:
        return jsonify({"error": "Run cleaning first"}), 400
    return jsonify({
        "clean_log":        _store["clean_log"],
        "reliability_pre":  _store["reliability_pre"],
        "reliability_post": _store["reliability_post"],
        "n_start":          _store["n_start"],
        "total_removed":    _store["total_removed"],
        "n_final":          _store["n_final"],
        "col_stats":        _store["col_stats"],
        "clean_columns":    _store["clean_columns"],
        "clean_preview":    _store["clean_preview"],
        "raw_shape":        _store.get("raw_shape"),
    })


# ── v56: pricing analysis from what students said they would pay ─────────
def _wilson(k, n, z=1.96):
    """95% range for a share (Wilson interval; stays inside 0-100%)."""
    if n <= 0: return (0.0, 0.0)
    p = k / n; d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * np.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return (max(0.0, float(c - h)), min(1.0, float(c + h)))

def _arc_elasticity(p1, q1, p2, q2):
    """Midpoint (arc) price elasticity between two points of a demand curve."""
    if p2 == p1 or (q1 + q2) <= 0 or (p1 + p2) <= 0: return None
    return float(((q2 - q1) / ((q1 + q2) / 2)) / ((p2 - p1) / ((p1 + p2) / 2)))

def _wtp_curve(series, col, rng):
    """Stated demand curve. Each answer is the most a student would pay; everyone
    whose band starts at or above a price is counted as still willing at that price."""
    s = series.dropna().astype(str).str.strip()
    s = s[s.ne("") & s.str.lower().ne("nan")]
    n = int(len(s))
    if n < 30: return None
    cats = ordered_cats(s, col)
    bands, prev_w = {}, None
    for c in cats:
        b = _budget_band(c, prev_w)
        if b is None: return None
        bands[c] = b
        if b[1] > b[0] and not _is_upper_open(c): prev_w = b[1] - b[0]
    counts = np.array([int((s == c).sum()) for c in cats], dtype=float)
    tail = counts[::-1].cumsum()[::-1]
    idx, pts = [], []
    for i, c in enumerate(cats):
        price = float(bands[c][0])
        if price <= 0: continue
        k = float(tail[i]); lo, hi = _wilson(k, n)
        idx.append(i)
        pts.append({"price": price, "band": c, "share": k / n, "lo": lo, "hi": hi,
                    "students": int(k), "rev100": price * k / n * 100})
    if len(pts) < 2: return None
    sims = rng.multinomial(n, counts / counts.sum(), size=1000).astype(float)
    tails = sims[:, ::-1].cumsum(axis=1)[:, ::-1] / n
    el = []
    for a in range(len(pts) - 1):
        b = a + 1
        e = _arc_elasticity(pts[a]["price"], pts[a]["share"], pts[b]["price"], pts[b]["share"])
        boot = [_arc_elasticity(pts[a]["price"], t[idx[a]], pts[b]["price"], t[idx[b]]) for t in tails]
        boot = [x for x in boot if x is not None and np.isfinite(x)]
        el.append({"from": pts[a]["price"], "to": pts[b]["price"], "e": e,
                   "lo": float(np.percentile(boot, 2.5)) if boot else None,
                   "hi": float(np.percentile(boot, 97.5)) if boot else None})
    best = max(pts, key=lambda r: r["rev100"])
    return {"n": n, "points": pts, "elasticity": el, "best_price": best["price"],
            "open_top": bool(_is_upper_open(cats[-1]))}

def _headroom(df):
    """How far each student's per-visit budget sits above what they spend now."""
    sc_, bc_ = "Avg_Spend_per_visit", "Budget_Per_Visit"
    if sc_ not in df.columns or bc_ not in df.columns: return None
    d = df[[sc_, bc_]].dropna().astype(str).apply(lambda c: c.str.strip())
    if len(d) < 30: return None
    s_cats, b_cats = ordered_cats(d[sc_], sc_), ordered_cats(d[bc_], bc_)
    sb, bb = budget_bands(d, sc_), budget_bands(d, bc_)
    if not sb or not bb: return None
    def top(bands, c): return bands[c][1] if c in bands else np.nan
    widths = [sb[c][1] - sb[c][0] for c in s_cats if c in sb and not _is_upper_open(c) and sb[c][1] > sb[c][0]]
    step = float(np.median(widths)) if widths else 200.0
    rows, tot = [], np.zeros(4)
    for c in s_cats:
        sub = d[d[sc_] == c]
        both_top = np.array([_is_upper_open(c) and _is_upper_open(b) for b in sub[bc_]])
        gap = np.array([top(bb, b) for b in sub[bc_]]) - top(sb, c)
        k = np.array([((gap <= step * 0.25) & ~both_top).sum(), ((gap > step * 0.25) & (gap <= step * 1.25) & ~both_top).sum(),
                      ((gap > step * 1.25) & ~both_top).sum(), both_top.sum()], dtype=float)
        tot += k
        rows.append({"spend": c, "n": int(len(sub)), "counts": [int(x) for x in k]})
    known = float(tot[:3].sum())
    return {"rows": rows, "n": int(tot.sum()), "known": int(known), "unknown": int(tot[3]),
            "shares": [float(x / known) for x in tot[:3]] if known else [0, 0, 0],
            "step": step, "budget_bands": b_cats}

_SEG_COLS = [("Visit_Frequency_per_week", "How often they would visit"),
             ("Avg_Spend_per_visit", "What they spend per visit now"),
             ("Distance_from_Campus", "Distance from campus"),
             ("Preferred_Seating_Type", "Preferred seat")]

def _wtp_segments(df, mem_mask, curve):
    """Share of interested students who would pay the 'middle' fee or more, by group."""
    col = "Membership_Price_Tolerance"
    if curve is None or col not in df.columns: return None
    cand = [p for p in curve["points"][1:] if p["share"] < 0.999] or curve["points"]
    thr = min(cand, key=lambda p: abs(p["share"] - 0.5))
    cats = ordered_cats(df.loc[mem_mask, col], col)
    high = set(cats[cats.index(thr["band"]):]) if thr["band"] in cats else set()
    base = df.loc[mem_mask].copy()
    base = base[base[col].notna()]
    base["_high"] = base[col].astype(str).str.strip().isin(high).astype(int)
    groups = []
    for seg, title in _SEG_COLS:
        if seg not in base.columns: continue
        sub = base[base[seg].notna()]
        levels = [l for l in ordered_cats(sub[seg], seg)]
        rows, table = [], []
        for l in levels:
            g = sub[sub[seg].astype(str).str.strip() == l]
            n_, k_ = int(len(g)), int(g["_high"].sum())
            if n_ < 30: continue
            lo, hi = _wilson(k_, n_)
            rows.append({"label": l, "n": n_, "share": k_ / n_, "lo": lo, "hi": hi})
            table.append([k_, n_ - k_])
        p = None
        if len(table) >= 2:
            try:
                from scipy.stats import chi2_contingency
                p = float(chi2_contingency(np.array(table))[1])
            except Exception:
                p = None
        if rows: groups.append({"key": seg, "title": title, "rows": rows, "p": p})
    return {"threshold": thr["price"], "threshold_band": thr["band"], "overall": float(base["_high"].mean()) if len(base) else 0.0,
            "n": int(len(base)), "groups": groups}

def _pricing_analysis(df):
    rng = np.random.default_rng(7)
    out = {}
    if "Membership_Price_Tolerance" in df.columns:
        mask = (yes_series(df["Willingness_for_Membership"]) == 1) if "Willingness_for_Membership" in df.columns \
               else pd.Series(True, index=df.index)
        out["membership"] = _wtp_curve(df.loc[mask, "Membership_Price_Tolerance"], "Membership_Price_Tolerance", rng)
        out["segments"] = _wtp_segments(df, mask, out["membership"])
    if "Budget_Per_Visit" in df.columns:
        out["visit"] = _wtp_curve(df["Budget_Per_Visit"], "Budget_Per_Visit", rng)
    out["headroom"] = _headroom(df)
    return out

@app.route("/analytics_data")
def analytics_data():
    import warnings; warnings.filterwarnings("ignore")
    if "clean_bytes" not in _store:
        return jsonify({"error": "Run cleaning first"}), 400
    if "analytics_cache" in _store:
        return jsonify(_store["analytics_cache"])
    try:
        df = pd.read_excel(io.BytesIO(_store["clean_bytes"]), sheet_name="Cleaned_Data")
        N  = len(df)

        # ── Ordinal encodings (data-driven; see ord_series) ────────────────
        df["_spend"]  = ord_series(df, "Avg_Spend_per_visit")
        df["_budget"] = ord_series(df, "Budget_Per_Visit")
        df["_vf"]     = ord_series(df, "Visit_Frequency_per_week")
        df["_sh"]     = ord_series(df, "Study_Hours")
        df["_dist"]   = ord_series(df, "Distance_from_Campus")
        df["_gs"]     = ord_series(df, "Group_Study_Frequency")
        df["_ln"]     = yes_series(df["Late_Night_Study_Preference"])
        df["_mem"]    = yes_series(df["Willingness_for_Membership"])

        # ── Value counts — every category present, in natural order ───────
        CAT_COLS = [
            "University","Distance_from_Campus","Study_Hours",
            "Visit_Frequency_per_week","Preferred_Time","Group_Study_Frequency",
            "Noise_Tolerance","Preferred_Seating_Type","Food_Preference",
            "Beverage_Preference","Avg_Spend_per_visit","Budget_Per_Visit",
            "Willingness_for_Membership","Membership_Price_Tolerance",
            "Late_Night_Study_Preference"
        ]
        vc = {}
        for col in CAT_COLS:
            if col in df.columns:
                counts = df[col].dropna().astype(str).str.strip().value_counts()
                vc[col] = {c: int(counts.get(c, 0)) for c in ordered_cats(df[col], col)}

        likert_means = {}
        for lc in LIKERT_COLS:
            if lc in df.columns:
                likert_means[lc] = round(float(df[lc].mean()), 2)

        # ── Age histogram — bucket edges derived from the data ────────────
        edges, width = age_bucket_edges(df["Age"]) if "Age" in df.columns else ([0], 1)
        age_labels   = age_bucket_labels(edges, width)
        age_counts   = [0] * len(age_labels)
        if "Age" in df.columns:
            for age in df["Age"].dropna():
                age_counts[age_bucket_index(age, edges, width)] += 1
        age_buckets = {lab: int(cnt) for lab, cnt in zip(age_labels, age_counts)}

        likert_dist = {}
        for lc in LIKERT_COLS:
            if lc in df.columns:
                likert_dist[lc] = [int((df[lc] == v).sum()) for v in [1, 2, 3, 4, 5]]

        # ── Cross-tabs — categories come from the data, so no respondent
        #    is ever left out of a chart because of a hard-coded label list.
        def _xtab(x_col, series_col):
            xl = ordered_cats(df[x_col], x_col); sl = ordered_cats(df[series_col], series_col)
            ct = pd.crosstab(df[series_col].astype(str).str.strip(), df[x_col].astype(str).str.strip())
            data = {s: [int(ct.loc[s, x]) if (s in ct.index and x in ct.columns) else 0 for x in xl] for s in sl}
            return xl, sl, data
        sh_l, vf_l, xtab_sh_vf       = _xtab("Study_Hours", "Visit_Frequency_per_week")
        dist_l, ln_l, xtab_dist_ln   = _xtab("Distance_from_Campus", "Late_Night_Study_Preference")
        gs_l, seat_l, xtab_gs_seat   = _xtab("Group_Study_Frequency", "Preferred_Seating_Type")
        bud_l, univ_l, xtab_univ_bud = _xtab("Budget_Per_Visit", "University")
        univs = sorted(univ_l)
        xtab_univ_bud = {u: xtab_univ_bud[u] for u in univs}

        analytics = {
            "n_rows": N,
            "age_mean": round(float(df["Age"].mean()),2) if "Age" in df.columns else 0,
            "age_std":  round(float(df["Age"].std()),2)  if "Age" in df.columns else 0,
            "age_min":  int(df["Age"].min())              if "Age" in df.columns else 0,
            "age_max":  int(df["Age"].max())              if "Age" in df.columns else 0,
            "age_histogram": age_buckets,
            "age_bucket_labels": age_labels,
            "likert_means": likert_means,
            "likert_dist":  likert_dist,
            "value_counts": vc,
            "value_count_order": {col: list(d_.keys()) for col, d_ in vc.items()},
            "xtab_sh_vf":       xtab_sh_vf,
            "xtab_dist_ln":     xtab_dist_ln,
            "xtab_gs_seat":     xtab_gs_seat,
            "xtab_univ_budget": xtab_univ_bud,
            "xtab_labels": {"sh": sh_l, "vf": vf_l, "dist": dist_l, "ln": ln_l,
                            "gs": gs_l, "seat": seat_l, "budget": bud_l},
            "xtab_rows_plotted": {
                "sh_vf":   int(sum(map(sum, xtab_sh_vf.values()))),
                "gs_seat": int(sum(map(sum, xtab_gs_seat.values()))),
            },
            "univs":            univs,
        }

        # ── CORR: Pearson on ordinal-encoded numeric features ──────────────
        num_cols = [c for c in ["Age","Power_Outlet_Requirement","WiFi_Importance",
            "Productivity_Tools_Interest","_dist","_sh","_vf","_spend","_budget","_gs","_ln","_mem"]
            if c in df.columns]
        corr_labels = ["Age","Power_Outlet","WiFi","Productivity","Distance from Campus",
            "Study Hours","Visit Freq/week","Avg Spend","Budget Per Visit",
            "Group Study Freq","Late Night Pref","Membership Willing"][:len(num_cols)]
        cdf = df[num_cols].dropna()
        if len(cdf) > 5:
            corr_mx = cdf.corr(method="pearson").round(3).fillna(0).values.tolist()
        else:
            corr_mx = [[1.0]*len(num_cols) for _ in num_cols]
        corr = {"labels": corr_labels, "matrix": corr_mx}

        # ── CLUSTERS: K-Means, PCA ─────────────────────────────────────────
        clusters = {"scatter2d":[],"scatter3d":[],"cluster_means":[],"elbow":[],"best_k":4}
        if SKLEARN_OK:
            from sklearn.cluster import KMeans
            from sklearn.decomposition import PCA
            from sklearn.preprocessing import StandardScaler
            feat_cols = [c for c in ["Age","Power_Outlet_Requirement","WiFi_Importance",
                "Productivity_Tools_Interest","_spend","_budget","_sh","_vf"] if c in df.columns]
            cdf2 = df[feat_cols + ["_mem"]].dropna().copy()
            if len(cdf2) >= 20:
                sc = StandardScaler()
                Xsc = sc.fit_transform(cdf2[feat_cols])
                elbow_data = []
                for k in range(2, min(9, len(cdf2))):
                    km = KMeans(n_clusters=k, random_state=42, n_init=10)
                    km.fit(Xsc)
                    elbow_data.append({"k": k, "inertia": round(float(km.inertia_), 1)})
                clusters["elbow"] = elbow_data
                best_k = min(4, max(2, len(elbow_data) + 1))
                if len(elbow_data) >= 3:
                    inertias = [e["inertia"] for e in elbow_data]
                    diffs = [inertias[i] - inertias[i+1] for i in range(len(inertias)-1)]
                    drops = [diffs[i] - diffs[i+1] for i in range(len(diffs)-1)]
                    best_k = elbow_data[drops.index(max(drops)) + 1]["k"]
                clusters["best_k"] = best_k
                km4 = KMeans(n_clusters=best_k, random_state=42, n_init=10)
                raw_labels = km4.fit_predict(Xsc)

                # K-Means numbers its clusters arbitrarily. Re-number them by
                # membership willingness (highest first) so cluster 0 always
                # means the same thing, whatever data is uploaded.
                mem_rate = {ci: float(cdf2["_mem"][raw_labels == ci].mean()) for ci in range(best_k)}
                order    = sorted(range(best_k), key=lambda ci: -mem_rate[ci])
                remap    = {old: new for new, old in enumerate(order)}
                labels4  = np.array([remap[l] for l in raw_labels])

                pca3 = PCA(n_components=min(3, len(feat_cols)), random_state=42)
                Xpca = pca3.fit_transform(Xsc)
                # What each principal component is made of, so the plots can say so.
                # A component is a weighted mix of the clustering variables; the variables
                # with the largest weights are what the axis mostly measures. The sign of a
                # component is arbitrary, so each is oriented so its weights sum positive.
                _pc_names = {"Age":"age","Power_Outlet_Requirement":"power-outlet need","WiFi_Importance":"WiFi importance",
                             "Productivity_Tools_Interest":"productivity-tools interest","_spend":"spend per visit",
                             "_budget":"budget per visit","_sh":"study hours","_vf":"visit frequency"}
                _pc_axes = []
                for _ci in range(pca3.components_.shape[0]):
                    _w = pca3.components_[_ci].astype(float)
                    if _w.sum() < 0:
                        _w = -_w; Xpca[:, _ci] = -Xpca[:, _ci]
                    _cut = 0.6 * float(np.abs(_w).max())
                    _big = sorted([(abs(v), j) for j, v in enumerate(_w) if abs(v) >= _cut], reverse=True)[:4]
                    _nm = lambda j: _pc_names.get(feat_cols[j], feat_cols[j].strip("_").replace("_", " ").lower())
                    _pc_axes.append({
                        "name": f"PC{_ci+1}",
                        "share_pct": round(float(pca3.explained_variance_ratio_[_ci]) * 100, 1),
                        "high": [_nm(j) for _, j in _big if _w[j] > 0],
                        "low":  [_nm(j) for _, j in _big if _w[j] < 0],
                        "weights": {_nm(j): round(float(v), 2) for j, v in enumerate(_w)},
                    })
                clusters["pca"] = {"axes": _pc_axes, "n_variables": len(feat_cols),
                                   "variables": [_pc_names.get(c, c.strip("_").replace("_", " ").lower()) for c in feat_cols],
                                   "total_share_pct": round(float(pca3.explained_variance_ratio_.sum()) * 100, 1)}
                if Xpca.shape[1] < 3:
                    Xpca = np.hstack([Xpca, np.zeros((len(Xpca), 3 - Xpca.shape[1]))])

                # Scatter plots: a seeded RANDOM sample of all rows (not the first 400)
                rng  = np.random.RandomState(42)
                n_pts = min(len(cdf2), 600)
                pick = np.sort(rng.choice(len(cdf2), size=n_pts, replace=False))
                s2, s3 = [], []
                for i in pick:
                    row = cdf2.iloc[i]
                    base = {"x": round(float(Xpca[i,0]),4), "y": round(float(Xpca[i,1]),4),
                            "cluster": int(labels4[i]),
                            "age": float(row.get("Age",0)), "wifi": float(row.get("WiFi_Importance",0))}
                    s2.append(dict(base, spend=float(row.get("_spend",0)), study=float(row.get("_sh",0))))
                    s3.append(dict(base, z=round(float(Xpca[i,2]),4)))
                clusters["scatter2d"] = s2
                clusters["scatter3d"] = s3
                clusters["scatter_sample"] = {"shown": int(n_pts), "total": int(len(cdf2))}

                # Cluster profiles: means, position within each feature's observed
                # range (0–100, comparable ACROSS clusters), and data-driven names.
                overall_mean = cdf2[feat_cols].mean(); overall_std = cdf2[feat_cols].std().replace(0, 1)
                f_min = cdf2[feat_cols].min(); f_max = cdf2[feat_cols].max()
                key = {"Age":"age","WiFi_Importance":"wifi","Power_Outlet_Requirement":"power",
                       "Productivity_Tools_Interest":"productivity","_spend":"spend",
                       "_budget":"budget","_sh":"study_hours","_vf":"visit_freq"}
                ages_all = df.loc[cdf2.index, "Age"] if "Age" in df.columns else None
                cm_list, used_names = [], set()
                for ci in range(best_k):
                    mask = labels4 == ci
                    sub  = cdf2[mask]
                    if len(sub) == 0: continue
                    entry = {"cluster": ci, "size": int(mask.sum()),
                             "membership": round(float(sub["_mem"].mean()), 2)}
                    norm = {}
                    for fc in feat_cols:
                        m = float(sub[fc].mean())
                        entry[key[fc]] = round(m, 1 if fc == "Age" else 2)
                        rngv = float(f_max[fc] - f_min[fc])
                        norm[key[fc]] = round((m - float(f_min[fc])) / rngv * 100, 1) if rngv > 0 else 50.0
                    for k_ in key.values(): entry.setdefault(k_, 0)
                    entry["norm"] = norm
                    z = ((sub[feat_cols].mean() - overall_mean) / overall_std)
                    top = z.abs().sort_values(ascending=False)
                    parts = [("High " if z[f] > 0 else "Low ") + FEATURE_LABELS.get(f, f)
                             for f in top.index[:2] if abs(z[f]) >= 0.25]
                    name = " · ".join(parts) if parts else "Average Profile"
                    if name in used_names:
                        extra = [f for f in top.index[2:] if abs(z[f]) >= 0.1]
                        if extra:
                            name += " · " + ("High " if z[extra[0]] > 0 else "Low ") + FEATURE_LABELS.get(extra[0], extra[0])
                    n_try = 2
                    base_name = name
                    while name in used_names:
                        name = f"{base_name} ({n_try})"; n_try += 1
                    used_names.add(name)
                    entry["name"] = name
                    # Age-bracket mix of the WHOLE cluster (every row, not a sample)
                    if ages_all is not None:
                        cnt = [0] * len(age_labels)
                        for a in ages_all[mask].dropna():
                            cnt[age_bucket_index(a, edges, width)] += 1
                        tot = sum(cnt) or 1
                        entry["age_dist"] = [round(c_ / tot * 100, 1) for c_ in cnt]
                    cm_list.append(entry)
                clusters["cluster_means"] = cm_list
                clusters["names"] = [c["name"] for c in cm_list]

        # ── RULES: pairwise association rules (deterministic order) ────────
        rules_list = []; rules_meta = {}
        try:
            from itertools import combinations
            arm_cols = [c for c in [
                "Avg_Spend_per_visit","Budget_Per_Visit","Preferred_Time",
                "Late_Night_Study_Preference","Willingness_for_Membership",
                "Food_Preference","Noise_Tolerance","Visit_Frequency_per_week",
                "Distance_from_Campus","Group_Study_Frequency","Preferred_Seating_Type",
                "Beverage_Preference","Study_Hours"
            ] if c in df.columns]
            tx_df = df[arm_cols].dropna().astype(str)
            n_tx  = len(tx_df)
            if n_tx > 50:
                min_sup, min_conf = 0.10, 0.45
                # A rule is reported only if it clearly beats chance: lift of at least 1.10 (10 % more likely than chance).
                MIN_LIFT = 1.10
                # Pairs of questions that the cleaning rules tie together. Any rule between
                # them (e.g. "Late Night" => late-night preference Yes, 100 %) is produced by
                # the cleaning, not by respondents, so it is not reported as a finding.
                STRUCTURAL_PAIRS = {frozenset(("Preferred_Time", "Late_Night_Study_Preference")),
                                    frozenset(("Avg_Spend_per_visit", "Budget_Per_Visit"))}
                n_structural = n_low_lift = 0
                try:
                    from scipy.stats import chi2_contingency as _chi2c
                except Exception:
                    _chi2c = None
                _pair_p = {}
                items = {}
                for col in arm_cols:
                    for val, cnt in tx_df[col].value_counts().items():
                        if cnt / n_tx >= min_sup:
                            items[f"{col}={val}"] = (col, val, cnt / n_tx)
                keys = list(items.keys())
                for a, b in combinations(keys, 2):
                    ca, va, sa = items[a]; cb, vb, sb = items[b]
                    if ca == cb: continue
                    _structural = frozenset((ca, cb)) in STRUCTURAL_PAIRS
                    sup = float(((tx_df[ca] == va) & (tx_df[cb] == vb)).sum()) / n_tx
                    if sup < min_sup: continue
                    _nab = int(round(sup * n_tx)); _na = int(round(sa * n_tx)); _nb = int(round(sb * n_tx))
                    _p = None
                    if _chi2c is not None:
                        try:
                            _tbl = [[_nab, _na - _nab], [_nb - _nab, n_tx - _na - _nb + _nab]]
                            _p = float(_chi2c(_tbl, correction=False)[1])
                            if not np.isfinite(_p): _p = None
                        except Exception:
                            _p = None
                    _pair_p[(a, b)] = _p
                    for ant, con, s_ant, s_con in [(a, b, sa, sb), (b, a, sb, sa)]:
                        conf = sup / s_ant
                        if conf < min_conf: continue
                        lift = conf / s_con
                        if _structural: n_structural += 1; continue
                        if lift < MIN_LIFT: n_low_lift += 1; continue
                        conv = round((1 - s_con) / max(1 - conf, 0.001), 4) if (1 - s_con) > 0.001 else 99.0
                        rules_list.append({
                            "antecedent": ant, "consequent": con,
                            "support": round(sup, 4), "confidence": round(conf, 4),
                            "lift": round(lift, 4), "conviction": round(conv, 4),
                            "n_both": _nab, "n_antecedent": int(round(s_ant * n_tx)),
                            "p_value": _p, "_pair": (a, b)})
                # Benjamini-Hochberg adjustment over every pair that was tested
                _tested = sorted([(p, k) for k, p in _pair_p.items() if p is not None])
                _m = len(_tested); _q = {}
                _run = 1.0
                for _rank in range(_m, 0, -1):
                    _p, _k = _tested[_rank - 1]
                    _run = min(_run, _p * _m / _rank)
                    _q[_k] = min(_run, 1.0)
                for _r in rules_list:
                    _qv = _q.get(_r.pop("_pair"))
                    _r["q_value"] = None if _qv is None else float(_qv)
                    _r["significant"] = None if _qv is None else bool(_qv < 0.05)
                rules_list.sort(key=lambda r: (-r["lift"], r["antecedent"], r["consequent"]))
                rules_list = rules_list[:200]
                rules_meta = {"shown": len(rules_list), "excluded_low_lift": n_low_lift,
                              "excluded_structural": n_structural, "min_lift": MIN_LIFT,
                              "min_support": min_sup, "min_confidence": min_conf,
                              "n_tx": int(n_tx), "n_tests": int(_m), "alpha": 0.05,
                              "test": "chi-square" if _chi2c is not None else None}
        except Exception:
            rules_list = []

        # ── Logistic Regression coefficients ──────────────────────────────
        reg_result = {"features":[], "coefs":[], "accuracy": 0, "budget_probs":[], "budget_labels":[], "baseline": 0}
        if SKLEARN_OK:
            from sklearn.linear_model import LogisticRegression
            from sklearn.preprocessing import StandardScaler
            from sklearn.model_selection import cross_val_score
            feat_reg = [c for c in ["Age","_vf","_spend","_sh","_dist","_gs",
                "Power_Outlet_Requirement","WiFi_Importance","Productivity_Tools_Interest",
                "_ln","_budget"] if c in df.columns]
            rdf = df[feat_reg+["_mem"]].dropna()
            if len(rdf) > 20 and rdf["_mem"].nunique() == 2:
                sc2 = StandardScaler()
                Xr  = sc2.fit_transform(rdf[feat_reg])
                yr  = rdf["_mem"]
                lr  = LogisticRegression(max_iter=500, random_state=42)
                try:
                    cv_acc = cross_val_score(lr, Xr, yr, cv=3, scoring="accuracy").mean()
                    lr.fit(Xr, yr)
                    feat_labels = {
                        "Age":"Age","_vf":"Visit Frequency","_spend":"Avg Spend",
                        "_sh":"Study Hours","_dist":"Distance from Campus",
                        "_gs":"Group Study Freq","Power_Outlet_Requirement":"Power Outlet",
                        "WiFi_Importance":"WiFi Importance","Productivity_Tools_Interest":"Productivity",
                        "_ln":"Late Night Pref","_budget":"Budget Per Visit"
                    }
                    pairs = sorted(zip(feat_reg, lr.coef_[0]), key=lambda x: -abs(x[1]))[:8]
                    reg_result["features"] = [feat_labels.get(p[0],p[0]) for p in pairs]
                    reg_result["coefs"]    = [round(float(p[1]),3) for p in pairs]
                    reg_result["accuracy"] = round(float(cv_acc),4)
                    reg_result["baseline"] = round(majority_share(yr), 4)
                    b_levels = sorted(rdf["_budget"].dropna().unique().tolist())
                    b_cats   = ordered_cats(df["Budget_Per_Visit"], "Budget_Per_Visit")
                    budget_probs = []
                    for bv in b_levels:
                        med_row = rdf[feat_reg].median().to_dict()
                        med_row["_budget"] = bv
                        row_sc = sc2.transform(pd.DataFrame([med_row])[feat_reg])
                        budget_probs.append(round(float(lr.predict_proba(row_sc)[0][1]), 3))
                    reg_result["budget_probs"]  = budget_probs
                    reg_result["budget_labels"] = b_cats if len(b_cats) == len(b_levels) else [str(b) for b in b_levels]
                except Exception:
                    pass

        try:
            pricing = _pricing_analysis(df)
        except Exception as _pex:
            pricing = {"error": str(_pex)}
        result = {
            "analytics": analytics,
            "pricing":   pricing,
            "corr":      corr,
            "clusters":  clusters,
            "rules":     rules_list,
            "rules_meta": rules_meta,
            "regression": reg_result,
        }
        result = _sanitise(result)
        _store["analytics_cache"] = result
        return jsonify(result)
    except Exception as ex:
        import traceback
        return jsonify({"error": str(ex), "traceback": traceback.format_exc()[-1200:]}), 500

@app.route("/download")
def download():
    if "clean_bytes" not in _store:
        return jsonify({"error": "No cleaned data"}), 400
    # Always create a fresh BytesIO from the stored bytes so repeated
    # downloads never fail due to a consumed/closed buffer.
    buf = io.BytesIO(_store["clean_bytes"])
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="study_cafe_cleaned.xlsx"
    )

@app.route("/shutdown", methods=["POST"])
def shutdown():
    """Free the port when the GUI tab is closed."""
    func = request.environ.get("werkzeug.server.shutdown")
    if func:
        func()
    else:
        threading.Thread(target=lambda: (__import__('time').sleep(0.5), os._exit(0)), daemon=True).start()
    return "Shutting down..."

BUNDLED_NAME = "study_cafe_survey_dataset.xlsx"

def _bundled_path():
    """The survey file shipped with the app: next to the exe (desktop build),
    next to this script, or wherever SC_BUNDLED_DATASET points."""
    import os, sys
    cands = []
    if os.environ.get("SC_BUNDLED_DATASET"):
        cands.append(os.environ["SC_BUNDLED_DATASET"])
    if getattr(sys, "frozen", False):
        d = os.path.dirname(sys.executable)
        cands += [os.path.join(d, BUNDLED_NAME), os.path.join(d, "_internal", BUNDLED_NAME)]
    try:
        cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), BUNDLED_NAME))
    except NameError:
        pass
    for p in cands:
        if p and os.path.isfile(p):
            return p
    return None

@app.route("/bundled_info")
def bundled_info():
    p = _bundled_path()
    if not p:
        return jsonify({"available": False})
    import os
    return jsonify({"available": True, "name": os.path.basename(p), "bytes": os.path.getsize(p)})

@app.route("/bundled_dataset")
def bundled_dataset():
    p = _bundled_path()
    if not p:
        return jsonify({"error": "No bundled dataset"}), 404
    return send_file(p, as_attachment=False, download_name=BUNDLED_NAME,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/reset", methods=["POST"])
def reset():
    """Clear all stored data so a new dataset can be uploaded fresh."""
    _store.clear()
    return jsonify({"ok": True})



# ════════════════════════════════════════════════════════════
#  WHAT-IF SCENARIO MODELLING ROUTE
# ════════════════════════════════════════════════════════════

# Assumed share of a respondent's weekly visits falling on Mon..Sun, chosen by how
# many visits a week they make. These are planning assumptions, not survey findings.
DAY_PROFILE_ASSUMPTIONS = [
    (6.0,  [1/7]*7),                                   # near-daily: every day alike
    (3.5,  [0.20, 0.22, 0.22, 0.20, 0.16, 0.00, 0.00]),  # 4-5 a week: weekdays
    (1.75, [0.15, 0.20, 0.25, 0.20, 0.15, 0.03, 0.02]),  # 2-3 a week: midweek
    (0.75, [0.10, 0.15, 0.20, 0.20, 0.15, 0.10, 0.10]),  # once a week
    (float("-inf"), [1/7]*7),                          # rarely: no pattern assumed
]
VARIABLE_COST_RATE = 0.35      # cost of goods as a share of walk-in revenue
STAFF_RATIO        = 20        # customers one staff member can serve at once
SLOT_WEIGHTS       = {"Morning": 4, "Afternoon": 4, "Evening": 5, "Late Night": 3}

@app.route("/whatif", methods=["POST"])
def whatif():
    import warnings; warnings.filterwarnings("ignore")
    if "clean_bytes" not in _store:
        return jsonify({"error": "Run cleaning first"}), 400
    try:
        body = request.get_json(force=True, silent=True) or {}
        # ── Inputs — every value is range-checked so no input can produce
        #    impossible results (negative members, >100 % uptake, etc.)
        catchment   = int(_clamp(body.get("catchment", 1000), 0, 10_000_000))
        growth_pct  = _clamp(body.get("monthly_growth", 5.0), -50, 100) / 100
        price_adj   = _clamp(body.get("price_adjustment", 0.0), -90, 300) / 100
        op_hours    = _clamp(body.get("operating_hours", 12), 1, 24)
        mem_target  = _clamp(body.get("membership_target", 50.0), 0, 100) / 100
        occupancy   = _clamp(body.get("occupancy_rate", 70.0), 0, 100) / 100
        seats       = int(_clamp(body.get("total_seats", 60), 0, 100_000))
        fixed_costs = _clamp(body.get("fixed_costs", 150000), 0, 1e12)
        comp_impact = _clamp(body.get("competitor_impact", 0.0), 0, 100) / 100
        staff_cost  = _clamp(body.get("staff_cost_per_shift", 800), 0, 1e9)
        tier1_price = _clamp(body.get("tier1_price", 750), 0, 1e9)
        tier1_uptake= _clamp(body.get("tier1_uptake", 60.0), 0, 100) / 100
        tier2_price = _clamp(body.get("tier2_price", 1500), 0, 1e9)

        df = pd.read_excel(io.BytesIO(_store["clean_bytes"]), sheet_name="Cleaned_Data")
        N  = len(df)
        df["visits_pw"] = num_series(df, "Visit_Frequency_per_week")
        df["spend_v"]   = num_series(df, "Avg_Spend_per_visit")
        df["study_h"]   = num_series(df, "Study_Hours")
        w = df[["visits_pw", "spend_v"]].dropna()
        visits = w["visits_pw"].values; spend = w["spend_v"].values

        base_avg_vpw   = float(visits.mean()) if len(w) else 0.0
        # spend of the average VISIT (frequent visitors weigh more), not of the average respondent
        base_avg_spend = float((visits * spend).sum() / visits.sum()) if len(w) and visits.sum() > 0 else 0.0
        # Average session length, weighted by how often each respondent visits
        sess = df[["study_h", "visits_pw"]].dropna()
        avg_session_h = float(np.average(sess["study_h"], weights=sess["visits_pw"])) \
                        if len(sess) and sess["visits_pw"].sum() > 0 else 2.0
        avg_session_h = max(avg_session_h, 0.25)
        willing_rate  = float(yes_series(df["Willingness_for_Membership"]).mean()) if N else 0.0

        # ── Demand response to price, taken from the survey itself ─────────
        # Each respondent states a budget band. Their true ceiling is assumed to
        # lie evenly between what they already spend and the top of that band.
        # Raising prices by p % keeps a respondent with probability
        #   P(ceiling >= spend x (1+p)).   Price cuts are not assumed to add visits.
        bands = budget_bands(df) if "Budget_Per_Visit" in df.columns else {}
        b_lab = df.loc[w.index, "Budget_Per_Visit"].astype(str).str.strip() if bands else None
        lo_arr = np.array([bands.get(b, (np.nan, np.nan))[0] for b in b_lab]) if bands else None
        hi_arr = np.array([bands.get(b, (np.nan, np.nan))[1] for b in b_lab]) if bands else None

        def retention(p):
            """Share of each respondent's visits kept after a price change of p."""
            if p <= 0 or not bands or len(w) == 0:
                return np.ones(len(w))
            new_price = spend * (1 + p)
            floor = np.maximum(np.nan_to_num(lo_arr, nan=0.0), spend)
            ceil_ = np.where(np.isnan(hi_arr), np.inf, np.maximum(hi_arr, spend))
            span  = ceil_ - floor
            with np.errstate(divide="ignore", invalid="ignore"):
                keep = np.where(span > 0, (ceil_ - new_price) / span, (new_price <= ceil_).astype(float))
            keep = np.where(np.isinf(ceil_), 1.0, keep)
            return np.clip(keep, 0.0, 1.0)

        capacity_day = seats * occupancy * (op_hours / avg_session_h)   # visits the seats can host

        def simulate(catch_eff, p, foot_mult=1.0, cap=None):
            """One consistent engine used by every table and chart."""
            keep = retention(p)
            kept_visits = float((visits * keep).sum())
            vpw_per_student = kept_visits / len(w) if len(w) else 0.0
            avg_sp = float((visits * keep * spend * (1 + p)).sum() / kept_visits) if kept_visits > 0 else 0.0
            demand = vpw_per_student * catch_eff / 7.0 * foot_mult
            monthly, served_m = [], []
            for m in range(12):
                d_m = demand * (1 + growth_pct) ** m
                s_m = min(d_m, capacity_day if cap is None else cap)
                served_m.append(s_m); monthly.append(s_m * avg_sp * 30)
            return {"demand": demand, "served": served_m[0], "avg_spend": avg_sp,
                    "monthly": monthly, "served_m": served_m, "annual": float(sum(monthly)),
                    "retained_pct": (kept_visits / float(visits.sum()) * 100) if visits.sum() > 0 else 100.0}

        eff_catchment = catchment * (1 - comp_impact)
        base = simulate(eff_catchment, price_adj)
        adj_spend    = base["avg_spend"]
        daily_demand = round(base["demand"], 1)
        daily_foot   = round(base["served"], 1)
        daily_rev    = round(base["served"] * adj_spend, 0)
        weekly_rev   = round(daily_rev * 7, 0)
        monthly_rev  = round(base["monthly"][0], 0)
        monthly_proj = [round(v, 0) for v in base["monthly"]]
        annual_rev   = round(base["annual"], 0)

        # ── Membership tiers ───────────────────────────────────────────────
        # Target applies to the students who said they are WILLING to join.
        total_members = int(round(eff_catchment * willing_rate * mem_target))
        tier1_count   = int(round(total_members * tier1_uptake))
        tier2_count   = total_members - tier1_count
        tier1_monthly = round(tier1_count * tier1_price, 0)
        tier2_monthly = round(tier2_count * tier2_price, 0)
        tier_total_monthly = tier1_monthly + tier2_monthly
        pt = num_series(df, "Membership_Price_Tolerance").dropna() if "Membership_Price_Tolerance" in df.columns else pd.Series(dtype=float)
        avg_price_tol = float(pt.mean()) if len(pt) else 0.0
        single_tier_monthly = round(total_members * avg_price_tol, 0)

        # ── Staffing — slots share the operating hours the user entered ────
        time_dist  = df["Preferred_Time"].astype(str).str.strip().value_counts(normalize=True)
        slot_names = ordered_cats(df["Preferred_Time"], "Preferred_Time")
        w_tot = sum(SLOT_WEIGHTS.get(s, 4) for s in slot_names) or 1
        slots_info = []
        for slot in slot_names:
            slot_hours = op_hours * SLOT_WEIGHTS.get(slot, 4) / w_tot
            slot_pct   = float(time_dist.get(slot, 0))
            slot_foot  = base["served"] * slot_pct
            concurrent = slot_foot * min(avg_session_h, slot_hours) / slot_hours if slot_hours > 0 else 0
            concurrent = min(concurrent, seats * occupancy) if seats > 0 else concurrent
            staff_needed = max(1, int(np.ceil(concurrent / STAFF_RATIO))) if slot_foot > 0 else 0
            labour_cost  = round(staff_needed * staff_cost, 0)          # one shift per slot
            slot_rev_val = round(slot_foot * adj_spend, 0)
            slots_info.append({
                "slot": slot, "hours": round(slot_hours, 1), "footfall": int(round(slot_foot)),
                "staff_needed": staff_needed, "labour_cost": int(labour_cost),
                "slot_revenue": int(slot_rev_val),
                "profit_after_labour": int(slot_rev_val - labour_cost),
            })
        total_daily_labour = sum(s["labour_cost"] for s in slots_info)
        labour_monthly     = total_daily_labour * 30

        def labour_daily_for(served, seats_=None, occ_=None):
            """Daily staff cost for a given number of visitors served (same rules as the table above)."""
            tot = 0.0
            for slot in slot_names:
                sh = op_hours * SLOT_WEIGHTS.get(slot, 4) / w_tot
                sf = served * float(time_dist.get(slot, 0))
                cc = sf * min(avg_session_h, sh) / sh if sh > 0 else 0
                _st = seats if seats_ is None else seats_; _oc = occupancy if occ_ is None else occ_
                cc = min(cc, _st * _oc) if _st > 0 else cc
                tot += (max(1, int(np.ceil(cc / STAFF_RATIO))) if sf > 0 else 0) * staff_cost
            return tot
        staffing = {
            "slots": slots_info, "total_daily_labour": int(total_daily_labour),
            "staff_cost_per_shift": int(staff_cost),
            "net_daily_revenue": int(daily_rev - total_daily_labour),
        }

        # ── Profit & break-even — ONE definition used everywhere ───────────
        # profit = walk-in revenue x (1 - variable cost) + membership fees
        #          - fixed costs - labour
        def profit_of(rev_month, served_day=None, seats_=None, occ_=None):
            lab = labour_monthly if served_day is None else labour_daily_for(served_day, seats_, occ_) * 30
            return rev_month * (1 - VARIABLE_COST_RATE) + tier_total_monthly - fixed_costs - lab
        profit_monthly = [round(profit_of(v, s_), 0) for v, s_ in zip(base["monthly"], base["served_m"])]
        breakeven_month = next((m + 1 for m, p_ in enumerate(profit_monthly) if p_ >= 0), None)
        variable_cost_m1 = monthly_rev * VARIABLE_COST_RATE
        total_monthly_cost = fixed_costs + labour_monthly + variable_cost_m1
        margin_per_visit = adj_spend * (1 - VARIABLE_COST_RATE)
        need = max(0.0, fixed_costs + labour_monthly - tier_total_monthly)
        breakeven_visitors = int(np.ceil(need / (margin_per_visit * 30))) if margin_per_visit > 0 else None
        breakeven = {
            "fixed_costs_monthly":      int(fixed_costs),
            "labour_cost_monthly":      int(labour_monthly),
            "variable_cost_monthly":    int(variable_cost_m1),
            "membership_income_monthly":int(tier_total_monthly),
            "variable_cost_rate_pct":   int(VARIABLE_COST_RATE * 100),
            "total_cost_monthly":       int(total_monthly_cost),
            "breakeven_daily_visitors": breakeven_visitors,
            "breakeven_within_capacity": (breakeven_visitors is not None and breakeven_visitors <= capacity_day),
            "breakeven_month":          breakeven_month,
            "profit_monthly_projection": profit_monthly,
            "is_profitable_at_base":    bool(profit_monthly[0] >= 0),
        }

        # ── Pricing sweep (same engine, so it includes the demand response) ─
        price_sweep = []
        for pct_adj in [-0.30, -0.20, -0.10, 0, 0.10, 0.20, 0.30, 0.40, 0.50]:
            s = simulate(eff_catchment, pct_adj)
            mr = [float(v) for v in s["monthly"]]
            mp = [float(profit_of(v, s_)) for v, s_ in zip(s["monthly"], s["served_m"])]
            full = next((m + 1 for m in range(12) if s["served_m"][m] >= capacity_day - 1e-9), None)
            price_sweep.append({
                "label":       f"{'+' if pct_adj >= 0 else ''}{int(round(pct_adj*100))}%",
                "adj_spend":   round(s["avg_spend"], 0),
                "demand_retained_pct": round(s["retained_pct"], 1),
                "monthly_rev": int(round(mr[0])),                     # month 1 (kept for older views)
                "profit":      int(round(mp[0])),
                "annual_rev":    int(round(sum(mr))), "annual_profit": int(round(sum(mp))),
                "q_rev":    [int(round(sum(mr[q*3:q*3+3]))) for q in range(4)],
                "q_profit": [int(round(sum(mp[q*3:q*3+3]))) for q in range(4)],
                "m_rev":    [int(round(v)) for v in mr], "m_profit": [int(round(v)) for v in mp],
                "full_from_month": full,
            })
        # The best price is the one with the most profit over the whole year, not in month 1.
        best_price = max(price_sweep, key=lambda r: r["annual_profit"])["label"] if price_sweep else None
        # Once the cafe is full, a different price can earn more each month; report that too.
        best_price_full = max(price_sweep, key=lambda r: r["m_profit"][-1])["label"] if price_sweep else None

        # ── Competitor sweep — measured against a true 0 % baseline ────────
        comp_base = simulate(catchment, price_adj)["annual"]
        competitor_scenarios = []
        for c_pct in sorted(set([0, 10, 20, 30, 40, int(round(comp_impact * 100))])):
            c_ann = simulate(catchment * (1 - c_pct / 100), price_adj)["annual"]
            competitor_scenarios.append({
                "label":         f"-{c_pct}% catchment",
                "catchment_eff": int(round(catchment * (1 - c_pct / 100))),
                "annual_rev":    int(round(c_ann)),
                "revenue_loss":  int(round(comp_base - c_ann)),
                "is_current":    bool(c_pct == int(round(comp_impact * 100))),
            })

        # ── v56: profit for the year, price against one other setting ───────
        def annual_profit_at(catch_eff, p, occ_=None, seats_=None):
            o_ = occupancy if occ_ is None else occ_; st_ = seats if seats_ is None else seats_
            s = simulate(catch_eff, p, cap=st_ * o_ * (op_hours / avg_session_h))
            return float(sum(profit_of(v, s_, st_, o_) for v, s_ in zip(s["monthly"], s["served_m"])))
        p_rows = sorted(set([-0.20, -0.10, 0.0, 0.10, 0.20, 0.30, round(price_adj, 4)]))
        cur_row = p_rows.index(round(price_adj, 4))
        def _grid(col_vals, fn):
            return [[int(round(fn(p, v))) for v in col_vals] for p in p_rows]
        c_mult = [0.5, 0.75, 1.0, 1.25, 1.5]
        c_vals = [catchment * m for m in c_mult]
        o_vals = sorted(set(round(o, 4) for o in [occupancy - 0.2, occupancy - 0.1, occupancy, occupancy + 0.1, occupancy + 0.2]
                             if 0.05 <= round(o, 4) <= 1.0))
        s_vals = sorted(set(int(round(seats * m)) for m in [0.5, 0.75, 1.0, 1.5, 2.0] if seats * m >= 1))
        sens_grid = {
            "rows": [f"{'+' if p >= 0 else '−'}{abs(int(round(p * 100)))}%" for p in p_rows], "current_row": cur_row,
            "grids": {
                "catchment": {"title": "Students in the catchment", "cols": [f"{int(round(v)):,}" for v in c_vals],
                              "current_col": 2, "cells": _grid(c_vals, lambda p, v: annual_profit_at(v * (1 - comp_impact), p))},
                "occupancy": {"title": "Seats in use on average", "cols": [f"{int(round(o * 100))}%" for o in o_vals],
                              "current_col": o_vals.index(round(occupancy, 4)) if round(occupancy, 4) in o_vals else None,
                              "cells": _grid(o_vals, lambda p, v: annual_profit_at(eff_catchment, p, occ_=v))},
                "seats": {"title": "Seats", "cols": [str(v) for v in s_vals],
                          "current_col": s_vals.index(seats) if seats in s_vals else None,
                          "cells": _grid(s_vals, lambda p, v: annual_profit_at(eff_catchment, p, seats_=v))},
            },
        }

        scenario_summary = {
            "catchment": catchment, "growth_pct": round(growth_pct * 100, 1),
            "price_adj_pct": round(price_adj * 100, 1), "comp_impact_pct": round(comp_impact * 100, 1),
            "fixed_costs": int(fixed_costs), "daily_revenue": int(daily_rev),
            "monthly_revenue": int(monthly_rev), "annual_revenue": int(annual_rev),
            "profit_monthly_m1": int(profit_monthly[0]), "breakeven_month": breakeven_month,
            "members": int(total_members), "tier_combined_monthly": int(tier_total_monthly),
        }

        utilisation = (base["served"] / capacity_day * 100) if capacity_day > 0 else 0.0
        seating = {
            "seats_used_daily":           int(round(seats * occupancy)),
            "avg_session_hours":          round(avg_session_h, 2),
            "avg_turnover_per_seat":      round(op_hours / avg_session_h, 1),
            "effective_capacity_per_day": int(round(capacity_day)),
            "daily_demand":               int(round(base["demand"])),
            "visitors_served":            int(round(base["served"])),
            "unmet_demand":               int(round(max(0.0, base["demand"] - base["served"]))),
            "capacity_utilisation_pct":   round(utilisation, 1),
            "capacity_limited":           bool(base["demand"] > capacity_day),
            "capacity_reached_month":     next((m + 1 for m in range(12)
                                               if base["demand"] * (1 + growth_pct) ** m > capacity_day), None),
        }

        sensitivity = {}
        for delta, label in [(-0.20,"-20%"),(-0.10,"-10%"),(0,"Base"),(0.10,"+10%"),(0.20,"+20%")]:
            s = simulate(eff_catchment, price_adj, foot_mult=1 + delta)
            sensitivity[label] = {
                "daily_footfall":  round(s["served"], 0),
                "daily_revenue":   round(s["served"] * s["avg_spend"], 0),
                "monthly_revenue": round(s["monthly"][0], 0),
            }
        slot_rev = {}
        for t in slot_names:
            sub = df[df["Preferred_Time"].astype(str).str.strip() == t]["spend_v"]
            slot_foot  = round(base["served"] * float(time_dist.get(t, 0)), 0)
            slot_spend = float(sub.mean() * (1 + price_adj)) if len(sub) else adj_spend
            slot_rev[t] = {"footfall": slot_foot, "revenue": round(slot_foot * slot_spend, 0)}

        notes = [
            f"Demand = catchment x average visits/week ({base_avg_vpw:.2f}) / 7. "
            f"Seats can host {int(round(capacity_day))} visits/day "
            f"(seats x occupancy x operating hours / average session of {avg_session_h:.1f} h); "
            "revenue is earned only on visitors who can be seated.",
            "Price rises lose visitors whose stated budget no longer covers the new price; "
            "price cuts are not assumed to add visits.",
            f"Profit = walk-in revenue x {int((1-VARIABLE_COST_RATE)*100)}% + membership fees - fixed costs - labour. "
            f"Membership target applies to the {willing_rate*100:.0f}% of students who said they would join.",
        ]

        return jsonify(_sanitise({
            "inputs": {
                "catchment": catchment, "monthly_growth_pct": round(growth_pct*100,1),
                "price_adjustment_pct": round(price_adj*100,1),
                "operating_hours": op_hours, "membership_target_pct": round(mem_target*100,1),
                "occupancy_rate_pct": round(occupancy*100,1), "total_seats": seats,
                "fixed_costs": int(fixed_costs), "competitor_impact_pct": round(comp_impact*100,1),
                "tier1_uptake_pct": round(tier1_uptake*100,1),
            },
            "base_metrics": {
                "avg_visits_per_week": round(base_avg_vpw,2),
                "avg_spend_base": round(base_avg_spend,1),
                "avg_spend_adjusted": round(adj_spend,1),
                "demand_retained_pct": round(base["retained_pct"],1),
                "willing_rate_pct": round(willing_rate*100,1),
            },
            "footfall": {"daily": daily_foot, "daily_demand": daily_demand},
            "revenue": {
                "daily": int(daily_rev), "weekly": int(weekly_rev),
                "monthly": int(monthly_rev), "annual": int(annual_rev),
                "monthly_projection": monthly_proj,
                # what the same demand would earn if every visitor could be seated
                "monthly_projection_uncapped": [round(base["demand"] * (1 + growth_pct) ** m * base["avg_spend"] * 30, 0) for m in range(12)],
            },
            "membership": {
                "members_estimated": int(total_members),
                "avg_price_tolerance": round(avg_price_tol,1),
                "monthly_revenue": int(tier_total_monthly),
                "annual_revenue":  int(tier_total_monthly * 12),
            },
            "seating": seating,
            "sensitivity": sensitivity,
            "sens_grid":   sens_grid,
            "slot_breakdown": slot_rev,
            "breakeven":     breakeven,
            "price_sweep":   price_sweep,
            "best_price_label": best_price,
            "best_price_full_label": best_price_full,
            "competitor":    competitor_scenarios,
            "staffing":      staffing,
            "tiers": {
                "total_members": int(total_members),
                "tier1_name": "Basic", "tier1_price": int(tier1_price),
                "tier1_count": int(tier1_count), "tier1_monthly_rev": int(tier1_monthly),
                "tier2_name": "Premium", "tier2_price": int(tier2_price),
                "tier2_count": int(tier2_count), "tier2_monthly_rev": int(tier2_monthly),
                "combined_monthly_rev": int(tier_total_monthly),
                "combined_annual_rev":  int(tier_total_monthly * 12),
                "single_tier_monthly":  int(single_tier_monthly),
                "uplift_vs_single":     int(tier_total_monthly - single_tier_monthly),
            },
            "scenario_summary": scenario_summary,
            "model_notes": notes,
            "arima_note": (
                "ARIMA validity for this dataset: Cross-sectional survey data "
                "records preferences at a single point in time — it has no time axis. "
                "ARIMA requires a time series (observations ordered chronologically). "
                "Applying ARIMA here would mean treating the 12-month projection as a "
                "time series, but those projections are themselves derived from a fixed "
                "growth assumption, not observed data. ARIMA would fit a model to a "
                "deterministic formula rather than real variation, producing statistically "
                "meaningless confidence intervals. Verdict: ARIMA is NOT validly applicable "
                "to this dataset. For genuine time-series forecasting, you would need at "
                "least 24-36 months of actual sales/footfall observations."
            ),
        }))
    except Exception as ex:
        import traceback
        return jsonify({"error": str(ex), "traceback": traceback.format_exc()[-800:]}), 500


# ════════════════════════════════════════════════════════════
#  PREDICTIVE ANALYTICS ROUTE
# ════════════════════════════════════════════════════════════

class _NoModel:
    """Stand-in estimator for fast mode: accepts any arguments, trains nothing."""
    def __init__(self, *a, **k): pass
    def get_params(self, deep=True): return {}
    def set_params(self, **k): return self
    def fit(self, X, y=None, *a, **k):
        n = X.shape[1] if hasattr(X, "shape") and len(getattr(X, "shape", ())) > 1 else 1
        self.feature_importances_ = np.zeros(n); self.coef_ = np.zeros((1, n)); self.intercept_ = np.zeros(1)
        self.classes_ = np.unique(np.asarray(y)) if y is not None and len(np.asarray(y)) else np.array([0])
        return self
    def predict(self, X):
        c = getattr(self, "classes_", np.array([0]))
        return np.full(len(X), c[0] if len(c) else 0)
    def predict_proba(self, X):
        c = getattr(self, "classes_", np.array([0, 1]))
        p = np.zeros((len(X), max(len(c), 1))); p[:, 0] = 1.0; return p
    def score(self, X, y=None): return 0.0

_EST_NAMES = ["LogisticRegression", "Ridge", "RandomForestClassifier", "GradientBoostingClassifier",
              "RandomForestRegressor", "DecisionTreeClassifier", "DecisionTreeRegressor",
              "KNeighborsClassifier", "KNeighborsRegressor", "GaussianNB", "SVC", "SVR"]
def _estimators(fast):
    if fast:
        d = {n: _NoModel for n in _EST_NAMES}
        d["cross_val_score"] = lambda est, X, y=None, *a, **k: np.array([0.0])
        d["train_test_split"] = lambda X, y, *a, **k: (X, X, y, y)
        return d
    g = globals()
    return {n: g[n] for n in _EST_NAMES + ["cross_val_score", "train_test_split"]}

def _predictive_compute(ex, fast=False):
    _E = _estimators(fast)
    (LogisticRegression, Ridge, RandomForestClassifier, GradientBoostingClassifier, RandomForestRegressor,
     DecisionTreeClassifier, DecisionTreeRegressor, KNeighborsClassifier, KNeighborsRegressor,
     GaussianNB, SVC, SVR) = [_E[n] for n in _EST_NAMES]
    cross_val_score = _E["cross_val_score"]; train_test_split = _E["train_test_split"]
    import warnings; warnings.filterwarnings("ignore")
    try:

        _cb = _store["clean_bytes"]
        if _store.get("_pred_df_src") is not _cb:
            _store["_pred_df"] = pd.read_excel(io.BytesIO(_cb), sheet_name="Cleaned_Data")
            _store["_pred_df_src"] = _cb
        df = _store["_pred_df"].copy()
        N  = len(df)
        result = {}

        if _store.get("_pred_prep_src") is not _cb:
            # ── Ordinal mappings ──────────────────────────────────────────────────────
            # v41: numeric encodings are data-driven (see num_series / yes_series)
            df["visits_pw"]  = num_series(df, "Visit_Frequency_per_week")
            df["spend_v"]    = num_series(df, "Avg_Spend_per_visit")
            df["study_h"]    = num_series(df, "Study_Hours")
            df["dist_n"]     = num_series(df, "Distance_from_Campus")
            df["gs_n"]       = num_series(df, "Group_Study_Frequency")
            df["time_n"]     = num_series(df, "Preferred_Time")
            df["late_n"]     = yes_series(df["Late_Night_Study_Preference"])
            df["mem_n"]      = yes_series(df["Willingness_for_Membership"])
            for _c in ["Preferred_Time","Distance_from_Campus","Avg_Spend_per_visit","Preferred_Seating_Type",
                       "Group_Study_Frequency","Visit_Frequency_per_week"]:
                df[_c] = df[_c].astype(str).str.strip()
            _store["_pred_prep"] = df.copy(); _store["_pred_prep_src"] = _cb
        else:
            df = _store["_pred_prep"].copy()

        # v48 levers: leave out the respondents whose answers the user unticked
        N_full = N
        _sec_err = {}
        if ex:
            _colmap = {"time": "Preferred_Time", "spend": "Avg_Spend_per_visit", "seat": "Preferred_Seating_Type"}
            for _k, _vals in ex.items():
                if _vals and _colmap.get(_k) in df.columns:
                    df = df[~df[_colmap[_k]].isin([str(v).strip() for v in _vals])]
            df = df.reset_index(drop=True)
            N = len(df)
            if N < 30:
                raise _TooFew(N)

        # ── Subsample for ML only (keep full df for descriptive stats) ──────
        # Caps training data at 2500 rows to prevent timeout on large uploads.
        # All frequency counts / distributions always use the full cleaned dataset.
        MAX_ML_ROWS = 2500
        if N > MAX_ML_ROWS:
            df_ml = df.dropna(subset=["Age","visits_pw","spend_v","study_h",
                                      "dist_n","gs_n","time_n","late_n","mem_n"])
            df_ml = df_ml.sample(n=MAX_ML_ROWS, random_state=42)
        else:
            df_ml = df.copy()

        try:
            # ═══════════════════════════════════════════════════════════════
            # 1. PEAK HOURS — distribution + RF multi-class classifier
            # ═══════════════════════════════════════════════════════════════
            time_vc   = df["Preferred_Time"].value_counts()
            time_pct  = {k: round(v/N*100,1) for k,v in time_vc.items()}
            time_order= ordered_cats(df["Preferred_Time"], "Preferred_Time")

            # v41: Late_Night_Study_Preference is NOT used here — it restates the
            # answer (every "Late Night" respondent has it set to Yes), which is leakage.
            feat_pk = df_ml[["Age","study_h","dist_n","gs_n","visits_pw",
                          "Power_Outlet_Requirement","WiFi_Importance"]].dropna()
            y_pk    = df_ml["Preferred_Time"].loc[feat_pk.index]
            from sklearn.preprocessing import LabelEncoder
            le = LabelEncoder()
            y_pk_enc = le.fit_transform(y_pk)
            sc1 = StandardScaler()
            Xpk = sc1.fit_transform(feat_pk)
            rf_pk = RandomForestClassifier(n_estimators=100, max_depth=6, random_state=42)
            cv_pk = cross_val_score(rf_pk, Xpk, y_pk_enc, cv=3, scoring="accuracy")
            rf_pk.fit(Xpk, y_pk_enc)
            fi_pk = {list(feat_pk.columns)[i]: round(float(v),4) for i,v in enumerate(rf_pk.feature_importances_)}

            # Spend by time slot
            sbt = df.groupby("Preferred_Time")["spend_v"].mean().round(1)
            spend_by_time = {t: round(float(sbt.get(t, 0)),1) for t in time_order}

            # Projected hourly revenue per 100 customers for each slot
            # Assume 3hr session per slot, customers spread within slot
            hrs_per_slot = 3
            rev_per_slot = {}
            for t in time_order:
                pct = time_pct.get(t,0)/100
                avg_cust_100 = pct * 100
                avg_spend    = spend_by_time.get(t,300)
                rev_per_slot[t] = round(avg_cust_100 * avg_spend, 0)

            top_time_slot   = max(time_pct, key=time_pct.get)
            top_time_pct    = time_pct[top_time_slot]
            min_time_slot   = min(time_pct, key=time_pct.get)
            min_time_pct    = time_pct[min_time_slot]
            pct_values      = list(time_pct.values())
            pct_range_t     = round(max(pct_values) - min(pct_values), 1)
            distribution_desc = ("nearly uniformly distributed" if pct_range_t < 5
                                 else "unevenly distributed")
            result["peak_hours"] = {
                "distribution":    {t: time_pct.get(t,0) for t in time_order},
                "spend_by_slot":   spend_by_time,
                "revenue_per_100": rev_per_slot,
                "rf_accuracy":     round(float(cv_pk.mean()),4),
                "rf_std":          round(float(cv_pk.std()),4),
                "baseline":        round(majority_share(y_pk), 4),
                "feature_importances": dict(sorted(fi_pk.items(), key=lambda x:-x[1])),
                "top_time_slot":   top_time_slot,
                "insight": (f"Time preferences are {distribution_desc} across slots. "
                            f"{top_time_slot} shows the highest preference ({top_time_pct:.1f}%), "
                            f"{min_time_slot} the lowest ({min_time_pct:.1f}%). "
                            f"The model predicts a student's slot correctly {cv_pk.mean()*100:.1f}% of the time, "
                            f"against {majority_share(y_pk)*100:.1f}% from always guessing the most popular slot.")
            }

        except Exception as _sec_ex:
            _sec_err['peak_hours'] = str(_sec_ex)[:300]

        try:
            # ═══════════════════════════════════════════════════════════════
            # 2. FOOTFALL / TRAFFIC FORECAST
            # ═══════════════════════════════════════════════════════════════
            vf_vc  = df["Visit_Frequency_per_week"].value_counts()
            vf_pct = {k: round(v/N*100,1) for k,v in vf_vc.items()}
            avg_vpw = round(float(df["visits_pw"].mean()),2)

            # Projected daily footfall for a 1000-respondent catchment
            catchment = 1000 * N / max(N_full, 1)
            daily_proj = round(avg_vpw * catchment / 7, 0)

            # Footfall by distance (who will actually come)
            dist_vf = df.groupby("Distance_from_Campus")["visits_pw"].mean().round(2)

            # Footfall by time slot (projected customers in each slot per day for 1000 catchment)
            slot_footfall = {}
            for t in time_order:
                pct_t = time_pct.get(t,0)/100
                slot_footfall[t] = round(daily_proj * pct_t, 0)

            # Footfall by distance group
            foot_by_dist = {}
            for label in ordered_cats(df["Distance_from_Campus"], "Distance_from_Campus"):
                sub = df[df["Distance_from_Campus"] == label]
                if len(sub):
                    foot_by_dist[label] = round(float(sub["visits_pw"].mean()),2)

            # Weekly footfall pattern — ILLUSTRATIVE. The survey never asked which days
            # students visit, so no day-of-week result can come from the data. What the data
            # does give is how many visits each respondent makes per week; those visits are
            # spread over the week using the assumed profiles in DAY_PROFILE_ASSUMPTIONS
            # (frequent visitors come every day, occasional ones cluster midweek).
            # Each respondent is weighted by their visits, so the week adds up to the
            # same total as the daily projection.
            days = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
            _day_visits = [0.0]*7
            for _v in df["visits_pw"].dropna():
                _prof = next(p for thr, p in DAY_PROFILE_ASSUMPTIONS if _v >= thr)
                for _d in range(7):
                    _day_visits[_d] += float(_v) * _prof[_d]
            _total_contrib = sum(_day_visits) or 1
            day_weights = [c / _total_contrib for c in _day_visits]
            weekly_pattern = {days[i]: round(daily_proj*7*day_weights[i],0) for i in range(7)}

            # Footfall regressor
            feat_ft = df_ml[["Age","study_h","dist_n","late_n","gs_n","spend_v",
                          "Power_Outlet_Requirement","WiFi_Importance","mem_n"]].dropna()
            y_ft    = df_ml["visits_pw"].loc[feat_ft.index]
            sc2     = StandardScaler()
            Xft     = sc2.fit_transform(feat_ft)
            rfr_ft  = RandomForestRegressor(n_estimators=80, random_state=42)
            cv_ft   = cross_val_score(rfr_ft, Xft, y_ft, cv=3, scoring="r2")
            rfr_ft.fit(Xft, y_ft)
            fi_ft   = {list(feat_ft.columns)[i]: round(float(v),4) for i,v in enumerate(rfr_ft.feature_importances_)}

            top_dist_group  = max(foot_by_dist, key=foot_by_dist.get) if foot_by_dist else "<1 km"
            peak_day        = days[day_weights.index(max(day_weights))]
            peak_day2_idx   = sorted(range(7), key=lambda i: -day_weights[i])[1]
            peak_day2       = days[peak_day2_idx]
            result["footfall"] = {
                "distribution":       vf_pct,
                "avg_visits_per_week": avg_vpw,
                "daily_proj_1000":    int(daily_proj),
                "slot_footfall":      slot_footfall,
                "foot_by_dist":       foot_by_dist,
                "weekly_pattern":     weekly_pattern,
                "weekly_pattern_basis": "assumed day profiles applied to stated visit frequency; the survey has no day-of-week question",
                "rf_r2":              round(float(cv_ft.mean()),4),
                "feature_importances": dict(sorted(fi_ft.items(), key=lambda x:-x[1])),
                "insight": (f"With a 1,000-student catchment, expect ~{int(daily_proj)} visits/day on average. "
                            f"Students in the {top_dist_group} band visit most frequently. "
                            "The weekly chart is illustrative only: the survey did not ask which days students "
                            f"visit, so it spreads their stated visits using assumed weekday profiles (peaking {peak_day} and {peak_day2}).")
            }

        except Exception as _sec_ex:
            _sec_err['footfall'] = str(_sec_ex)[:300]

        try:
            # ═══════════════════════════════════════════════════════════════
            # 3. REVENUE / SPEND FORECAST
            # ═══════════════════════════════════════════════════════════════
            spend_vc   = df["Avg_Spend_per_visit"].value_counts()
            spend_pct  = {k: round(v/N*100,1) for k,v in spend_vc.items()}
            avg_spend  = round(float(df["spend_v"].mean()),1)

            # Daily revenue projection
            spend_per_visit = (float((df["visits_pw"]*df["spend_v"]).sum() / df["visits_pw"].sum())
                               if df["visits_pw"].sum() > 0 else avg_spend)
            daily_rev  = round(daily_proj * spend_per_visit, 0)
            weekly_rev = round(daily_rev * 7, 0)
            monthly_rev= round(daily_rev * 30, 0)

            # Revenue by time slot per day (1000 catchment)
            rev_by_slot_day = {t: round(slot_footfall[t]*spend_by_time.get(t,300),0) for t in time_order}

            # Spend regressor
            feat_sp = df_ml[["Age","study_h","dist_n","gs_n","time_n","mem_n",
                          "Power_Outlet_Requirement","WiFi_Importance"]].dropna()
            y_sp    = df_ml["spend_v"].loc[feat_sp.index]
            sc3     = StandardScaler()
            Xsp     = sc3.fit_transform(feat_sp)
            rfr_sp  = RandomForestRegressor(n_estimators=80, random_state=42)
            cv_sp   = cross_val_score(rfr_sp, Xsp, y_sp, cv=3, scoring="r2")
            rfr_sp.fit(Xsp, y_sp)
            fi_sp   = {list(feat_sp.columns)[i]: round(float(v),4) for i,v in enumerate(rfr_sp.feature_importances_)}

            # Spend by university
            sbu = df.groupby("University")["spend_v"].agg(["mean","count"]).round(1)
            spend_by_univ = {str(u): {"mean":round(float(r["mean"]),1),"n":int(r["count"])}
                             for u,r in sbu.iterrows()}

            # Cumulative revenue projection (months 1-12)
            monthly_growth = 0.05  # 5% monthly growth assumption
            revenue_projection = [round(monthly_rev*(1+monthly_growth)**m,0) for m in range(12)]

            top_spend_slot = max(spend_by_time, key=spend_by_time.get)
            result["revenue"] = {
                "distribution":     spend_pct,
                "avg_spend":        avg_spend,
                "spend_per_visit":  round(spend_per_visit, 1),
                "daily_revenue":    int(daily_rev),
                "weekly_revenue":   int(weekly_rev),
                "monthly_revenue":  int(monthly_rev),
                "rev_by_slot":      rev_by_slot_day,
                "spend_by_univ":    spend_by_univ,
                "monthly_projection":revenue_projection,
                "rf_r2":            round(float(cv_sp.mean()),4),
                "feature_importances": dict(sorted(fi_sp.items(), key=lambda x:-x[1])),
                "top_spend_slot":   top_spend_slot,
                "insight": (f"The average respondent spends ₹{avg_spend:.0f} a visit; weighted by how often each one visits, "
                            f"the average visit is worth ₹{spend_per_visit:.0f}. "
                            "These figures are demand potential before any seating limit (see What-If for capacity). "
                            f"{top_spend_slot} slot generates highest avg spend (₹{spend_by_time.get(top_spend_slot,0)}). "
                            f"Projected monthly revenue (1,000-student catchment): ₹{int(monthly_rev):,}.")
            }

        except Exception as _sec_ex:
            _sec_err['revenue'] = str(_sec_ex)[:300]

        try:
            # ═══════════════════════════════════════════════════════════════
            # 4. SEATING DEMAND FORECAST
            # ═══════════════════════════════════════════════════════════════
            seat_vc  = df["Preferred_Seating_Type"].value_counts()
            seat_pct = {k: round(v/N*100,1) for k,v in seat_vc.items()}

            # Seating demand by time slot
            seat_by_time = {}
            for t in time_order:
                sub = df[df["Preferred_Time"]==t]["Preferred_Seating_Type"].value_counts()
                seat_by_time[t] = {k: round(v/sub.sum()*100,1) for k,v in sub.items()} if sub.sum() else {}

            # Seating demand by group study frequency
            seat_by_gs = {}
            for gs in ordered_cats(df["Group_Study_Frequency"], "Group_Study_Frequency"):
                sub = df[df["Group_Study_Frequency"]==gs]["Preferred_Seating_Type"].value_counts()
                if sub.sum(): seat_by_gs[gs] = {k: round(v/sub.sum()*100,1) for k,v in sub.head(3).items()}

            # Seating classifier
            feat_se = df_ml[["Age","study_h","gs_n","time_n","dist_n","late_n",
                          "Power_Outlet_Requirement","WiFi_Importance"]].dropna()
            y_se    = df_ml["Preferred_Seating_Type"].loc[feat_se.index]
            le2     = LabelEncoder()
            y_se_enc= le2.fit_transform(y_se)
            sc4     = StandardScaler()
            Xse     = sc4.fit_transform(feat_se)
            rf_se   = RandomForestClassifier(n_estimators=100, max_depth=6, random_state=42)
            cv_se   = cross_val_score(rf_se, Xse, y_se_enc, cv=3, scoring="accuracy")
            rf_se.fit(Xse, y_se_enc)
            fi_se   = {list(feat_se.columns)[i]: round(float(v),4) for i,v in enumerate(rf_se.feature_importances_)}

            top_seat = max(seat_pct, key=seat_pct.get)
            top_seat_pct = seat_pct[top_seat]
            # Compute top seating type for Evening/Late Night vs Morning dynamically
            _evening_seats  = seat_by_time.get("Evening", {})
            _latenight_seats= seat_by_time.get("Late Night", {})
            _morning_seats  = seat_by_time.get("Morning", {})
            _eve_top  = max(_evening_seats,   key=_evening_seats.get)   if _evening_seats   else "Group Table"
            _ln_top   = max(_latenight_seats, key=_latenight_seats.get) if _latenight_seats else "Group Table"
            _morn_top = max(_morning_seats,   key=_morning_seats.get)   if _morning_seats   else "Individual Desk"
            _eve_ln_desc = (_eve_top if _eve_top == _ln_top
                            else f"{_eve_top} (Evening) and {_ln_top} (Late Night)")
            n_seat_classes  = len(seat_pct)
            # v41: the fair baseline is "always guess the most common seat", not 1/number of seats
            seat_baseline   = round(majority_share(y_se) * 100, 1)
            result["seating"] = {
                "overall_pct":       seat_pct,
                "by_time_slot":      seat_by_time,
                "by_group_freq":     seat_by_gs,
                "rf_accuracy":       round(float(cv_se.mean()),4),
                "n_seat_classes":    n_seat_classes,
                "seat_baseline":     seat_baseline,
                "baseline":          round(seat_baseline/100, 4),
                "feature_importances": dict(sorted(fi_se.items(), key=lambda x:-x[1])),
                "insight": (f"{top_seat} is most demanded ({top_seat_pct:.0f}%). Seating demand shifts by time — "
                            f"Evening/Late Night visitors favour {_eve_ln_desc}. "
                            f"Morning visitors most often prefer {_morn_top}. "
                            f"Model accuracy {cv_se.mean()*100:.1f}% vs {seat_baseline:.1f}% from always guessing {top_seat}.")
            }

        except Exception as _sec_ex:
            _sec_err['seating'] = str(_sec_ex)[:300]

        try:
            # ═══════════════════════════════════════════════════════════════
            # 5. MEMBERSHIP WILLINGNESS — improved with feature engineering
            # ═══════════════════════════════════════════════════════════════
            feat_mem = df_ml[["Age","visits_pw","spend_v","study_h","dist_n","gs_n",
                           "Power_Outlet_Requirement","WiFi_Importance","Productivity_Tools_Interest",
                           "late_n","time_n"]].dropna()
            # Interaction features (feature engineering)
            feat_mem = feat_mem.copy()
            feat_mem["spend_x_visit"] = feat_mem["spend_v"]*feat_mem["visits_pw"]
            feat_mem["dist_x_late"]   = feat_mem["dist_n"]*feat_mem["late_n"]
            feat_mem["study_x_visit"] = feat_mem["study_h"]*feat_mem["visits_pw"]
            y_mem = df_ml["mem_n"].loc[feat_mem.index]
            sc5 = StandardScaler()
            Xm  = sc5.fit_transform(feat_mem)

            models_mem = {
                "Logistic Regression (engineered)": LogisticRegression(max_iter=2000,C=0.5,random_state=42),
                "Random Forest (engineered)":       RandomForestClassifier(n_estimators=100,max_depth=6,random_state=42),
                "Gradient Boosting (engineered)":   GradientBoostingClassifier(n_estimators=80,learning_rate=0.1,max_depth=4,random_state=42),
            }
            mem_scores = {}
            for name, mdl in models_mem.items():
                cv = cross_val_score(mdl, Xm, y_mem, cv=3, scoring="accuracy")
                mem_scores[name] = {"mean":round(float(cv.mean()),4),"std":round(float(cv.std()),4)}

            from sklearn.base import clone as _clone
            from sklearn.metrics import accuracy_score as _acc, precision_recall_fscore_support as _prfs
            best_name = max(mem_scores, key=lambda k: mem_scores[k]["mean"])
            X_tr,X_te,y_tr,y_te = train_test_split(Xm,y_mem,test_size=0.25,random_state=42,stratify=y_mem)
            best_mem = _clone(models_mem[best_name]); best_mem.fit(X_tr,y_tr); y_pred=best_mem.predict(X_te)
            cm = confusion_matrix(y_te,y_pred,labels=[0,1]).tolist()
            _p,_r,_f,_ = _prfs(y_te,y_pred,labels=[0,1],zero_division=0)
            cr = {"0":{"precision":_p[0],"recall":_r[0],"f1-score":_f[0]},
                  "1":{"precision":_p[1],"recall":_r[1],"f1-score":_f[1]},
                  "accuracy": float(_acc(y_te,y_pred))}
            _rf_fi = RandomForestClassifier(n_estimators=100,max_depth=6,random_state=42).fit(Xm, y_mem)
            fi_mem = {list(feat_mem.columns)[i]:round(float(v),4) for i,v in enumerate(_rf_fi.feature_importances_)}
            mem_baseline = majority_share(y_mem)

            # Membership uptake by spend tier
            mem_by_spend = {}
            for tier in ordered_cats(df["Avg_Spend_per_visit"], "Avg_Spend_per_visit"):
                sub = df[df["Avg_Spend_per_visit"]==tier]["mem_n"]
                if len(sub): mem_by_spend[tier] = round(float(sub.mean()*100),1)

            # Membership uptake by distance
            mem_by_dist = {}
            for dist_label in ordered_cats(df["Distance_from_Campus"], "Distance_from_Campus"):
                sub = df[df["Distance_from_Campus"]==dist_label]["mem_n"]
                if len(sub): mem_by_dist[dist_label] = round(float(sub.mean()*100),1)

            mem_overall_rate = round(float(df["mem_n"].mean()*100),1)
            best_mem_acc = round(max(v["mean"] for v in mem_scores.values())*100, 1) if mem_scores else round(cr["accuracy"]*100,1)
            # v41: the wording is written from the results, not fixed in advance
            _lift = best_mem_acc - mem_baseline*100
            _strength = "a strong" if _lift >= 15 else "a moderate" if _lift >= 5 else "almost no"
            _mem_insight = (f"Membership willingness: {mem_overall_rate:.0f}% willing vs {100-mem_overall_rate:.0f}% not. "
                            f"The best model ({best_name.replace(' (engineered)','')}) is right {best_mem_acc:.1f}% of the time, "
                            f"against {mem_baseline*100:.1f}% from always guessing the larger group — {_strength} improvement.")
            if len(mem_by_spend) >= 2:
                _lo = min(mem_by_spend, key=mem_by_spend.get); _hi = max(mem_by_spend, key=mem_by_spend.get)
                _mem_insight += (f" Across spend tiers willingness runs from {mem_by_spend[_lo]:.0f}% ({_lo}) "
                                 f"to {mem_by_spend[_hi]:.0f}% ({_hi}).")
            result["membership"] = {
                "model_scores":      mem_scores,
                "confusion_matrix":  cm,
                "labels":            ["No Membership","Membership"],
                "precision_0":  round(cr["0"]["precision"],3),
                "recall_0":     round(cr["0"]["recall"],3),
                "f1_0":         round(cr["0"]["f1-score"],3),
                "precision_1":  round(cr["1"]["precision"],3),
                "recall_1":     round(cr["1"]["recall"],3),
                "f1_1":         round(cr["1"]["f1-score"],3),
                "accuracy":     round(cr["accuracy"],3),
                "feature_importances": dict(sorted(fi_mem.items(), key=lambda x:-x[1])),
                "by_spend_tier": mem_by_spend,
                "by_distance":   mem_by_dist,
                "overall_rate":  mem_overall_rate,
                "best_model":       best_name.replace(" (engineered)", ""),
                "best_cv_accuracy": round(best_mem_acc/100, 4),
                "baseline":         round(mem_baseline, 4),
                "insight": _mem_insight,
            }

        except Exception as _sec_ex:
            _sec_err['membership'] = str(_sec_ex)[:300]

        try:
            # ═══════════════════════════════════════════════════════════════
            # 6. MODEL PERFORMANCE SUMMARY — all algorithms, all tasks
            # ═══════════════════════════════════════════════════════════════
            model_summary = {"classification": [], "regression": []}

            # ── Common feature set for classification tasks ───────────────
            _clf_feat = df_ml[["Age","visits_pw","spend_v","study_h","dist_n","gs_n",
                                "late_n","time_n","Power_Outlet_Requirement",
                                "WiFi_Importance","Productivity_Tools_Interest"]].dropna()
            _clf_y_mem = df_ml["mem_n"].loc[_clf_feat.index]
            _sc_clf = StandardScaler()
            _Xclf = _sc_clf.fit_transform(_clf_feat)

            _clf_models = [
                ("Logistic Regression",  LogisticRegression(max_iter=500, C=1.0, random_state=42)),
                ("Random Forest",        RandomForestClassifier(n_estimators=100, max_depth=6, random_state=42)),
                ("Gradient Boosting",    GradientBoostingClassifier(n_estimators=80, max_depth=4, random_state=42)),
                ("Decision Tree",        DecisionTreeClassifier(max_depth=6, random_state=42)),
                ("K-Nearest Neighbours", KNeighborsClassifier(n_neighbors=7)),
                ("Naïve Bayes",          GaussianNB()),
                ("SVM (RBF kernel)",     SVC(kernel="rbf", C=1.0, random_state=42)),
            ]
            for name, mdl in _clf_models:
                try:
                    cv_scores = cross_val_score(mdl, _Xclf, _clf_y_mem, cv=3, scoring="accuracy")
                    model_summary["classification"].append({
                        "model": name,
                        "task": "Membership Prediction",
                        "metric": "Accuracy",
                        "score": round(float(cv_scores.mean()) * 100, 1),
                        "std":   round(float(cv_scores.std())  * 100, 1),
                    })
                except Exception:
                    pass

            # ── Seating classifier comparison ─────────────────────────────
            _seat_feat = df_ml[["Age","study_h","gs_n","time_n","dist_n","late_n",
                                 "Power_Outlet_Requirement","WiFi_Importance"]].dropna()
            _seat_y = df_ml["Preferred_Seating_Type"].loc[_seat_feat.index]
            _le_seat = LabelEncoder()
            _seat_y_enc = _le_seat.fit_transform(_seat_y)
            _sc_seat = StandardScaler()
            _Xseat = _sc_seat.fit_transform(_seat_feat)

            _seat_clf_models = [
                ("Random Forest",        RandomForestClassifier(n_estimators=100, max_depth=6, random_state=42)),
                ("Gradient Boosting",    GradientBoostingClassifier(n_estimators=80, max_depth=4, random_state=42)),
                ("Decision Tree",        DecisionTreeClassifier(max_depth=6, random_state=42)),
                ("K-Nearest Neighbours", KNeighborsClassifier(n_neighbors=7)),
                ("Naïve Bayes",          GaussianNB()),
                ("Logistic Regression",  LogisticRegression(max_iter=500, C=1.0, random_state=42)),
            ]
            for name, mdl in _seat_clf_models:
                try:
                    cv_scores = cross_val_score(mdl, _Xseat, _seat_y_enc, cv=3, scoring="accuracy")
                    model_summary["classification"].append({
                        "model": name,
                        "task": "Seating Type Prediction",
                        "metric": "Accuracy",
                        "score": round(float(cv_scores.mean()) * 100, 1),
                        "std":   round(float(cv_scores.std())  * 100, 1),
                    })
                except Exception:
                    pass

            # ── Regression comparison — Visit Frequency ───────────────────
            _reg_feat = df_ml[["Age","study_h","dist_n","late_n","gs_n","spend_v",
                                "Power_Outlet_Requirement","WiFi_Importance","mem_n"]].dropna()
            _reg_y_vf = df_ml["visits_pw"].loc[_reg_feat.index]
            _reg_y_sp = df_ml["spend_v"].loc[_reg_feat.index]
            _sc_reg = StandardScaler()
            _Xreg = _sc_reg.fit_transform(_reg_feat)
            # v41: the spend models must not see spend itself as an input (target leakage)
            _Xreg_sp = StandardScaler().fit_transform(_reg_feat.drop(columns=["spend_v"]))

            _reg_models = [
                ("Random Forest Regressor", RandomForestRegressor(n_estimators=80, random_state=42)),
                ("Ridge Regression",        Ridge(alpha=1.0)),
                ("Decision Tree Regressor", DecisionTreeRegressor(max_depth=6, random_state=42)),
                ("K-NN Regressor",          KNeighborsRegressor(n_neighbors=7)),
                ("SVR (RBF kernel)",        SVR(kernel="rbf", C=1.0)),
            ]
            for name, mdl in _reg_models:
                try:
                    cv_r2_vf = cross_val_score(mdl, _Xreg, _reg_y_vf, cv=3, scoring="r2")
                    model_summary["regression"].append({
                        "model": name,
                        "task": "Visit Frequency",
                        "metric": "R²",
                        "score": round(float(cv_r2_vf.mean()), 3),
                        "std":   round(float(cv_r2_vf.std()),  3),
                    })
                except Exception:
                    pass
                try:
                    cv_r2_sp = cross_val_score(mdl, _Xreg_sp, _reg_y_sp, cv=3, scoring="r2")
                    model_summary["regression"].append({
                        "model": name,
                        "task": "Avg Spend per Visit",
                        "metric": "R²",
                        "score": round(float(cv_r2_sp.mean()), 3),
                        "std":   round(float(cv_r2_sp.std()),  3),
                    })
                except Exception:
                    pass

            result["model_summary"] = model_summary

        except Exception as _sec_ex:
            _sec_err['model_summary'] = str(_sec_ex)[:300]

        result["_n"] = int(N); result["_n_full"] = int(N_full); result["_section_errors"] = _sec_err
        return _sanitise(result)
    except _TooFew:
        raise
    except Exception as _pred_ex:
        import traceback
        return {"error": str(_pred_ex), "traceback": traceback.format_exc()[-1200:]}





class _TooFew(Exception):
    def __init__(self, n):
        super().__init__(n); self.n = n

_LEVER_KEYS = ("time", "spend", "seat")
_SECTION_NAMES = {"peak_hours": "Peak hours", "footfall": "Footfall and traffic", "revenue": "Revenue",
                  "seating": "Seating demand", "membership": "Membership willingness", "model_summary": "Model summary"}

def _full_predictive():
    """Result for every cleaned response (cached)."""
    if "predictive_cache" in _store:
        return _store["predictive_cache"]
    res = _predictive_compute({})
    if "error" in res:
        return res
    if res.get("_section_errors"):
        k0 = next(iter(res["_section_errors"]))
        return {"error": f"{_SECTION_NAMES.get(k0, k0)}: {res['_section_errors'][k0]}"}
    res["lever_notes"] = {}
    res["lever_options"] = {
        "time":  list((res.get("peak_hours") or {}).get("distribution", {}).keys()),
        "spend": list((res.get("revenue") or {}).get("distribution", {}).keys()),
        "seat":  list((res.get("seating") or {}).get("overall_pct", {}).keys()),
    }
    _store["predictive_cache"] = res
    _gen = time.time(); _store["lever_gen"] = _gen
    _store["lever_progress"] = {"state": "ondemand"}   # browser edition: no threads; drill-downs run on request
    return res

import re as _re, zlib as _zlib, itertools as _it
_MODEL_FIELDS = {
    "peak_hours": ["feature_importances", "rf_accuracy", "rf_std", "baseline"],
    "footfall":   ["feature_importances", "rf_r2"],
    "revenue":    ["feature_importances", "rf_r2"],
    "seating":    ["feature_importances", "rf_accuracy", "baseline", "seat_baseline", "n_seat_classes"],
    "membership": ["model_scores", "confusion_matrix", "labels", "precision_0", "recall_0", "f1_0",
                   "precision_1", "recall_1", "f1_1", "accuracy", "feature_importances",
                   "best_model", "best_cv_accuracy", "baseline"],
}
_MODEL_SENTENCE = {"peak_hours": r"The model predicts.*?slot\.", "seating": r"Model accuracy .*$",
                   "membership": r"The best model .*?improvement\."}
_MODEL_NOTE = ("The model is trained on all responses. Drilling down changes which students the counts, shares and "
               "averages describe, not the model, so this result stays the same.")

def _lever_merge(res, full, ex):
    notes = {}
    for sec, msg in (res.get("_section_errors") or {}).items():
        res[sec] = full.get(sec)
        notes["sec:" + sec] = (f"{_SECTION_NAMES.get(sec, sec)} could not be worked out for this drill-down "
                               f"({msg}). It shows the result for all responses instead.")
    for sec, fields in _MODEL_FIELDS.items():
        r, f = res.get(sec), full.get(sec)
        if not isinstance(r, dict) or not isinstance(f, dict) or ("sec:" + sec) in notes: continue
        for k in fields:
            if k in f: r[k] = f[k]
        pat = _MODEL_SENTENCE.get(sec)
        if pat and isinstance(r.get("insight"), str) and isinstance(f.get("insight"), str):
            m = _re.search(pat, f["insight"])
            if m: r["insight"] = _re.sub(pat, m.group(0).rstrip(".") + " (model trained on all responses).", r["insight"])
    res["model_summary"] = full.get("model_summary")
    for cid in ("ch-peak-fi", "ch-mem-models", "ch-mem-fi", "mem-conf-matrix"):
        notes[cid] = _MODEL_NOTE
    res["lever_notes"] = notes
    res["lever_options"] = full.get("lever_options", {})
    res["levers"] = {"excluded": ex, "n": res.get("_n"), "n_full": res.get("_n_full")}
    res.pop("_section_errors", None)
    return res

def _lever_key(ex):
    return json.dumps({k: sorted(v) for k, v in ex.items() if v}, sort_keys=True)

def _lever_one(ex, full):
    """Result for one lever combination (fast mode). Returns (status, payload dict)."""
    try:
        res = _predictive_compute(ex, fast=True)
    except _TooFew as t:
        return 400, {"error": f"Only {t.n} students are in this drill-down, too few to show reliably "
                              f"(at least 30 are needed). Tick more options.", "too_few": True}
    if "error" in res:
        return 500, res
    return 200, _lever_merge(res, full, ex)

def _lever_combos(opts):
    """Every way of leaving options out, keeping at least one option per lever."""
    per = []
    for k in _LEVER_KEYS:
        o = list(opts.get(k, []))
        subs = [list(c) for r in range(0, len(o)) for c in _it.combinations(o, r)]
        per.append(subs)
    for t, s, se in _it.product(*per):
        ex = {k: v for k, v in zip(_LEVER_KEYS, (t, s, se)) if v}
        if ex: yield ex

_LEVER_MAX = 20000      # beyond this many combinations, the rest are worked out when first asked for
def _lever_precompute(gen):
    """Background job: prepare every lever combination once, right after the first run."""
    full = _store.get("predictive_cache")
    if not full: return
    combos = list(_it.islice(_lever_combos(full.get("lever_options", {})), _LEVER_MAX))
    prog = {"done": 0, "total": len(combos), "state": "running", "started": time.time()}
    _store["lever_progress"] = prog
    cache = _store.setdefault("predictive_cache_map", {})
    for ex in combos:
        if _store.get("lever_gen") != gen: return            # a new dataset was loaded: stop
        key = _lever_key(ex)
        if key not in cache:
            try:
                st, payload = _lever_one(ex, full)
                cache[key] = (st, _zlib.compress(json.dumps(payload).encode("utf-8"), 6))
            except Exception:
                pass
        prog["done"] += 1
    prog["state"] = "done"; prog["seconds"] = round(time.time() - prog["started"], 1)

@app.route("/lever_status")
def lever_status():
    return jsonify(_store.get("lever_progress") or {"state": "idle"})

@app.route("/predictive_data")
def predictive_data():
    import warnings; warnings.filterwarnings("ignore")
    if "clean_bytes" not in _store:
        return jsonify({"error": "Run cleaning first"}), 400
    if not SKLEARN_OK:
        return jsonify({"error": "scikit-learn not installed"}), 500
    try:
        raw = json.loads(request.args.get("exclude") or "{}")
        if not isinstance(raw, dict): raw = {}
    except Exception:
        raw = {}
    full = _full_predictive()
    if "error" in full:
        return jsonify(full), 500
    opts = full.get("lever_options", {})
    ex = {}
    for k in _LEVER_KEYS:
        vals = sorted({str(v).strip() for v in (raw.get(k) or [])} & set(map(str, opts.get(k, []))))
        if vals: ex[k] = vals
    if not ex:
        return jsonify(full)
    for k, vals in ex.items():
        if len(vals) >= len(opts.get(k, [])):
            return jsonify({"error": "Keep at least one option ticked in each chart.", "too_few": True}), 400  # drill-down needs one
    key = _lever_key(ex)
    cmap = _store.setdefault("predictive_cache_map", {})
    hit = cmap.get(key)
    if hit is None:                                   # not prepared yet: work it out now (fast mode)
        st, payload = _lever_one(ex, full)
        hit = (st, _zlib.compress(json.dumps(payload).encode("utf-8"), 6))
        cmap[key] = hit
    st, blob = hit
    return app.response_class(_zlib.decompress(blob), status=st, mimetype="application/json")


# ════════════════════════════════════════════════════════════
#  /predict_membership  — live single-student prediction
#  v41: the model is trained once and cached; the form's choices
#  come from the uploaded data; and the "factors" are computed
#  for THIS student, not copied from the overall model.
# ════════════════════════════════════════════════════════════
PRED_FIELDS = {   # form field -> (dataset column, feature name)
    "visit_freq":  ("Visit_Frequency_per_week", "visits_pw"),
    "spend":       ("Avg_Spend_per_visit",      "spend_v"),
    "study_hours": ("Study_Hours",              "study_h"),
    "distance":    ("Distance_from_Campus",     "dist_n"),
    "group_study": ("Group_Study_Frequency",    "gs_n"),
    "pref_time":   ("Preferred_Time",           "time_n"),
}
PRED_BASE = ["Age","visits_pw","spend_v","study_h","dist_n","gs_n",
             "Power_Outlet_Requirement","WiFi_Importance","Productivity_Tools_Interest",
             "late_n","time_n"]
PRED_FACTOR_LABELS = {
    "spend_v":"Avg Spend per Visit","visits_pw":"Visit Frequency",
    "study_h":"Study Hours","dist_n":"Distance from Campus",
    "gs_n":"Group Study Frequency","late_n":"Late Night Study",
    "time_n":"Preferred Time","Age":"Age",
    "Power_Outlet_Requirement":"Power Outlet Need",
    "WiFi_Importance":"WiFi Importance",
    "Productivity_Tools_Interest":"Productivity Tools",
}

def _pred_add_interactions(X):
    X = X.copy()
    X["spend_x_visit"] = X["spend_v"] * X["visits_pw"]
    X["dist_x_late"]   = X["dist_n"]  * X["late_n"]
    X["study_x_visit"] = X["study_h"] * X["visits_pw"]
    return X

def _get_membership_model():
    if "mem_model" in _store:
        return _store["mem_model"]
    df = pd.read_excel(io.BytesIO(_store["clean_bytes"]), sheet_name="Cleaned_Data")
    maps, options = {}, {}
    for field, (col, feat) in PRED_FIELDS.items():
        nums = num_series(df, col)
        labs = df[col].astype(str).str.strip()
        cats = ordered_cats(df[col], col)
        maps[field]    = {c: float(nums[labs == c].iloc[0]) for c in cats if (labs == c).any()}
        options[field] = cats
        df[feat] = nums
    df["late_n"] = yes_series(df["Late_Night_Study_Preference"])
    df["mem_n"]  = yes_series(df["Willingness_for_Membership"])
    train_df = df[PRED_BASE + ["mem_n"]].dropna()
    Xtrain   = _pred_add_interactions(train_df[PRED_BASE])
    ytrain   = train_df["mem_n"]
    sc  = StandardScaler()
    clf = RandomForestClassifier(n_estimators=100, max_depth=6, random_state=42)
    clf.fit(sc.fit_transform(Xtrain), ytrain)
    options["late_night"] = ordered_cats(df["Late_Night_Study_Preference"], "Late_Night_Study_Preference")
    # Which (preferred time, late-night) combinations actually occur in the cleaned data
    combos = df.groupby([df["Preferred_Time"].astype(str).str.strip(),
                         df["Late_Night_Study_Preference"].astype(str).str.strip()]).size()
    model = {
        "sc": sc, "clf": clf, "columns": list(Xtrain.columns), "maps": maps, "options": options,
        "medians": train_df[PRED_BASE].median().to_dict(),
        "age_min": int(train_df["Age"].min()), "age_max": int(train_df["Age"].max()),
        "valid_time_ln": [[t, l] for (t, l), n in combos.items() if n > 0],
        "n_train": int(len(train_df)),
    }
    _store["mem_model"] = model
    return model

@app.route("/predictor_options")
def predictor_options():
    if "clean_bytes" not in _store:
        return jsonify({"error": "Run cleaning first"}), 400
    try:
        m = _get_membership_model()
        return jsonify({"options": m["options"], "age_min": m["age_min"], "age_max": m["age_max"],
                        "valid_time_ln": m["valid_time_ln"], "n_train": m["n_train"]})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "traceback": traceback.format_exc()[-800:]}), 500

@app.route("/predict_membership", methods=["POST"])
def predict_membership():
    import warnings as _w; _w.filterwarnings("ignore")
    if "clean_bytes" not in _store:
        return jsonify({"error": "Run cleaning first"}), 400
    if not SKLEARN_OK:
        return jsonify({"error": "scikit-learn not available"}), 500
    body = request.get_json(force=True, silent=True) or {}
    try:
        m = _get_membership_model()
    except Exception as e:
        return jsonify({"error": f"Could not train the model: {e}"}), 500

    # ── Validate against what the model was actually trained on ──
    problems, warnings_out, vals = [], [], {}
    for field, (col, feat) in PRED_FIELDS.items():
        v = str(body.get(field, "")).strip()
        if v not in m["maps"][field]:
            problems.append(f"{col}: '{v}' is not a value in the uploaded data "
                            f"(choose one of: {', '.join(m['options'][field])})")
        else:
            vals[feat] = m["maps"][field][v]
    try:
        age = float(body.get("age", ""))
    except Exception:
        problems.append("Age must be a number"); age = None
    lik = {}
    for key, col in [("power_outlet","Power_Outlet_Requirement"),("wifi","WiFi_Importance"),
                     ("productivity_tools","Productivity_Tools_Interest")]:
        try:
            lik[col] = float(body.get(key, 3))
            if not (1 <= lik[col] <= 5): problems.append(f"{col} must be between 1 and 5")
        except Exception:
            problems.append(f"{col} must be a number")
    if problems:
        return jsonify({"error": "Invalid input — " + "; ".join(problems)}), 400
    if age < m["age_min"] or age > m["age_max"]:
        warnings_out.append(f"Age {age:g} is outside the ages in the data ({m['age_min']}–{m['age_max']}); "
                            "the prediction is an extrapolation.")
    late = 1.0 if str(body.get("late_night", "No")).strip().lower() == "yes" else 0.0

    base_row = dict(vals); base_row.update(lik); base_row["Age"] = age; base_row["late_n"] = late

    def prob_of(row_dict):
        X = _pred_add_interactions(pd.DataFrame([row_dict])[PRED_BASE])[m["columns"]]
        return float(m["clf"].predict_proba(m["sc"].transform(X))[0][1])

    prob = prob_of(base_row)
    verdict    = "Yes" if prob >= 0.5 else "No"
    confidence = "High" if abs(prob-0.5) > 0.3 else "Moderate" if abs(prob-0.5) > 0.15 else "Low"

    # ── Factors for THIS prediction ──────────────────────────────
    # For each input, swap in the typical (median) respondent's value and see
    # how far the probability moves. That movement is the input's effect here.
    effects = []
    for f in PRED_BASE:
        alt = dict(base_row); alt[f] = m["medians"][f]
        delta = (prob - prob_of(alt)) * 100
        effects.append((f, delta))
    effects.sort(key=lambda x: -abs(x[1]))
    factors = [{"name": PRED_FACTOR_LABELS.get(k, k), "importance": round(abs(d), 1),
                "effect_pts": round(d, 1), "direction": "raises" if d > 0 else "lowers" if d < 0 else "no effect"}
               for k, d in effects[:3]]

    return jsonify({
        "probability": round(prob, 4),
        "probability_pct": round(prob*100, 1),
        "verdict": verdict,
        "confidence": confidence,
        "top_factors": factors,
        "factor_note": "Change in probability compared with a typical respondent's answer for that question.",
        "warnings": warnings_out,
    })


# ════════════════════════════════════════════════════════════
#  FULL HTML / CSS / JS  PAGE
# ════════════════════════════════════════════════════════════
HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Study café feasibility</title>
<script>
(function(){var t=null;try{t=localStorage.getItem('sc-theme');}catch(e){}
if(t!=='light'&&t!=='dark'){t=(window.matchMedia&&window.matchMedia('(prefers-color-scheme: dark)').matches)?'dark':'light';}
document.documentElement.setAttribute('data-theme',t);
var v=null;try{v=localStorage.getItem('sc-view');}catch(e){}
document.documentElement.setAttribute('data-view',v==='deep'?'deep':'surface');})();
</script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600&family=Public+Sans:wght@400;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#F5F7F4;--panel:#FFFFFF;--card:#FFFFFF;
  --border:#DCE3DD;--border2:#C5D0C8;
  --accent:#2b6ea3;--accent2:#2f8f6a;--accent3:#d2692c;
  --warn:#cfae2a;--danger:#B3261E;--info:#c4534a;--purple:#a8508e;
  --text:#17231F;--muted:#6E7B75;--muted2:#55635D;
  --r:12px;--r2:8px;
  --c0:#2b6ea3;--c1:#2f8f6a;--c2:#d2692c;--c3:#7f9a24;
}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;min-height:100vh;overflow-x:hidden}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--panel)}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
.screen{display:none;animation:fadeIn .3s ease}
.screen.active{display:block}
#s-upload.active,#s-progress.active,#s-nav.active{display:flex}
@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}

/* ── Upload ── */
#s-upload{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;padding:40px 20px;box-sizing:border-box}
.upload-header{text-align:center;margin-bottom:40px;transition:opacity .4s ease,max-height .4s ease,margin .4s ease;overflow:hidden}
.upload-header h1{font-size:2.1rem;font-weight:800}
.upload-header h1 em{color:var(--accent);font-style:normal}
.upload-header p{color:var(--muted2);margin-top:8px;font-size:.95rem}
.drop-card{width:100%;max-width:520px;background:var(--panel);border:2px dashed var(--border2);border-radius:18px;padding:52px 28px;text-align:center;cursor:pointer;transition:border-color .2s,background .2s}
/* ── v55 design upgrade ── */
.wi-group{border:0;margin:0 0 22px;padding:0;min-width:0}
.wi-group legend{font-family:var(--f-head);font-size:15px;font-weight:600;color:var(--text);padding:0;margin:0 0 10px}
.wi-group legend{float:left;width:100%}
.wi-group legend + *{clear:both}
.wi-group + .wi-group{border-top:1px solid var(--border);padding-top:18px}
.wi-controls h3{color:var(--text)!important}
.screen button:not(.nav-item),.screen select,.tool-btn{min-height:36px}
.likert-row input[type=range]{height:32px;cursor:pointer}
label.has-val{display:flex;align-items:baseline;justify-content:space-between;gap:8px}
label.has-val .likert-val{font-variant-numeric:tabular-nums;color:var(--text);font-weight:600;min-width:auto}
label.has-val .likert-val::after{content:' / 5';color:var(--muted2);font-weight:400}
.mk-val{font-variant-numeric:tabular-nums}
@media (pointer:coarse){.screen button,.screen select,.screen input:not([type=range]),.tool-btn,#app-nav button{min-height:44px}}
@media (max-width:860px){#app-nav.more-right{-webkit-mask-image:linear-gradient(90deg,#000 82%,transparent);mask-image:linear-gradient(90deg,#000 82%,transparent)}}
/* ── v57 report level ── */
.exp-level{border:0;margin:0 0 18px;padding:0}
.exp-level legend{font-size:13px;font-weight:600;color:var(--text);margin:0 0 8px;padding:0}
.exp-choice{display:flex;gap:10px;align-items:flex-start;padding:10px 12px;border:1px solid var(--border2);border-radius:8px;margin-bottom:8px;cursor:pointer}
.exp-choice:has(input:checked){border-color:var(--accent);box-shadow:inset 0 0 0 1px var(--accent)}
.exp-choice input{margin-top:3px;accent-color:var(--accent)}
.exp-choice b{display:block;font-size:14px;color:var(--text)}
.exp-desc{display:block;font-size:12.5px;color:var(--muted2);margin-top:2px}
/* ── v56 pricing ── */
.pricing-note{max-width:80ch;color:var(--muted2);font-size:14px;line-height:1.55;margin:-4px 0 16px}
.pr-table{margin-top:14px;font-size:13px}
.pr-table td,.pr-table th{vertical-align:top}
.pr-table tr.pr-best td{font-weight:600}
.pr-sub{display:block;font-size:11.5px;color:var(--muted2);font-weight:400;margin-top:2px}
.sgrid caption{caption-side:top;text-align:left;font-size:12.5px;color:var(--muted2);padding:0 0 6px}
.seg-tabs{display:flex;flex-wrap:wrap;gap:6px;margin:4px 0 10px}
.seg-tabs .tool-btn[aria-pressed="true"]{background:var(--accent)!important;color:var(--on-accent,#fff)!important;border-color:var(--accent)!important}
.chart-card .card-explain{font-size:13.5px;line-height:1.55;color:var(--text);margin:2px 0 10px;max-width:72ch}
.sgrid th,.sgrid td{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
.sgrid th[scope=row],.sgrid thead th:first-child{text-align:left}
.sgrid td.sg-best{font-weight:700}
.sgrid td.sg-now{outline:2px solid var(--text);outline-offset:-2px}
.sgrid td.sg-p1{background:#e3f1ea}.sgrid td.sg-p2{background:#c3e2d2}.sgrid td.sg-p3{background:#9fd0b7}
.sgrid td.sg-n1{background:#f8e3e0}.sgrid td.sg-n2{background:#f0c6c0}.sgrid td.sg-n3{background:#e6a59c}
.sgrid td{color:#17231f}
html[data-theme="dark"] .sgrid td{color:#e9efea}
html[data-theme="dark"] .sgrid td.sg-p1{background:#1d3a2f}html[data-theme="dark"] .sgrid td.sg-p2{background:#24513f}html[data-theme="dark"] .sgrid td.sg-p3{background:#2d6a51}
html[data-theme="dark"] .sgrid td.sg-n1{background:#3d2522}html[data-theme="dark"] .sgrid td.sg-n2{background:#55302b}html[data-theme="dark"] .sgrid td.sg-n3{background:#6e3a33}
.sgrid-key{font-size:12.5px;color:var(--muted2);margin:8px 0 0}
.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
/* ── Bundled dataset icon (v54) ── */
.bundled-file{display:flex;flex-direction:column;align-items:center;gap:6px;margin:0 auto 22px;padding:14px 18px 12px;
  border-radius:10px;cursor:grab;user-select:none;-webkit-user-select:none;outline:none;transition:background .15s,transform .15s}
.bundled-file:hover{background:var(--panel)}
.bundled-file:focus-visible{box-shadow:0 0 0 2px var(--accent)}
.bundled-file:active{cursor:grabbing}
.bundled-file.dragging{opacity:.45}
.bundled-file .bf-icon{filter:drop-shadow(0 2px 3px rgba(0,0,0,.15))}
.bundled-file .bf-name{font-size:13px;font-weight:600;color:var(--text)}
.bundled-file .bf-hint{font-size:12px;color:var(--muted2)}
html[data-theme="dark"] .bundled-file .bf-icon path:first-child{fill:#e9efe9}
/* ── Post-reload reconnecting overlay ── */
#reconnect-overlay{display:none;flex-direction:column;align-items:center;gap:14px;
  padding:40px 28px;text-align:center;width:100%;max-width:520px}
#reconnect-overlay.visible{display:flex}
.reconnect-spinner{width:36px;height:36px;border:3px solid var(--border2);
  border-top-color:var(--accent);border-radius:50%;animation:spin .8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.reconnect-msg{font-size:.9rem;color:var(--muted2);line-height:1.6}
.reconnect-msg strong{color:var(--accent);display:block;font-size:1rem;margin-bottom:4px}
.drop-card:hover,.drop-card.over{border-color:var(--accent);background:rgba(43,110,163,.07)}
.drop-card .icon{font-size:3rem;margin-bottom:14px}
.drop-card h3{font-size:1.05rem;margin-bottom:6px}
.drop-card p{color:var(--muted2);font-size:.86rem}
.drop-card strong{color:var(--accent)}
#file-input{display:none}
.file-chosen{margin-top:14px;font-size:.85rem;color:var(--accent2);background:rgba(47,143,106,.08);border-radius:8px;padding:9px 14px;display:none}
.up-prog-wrap{margin-top:8px;display:none;flex-direction:column;gap:5px}
.up-prog-track{height:6px;border-radius:3px;background:rgba(23,35,31,.1);overflow:hidden}
.up-prog-fill{height:100%;width:0%;border-radius:3px;background:#1F5C4A;transition:width .12s ease}
.up-prog-row{display:flex;justify-content:space-between;font-size:.72rem;color:#55635D}

/* ── Progress ── */
#s-progress{flex-direction:column;align-items:center;justify-content:center;min-height:100vh;padding:40px 20px;text-align:center;box-sizing:border-box}
.prog-title{font-size:1.5rem;font-weight:700;margin-bottom:6px}
.prog-sub{color:var(--muted2);font-size:.92rem;margin-bottom:48px}
.big-prog-wrap{width:100%;max-width:520px}
.big-prog-track{height:18px;border-radius:9px;background:var(--panel);border:1px solid var(--border2);overflow:hidden}
.big-prog-fill{height:100%;border-radius:9px;background:#1F5C4A;width:0%;transition:width .18s linear;box-shadow:0 0 18px rgba(43,110,163,.5)}
.prog-pct{font-size:2.6rem;font-weight:800;background:#1F5C4A;-webkit-background-clip:text;-webkit-text-fill-color:transparent;margin:20px 0 8px}
.prog-step-label{color:var(--muted2);font-size:.88rem;min-height:22px}
.step-list{margin-top:36px;width:100%;max-width:520px;display:flex;flex-direction:column;gap:7px;text-align:left}
.step-row{display:flex;align-items:center;gap:10px;padding:10px 14px;border-radius:8px;background:var(--panel);border:1px solid var(--border);font-size:.84rem;color:var(--muted2);transition:border-color .3s,color .3s}
.step-row.done{border-color:var(--accent2);color:var(--text)}
.step-row.active{border-color:var(--accent);color:var(--text)}
.step-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;background:var(--border2);transition:background .3s}
.step-row.done .step-dot{background:var(--accent2)}
.step-row.active .step-dot{background:var(--accent);box-shadow:0 0 8px var(--accent)}
.step-tick{margin-left:auto;font-size:.8rem;opacity:0;transition:opacity .3s}
.step-row.done .step-tick{opacity:1}

/* ── Nav ── */
#s-nav{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;padding:40px 20px;box-sizing:border-box}
.nav-header{text-align:center;margin-bottom:54px}
.nav-header h1{font-size:2rem;font-weight:800}
.nav-header h1 em{color:var(--accent2);font-style:normal}
.nav-header p{color:var(--muted2);margin-top:8px;font-size:.92rem}
.nav-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;width:100%;max-width:800px}
@media(max-width:640px){.nav-grid{grid-template-columns:1fr}}
.nav-tile{background:var(--panel);border:1px solid var(--border2);border-radius:18px;padding:34px 24px;text-align:center;cursor:pointer;transition:transform .2s,border-color .2s,box-shadow .2s;position:relative;overflow:hidden}
.nav-tile::before{content:'';position:absolute;inset:0;opacity:0;transition:opacity .2s}
.nav-tile:nth-child(1)::before{background:radial-gradient(circle at 50% 0,rgba(43,110,163,.18),transparent 70%)}
.nav-tile:nth-child(2)::before{background:radial-gradient(circle at 50% 0,rgba(47,143,106,.18),transparent 70%)}
.nav-tile:nth-child(3)::before{background:radial-gradient(circle at 50% 0,rgba(210,105,44,.18),transparent 70%)}
.nav-tile:hover{transform:translateY(-4px);box-shadow:0 12px 40px rgba(0,0,0,.4)}
.nav-tile:hover::before{opacity:1}
.nav-tile:nth-child(1):hover{border-color:var(--accent)}
.nav-tile:nth-child(2):hover{border-color:var(--accent2)}
.nav-tile:nth-child(3):hover{border-color:var(--accent3)}
.nav-tile.disabled{opacity:.38;cursor:default;pointer-events:none}
/* ── 6th utility tile ── */
.nav-tile-utility{background:var(--panel);border:1px solid var(--border2);border-radius:18px;
  padding:24px 20px;display:flex;flex-direction:column;gap:10px;cursor:default}
.util-btn{display:flex;align-items:center;gap:10px;width:100%;padding:11px 14px;
  border-radius:10px;border:1px solid var(--border2);background:var(--card);
  color:var(--text);font-size:.84rem;font-weight:600;cursor:pointer;
  transition:border-color .15s,background .15s,transform .1s;text-align:left}
.util-btn:hover{border-color:var(--accent);background:rgba(43,110,163,.08);transform:translateY(-1px)}
.util-btn .ub-icon{font-size:1.2rem;flex-shrink:0}
.util-btn .ub-text{display:flex;flex-direction:column;gap:2px}
.util-btn .ub-title{font-size:.84rem;font-weight:700;color:var(--text)}
.util-btn .ub-sub{font-size:.69rem;color:var(--muted2);font-weight:400}
.util-btn.danger:hover{border-color:var(--danger);background:rgba(224,92,92,.08)}
.util-btn.success:hover{border-color:var(--accent2);background:rgba(47,143,106,.08)}
#dl-progress-wrap{display:none;margin-top:4px}
#dl-progress-bar{height:4px;border-radius:2px;background:#1F5C4A;width:0%;transition:width .2s}
#dl-progress-label{font-size:.72rem;color:var(--muted2);margin-top:4px}
/* ── Export progress dialog ── */
#export-prog-backdrop{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:950;align-items:center;justify-content:center}
#export-prog-backdrop.open{display:flex}
#export-prog-box{background:var(--card);border:1px solid var(--border2);border-radius:16px;padding:32px 28px;width:100%;max-width:460px;box-shadow:0 24px 80px rgba(23,35,31,.45)}
#export-prog-box h3{font-size:.95rem;font-weight:700;margin:0 0 18px;color:var(--text)}
#export-prog-track{height:8px;border-radius:4px;background:var(--border);overflow:hidden;margin-bottom:10px}
#export-prog-fill{height:100%;border-radius:4px;background:#1F5C4A;width:0%;transition:width .25s ease}
#export-prog-pct{font-size:1.4rem;font-weight:800;color:var(--accent);margin-bottom:4px}
#export-prog-step{font-size:.8rem;color:var(--muted2);min-height:20px;margin-bottom:16px}
#export-prog-log{font-size:.72rem;color:var(--muted2);line-height:1.7;max-height:120px;overflow-y:auto;
  border:1px solid var(--border);border-radius:6px;padding:8px 10px;background:var(--panel)}
#export-modal-backdrop{display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:900;align-items:center;justify-content:center}
#export-modal-backdrop.open{display:flex}
#export-modal{background:var(--card);border:1px solid var(--border2);border-radius:16px;padding:32px 28px;width:100%;max-width:480px;box-shadow:0 24px 80px rgba(23,35,31,.45)}
#export-modal h3{font-size:1rem;font-weight:700;margin:0 0 6px;color:var(--text)}
#export-modal p{font-size:.8rem;color:var(--muted2);margin:0 0 22px;line-height:1.6}
.export-option{display:flex;align-items:flex-start;gap:14px;padding:16px;border:1px solid var(--border2);border-radius:10px;cursor:pointer;transition:border-color .15s,background .15s;margin-bottom:10px;background:var(--panel)}
.export-option:hover{border-color:var(--accent);background:rgba(43,110,163,.07)}
.export-option .eo-icon{font-size:1.8rem;flex-shrink:0;margin-top:2px}
.export-option .eo-body{}
.export-option .eo-title{font-size:.88rem;font-weight:700;color:var(--text);margin-bottom:4px}
.export-option .eo-desc{font-size:.75rem;color:var(--muted2);line-height:1.5}
.export-option .eo-tag{display:inline-block;font-size:.62rem;font-weight:700;padding:2px 7px;border-radius:4px;margin-left:6px;text-transform:uppercase;letter-spacing:.06em}
.tag-static{background:rgba(210,105,44,.18);color:#d2692c}
.tag-live{background:rgba(47,143,106,.18);color:var(--accent2)}
#export-modal-cancel{width:100%;margin-top:6px;padding:10px;border-radius:8px;border:1px solid var(--border2);background:none;color:var(--muted2);font-size:.82rem;cursor:pointer;transition:border-color .15s,color .15s}
#export-modal-cancel:hover{border-color:var(--accent);color:var(--text)}
.nav-tile .t-icon{font-size:2.4rem;margin-bottom:14px}
.nav-tile h3{font-size:1rem;font-weight:700;margin-bottom:6px}
.nav-tile p{font-size:.8rem;color:var(--muted2);line-height:1.5}
.nav-tile .coming{position:absolute;top:12px;right:12px;font-size:.65rem;font-weight:700;letter-spacing:.07em;background:rgba(23,35,31,.07);color:var(--muted2);border-radius:20px;padding:3px 9px;text-transform:uppercase}

/* ── Shared topbar ── */
.topbar{display:flex;align-items:center;gap:14px;padding:16px 24px;border-bottom:1px solid var(--border);background:var(--panel)}
.topbar h2{font-size:1rem;font-weight:700;flex:1}
.back-btn{display:inline-flex;align-items:center;gap:7px;padding:8px 18px;border-radius:8px;font-size:.85rem;font-weight:600;background:rgba(23,35,31,.07);color:var(--text);border:1px solid var(--border2);cursor:pointer;transition:background .15s;white-space:nowrap}
.back-btn:hover{background:rgba(23,35,31,.12)}
.dl-btn{display:inline-flex;align-items:center;gap:7px;padding:8px 18px;border-radius:8px;font-size:.85rem;font-weight:600;background:var(--accent2);color:#F5F7F4;border:none;cursor:pointer;transition:opacity .15s;white-space:nowrap}
.dl-btn:hover{opacity:.85}

/* ── Dataset Summary (screen 4) ── */
#s-summary{padding:0;min-height:100vh}
.sum-body{padding:24px;max-width:1100px;margin:0 auto}
.sum-section{margin-bottom:36px}
.sum-section h3{font-size:.72rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:14px;display:flex;align-items:center;gap:8px}
.sum-section h3::after{content:'';flex:1;height:1px;background:var(--border)}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:12px}
.kpi{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:16px}
.kpi .k-label{font-size:.7rem;color:var(--muted2);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px}
.kpi .k-val{font-size:1.8rem;font-weight:800}
.kpi .k-sub{font-size:.73rem;color:var(--muted2);margin-top:2px}
.clean-table{width:100%;border-collapse:collapse;font-size:.82rem}
.clean-table thead th{background:rgba(43,110,163,.12);color:var(--accent);padding:9px 12px;text-align:left;font-weight:600;white-space:nowrap}
.clean-table tbody tr:nth-child(even){background:rgba(23,35,31,.02)}
.clean-table tbody td{padding:9px 12px;border-bottom:1px solid var(--border);vertical-align:top}
.clean-table tbody td:nth-child(2){font-size:1.05rem;font-weight:700;color:var(--warn);text-align:center;white-space:nowrap}
.rel-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:14px}
.rel-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:18px}
.rel-card .rc-title{font-size:.87rem;font-weight:700;margin-bottom:4px;display:flex;align-items:center;gap:8px}
.rel-card .rc-desc{font-size:.76rem;color:var(--muted2);line-height:1.55;margin-bottom:10px}
.rel-card .rc-val{font-size:1.5rem;font-weight:800;margin-bottom:2px}
.rc-rows{display:flex;flex-direction:column;gap:4px}
.rc-row{display:flex;justify-content:space-between;align-items:center;font-size:.78rem;padding:4px 0;border-bottom:1px solid var(--border)}
.rc-row:last-child{border-bottom:none}
.rc-row .rk{color:var(--muted2)}
.rc-row .rv{font-weight:600}
.col-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:10px}
.col-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r2);padding:12px 14px}
.col-card .cc-name{font-size:.82rem;font-weight:700;color:var(--accent);margin-bottom:7px;word-break:break-word}
.col-card .cc-row{display:flex;justify-content:space-between;font-size:.75rem;padding:3px 0;border-bottom:1px solid rgba(23,35,31,.04)}
.col-card .cc-row:last-child{border-bottom:none}
.col-card .cc-k{color:var(--muted2)}
.col-card .cc-v{font-weight:600;text-align:right;word-break:break-word;max-width:60%}
.tbl-wrap{overflow-x:auto;border-radius:var(--r2);border:1px solid var(--border)}
.pill{display:inline-block;padding:2px 8px;border-radius:20px;font-size:.7rem;font-weight:700;letter-spacing:.04em}
.pill-green{background:rgba(47,143,106,.15);color:var(--accent2)}
.pill-blue{background:rgba(43,110,163,.15);color:var(--accent)}
.pill-orange{background:rgba(210,105,44,.15);color:var(--accent3)}
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:.7rem;font-weight:700;background:rgba(43,110,163,.15);color:var(--accent);margin-left:6px}
.c-accent{color:var(--accent)}.c-green{color:var(--accent2)}.c-orange{color:var(--accent3)}
.c-warn{color:var(--warn)}.c-danger{color:var(--danger)}.c-info{color:var(--info)}
#sum-loading{position:fixed;inset:0;background:rgba(245,247,244,.85);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;z-index:200}
#sum-loading p{color:var(--muted2);font-size:.9rem}

/* ── Analytics Dashboard (screen 5) ── */
#s-analytics{padding:0;min-height:100vh}

/* ── Predictive Analytics (screen 6) ── */
#s-predictive{padding:0;min-height:100vh}
.pred-body{padding:20px;max-width:1400px;margin:0 auto}
.model-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin-bottom:20px}
.model-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:16px;text-align:center}
.model-card .mc-name{font-size:.75rem;color:var(--muted2);margin-bottom:6px;font-weight:600;text-transform:uppercase;letter-spacing:.05em}
.model-card .mc-acc{font-size:1.9rem;font-weight:800;margin-bottom:2px}
.model-card .mc-std{font-size:.7rem;color:var(--muted2)}
.model-card.best{border-color:var(--accent2);box-shadow:0 0 18px rgba(47,143,106,.15)}
.conf-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:4px;max-width:220px;margin:12px auto}
.conf-cell{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:14px 8px;border-radius:6px;font-size:.82rem}
.conf-cell .cv{font-size:1.5rem;font-weight:800;margin-bottom:2px}
.conf-cell .cl{font-size:.65rem;color:var(--muted2);text-align:center}
.fi-bar-wrap{display:flex;flex-direction:column;gap:8px;margin-top:10px}
.fi-row{display:flex;align-items:center;gap:8px;font-size:.8rem}
.fi-name{flex:0 0 130px;color:var(--muted2);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.fi-track{flex:1;height:10px;border-radius:5px;background:var(--border);overflow:hidden}
.fi-fill{height:100%;border-radius:5px;background:#1F5C4A}
.fi-val{flex:0 0 40px;text-align:right;font-weight:700;color:var(--text)}
.scorecard-table{width:100%;border-collapse:collapse;font-size:.81rem}
.scorecard-table thead th{background:rgba(43,110,163,.15);color:var(--accent);padding:9px 14px;text-align:left;font-weight:600;cursor:pointer;user-select:none}
.scorecard-table thead th:hover{background:rgba(43,110,163,.25)}
.scorecard-table tbody tr:nth-child(even){background:rgba(23,35,31,.02)}
.scorecard-table tbody tr:hover{background:rgba(43,110,163,.08)}
.scorecard-table tbody td{padding:8px 14px;border-bottom:1px solid var(--border);color:var(--muted2)}
.scorecard-table .score-cell{font-weight:700;font-family:monospace;font-size:.88rem}
.score-bar-wrap{display:flex;align-items:center;gap:8px}
.score-bar-bg{flex:1;height:6px;border-radius:3px;background:rgba(23,35,31,.07);overflow:hidden}
.score-bar-fill{height:100%;border-radius:3px;transition:width .4s ease}
.rank-badge{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-size:.7rem;font-weight:700;background:rgba(43,110,163,.2);color:var(--accent)}
.rank-badge.gold{background:rgba(207,174,42,.25);color:#cfae2a}
.rank-badge.silver{background:rgba(150,150,150,.2);color:#aaa}
.rank-badge.bronze{background:rgba(180,100,50,.2);color:#A8741A}

/* ── What-If Scenario Modelling ── */
#s-whatif{padding:0;min-height:100vh}
.wi-body{padding:20px;max-width:1200px;margin:0 auto}
.wi-controls{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:24px;margin-bottom:24px}
.wi-controls h3{font-size:.78rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted2);margin-bottom:18px}
.wi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:16px}
.wi-field{display:flex;flex-direction:column;gap:6px}
.wi-field label{font-size:.75rem;color:var(--muted2);font-weight:600;text-transform:uppercase;letter-spacing:.05em}
.wi-field input,.wi-field select{background:var(--panel);border:1px solid var(--border2);border-radius:6px;color:var(--text);padding:8px 12px;font-size:.88rem;outline:none;width:100%}
.wi-field input:focus,.wi-field select:focus{border-color:var(--accent)}
.wi-run-btn{margin-top:18px;padding:12px 32px;border-radius:8px;background:var(--accent);color:#fff;font-weight:700;font-size:.9rem;border:none;cursor:pointer;transition:opacity .15s}
.wi-run-btn:hover{opacity:.85}
.wi-run-btn:disabled{opacity:.4;cursor:default}
.wi-results{display:none}
.wi-note{background:rgba(207,174,42,.07);border:1px solid rgba(207,174,42,.3);border-radius:var(--r);padding:16px 20px;font-size:.8rem;color:var(--muted2);line-height:1.7;margin-top:20px}
.wi-note strong{color:var(--warn)}
.sens-table{width:100%;border-collapse:collapse;font-size:.82rem;margin-top:10px}
.sens-table thead th{background:rgba(43,110,163,.12);color:var(--accent);padding:8px 12px;text-align:left;font-weight:600}
.sens-table tbody tr:nth-child(even){background:rgba(23,35,31,.02)}
.sens-table tbody td{padding:8px 12px;border-bottom:1px solid var(--border)}
.sens-table .base-row{background:rgba(47,143,106,.08)!important;font-weight:700}
.dash-body{padding:20px;max-width:1400px;margin:0 auto}
.dash-section{margin-bottom:40px}
.section-title{font-size:.72rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:16px;display:flex;align-items:center;gap:8px}
.section-title::after{content:'';flex:1;height:1px;background:var(--border)}

/* chart grid layouts */
.chart-grid-2{display:grid;grid-template-columns:repeat(2,1fr);gap:18px}
.chart-grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}
.chart-grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.chart-grid-1-2{display:grid;grid-template-columns:1fr 2fr;gap:18px}
.chart-grid-2-1{display:grid;grid-template-columns:2fr 1fr;gap:18px}
@media(max-width:900px){.chart-grid-2,.chart-grid-3,.chart-grid-4,.chart-grid-1-2,.chart-grid-2-1{grid-template-columns:1fr}}

.chart-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:18px;position:relative}
.chart-card.span2{grid-column:span 2}
.chart-card .cc-title{font-size:.82rem;font-weight:700;color:var(--text);margin-bottom:4px}
.chart-card .cc-sub{font-size:.73rem;color:var(--muted2);margin-bottom:14px}
.chart-card canvas{display:block}
.chart-card .chart-wrap{position:relative}

/* 3D canvas wrapper */
.canvas3d{width:100%;height:340px;border-radius:8px;background:var(--bg)}

/* KPI mini row */
.mini-kpi-row{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px}
.mini-kpi{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:12px 16px;flex:1;min-width:120px;text-align:center}
.mini-kpi .mk-val{font-size:1.4rem;font-weight:800}
.mini-kpi .mk-lab{font-size:.7rem;color:var(--muted2);text-transform:uppercase;margin-top:3px}

/* Tab strip for ARM */
.tab-strip{display:flex;gap:8px;margin-bottom:18px;flex-wrap:wrap}
.tab-btn{padding:7px 16px;border-radius:20px;border:1px solid var(--border2);background:transparent;color:var(--muted2);font-size:.82rem;font-weight:600;cursor:pointer;transition:all .15s}
.tab-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}
.tab-btn:hover:not(.active){border-color:var(--accent);color:var(--text)}

/* ARM table */
.arm-table{width:100%;border-collapse:collapse;font-size:.8rem}
.arm-table thead th{background:rgba(43,110,163,.12);color:var(--accent);padding:9px 10px;text-align:left;font-weight:600;white-space:nowrap;cursor:pointer;user-select:none}
.arm-table thead th:hover{color:var(--text)}
.arm-table tbody tr:hover{background:rgba(23,35,31,.04)}
.arm-table tbody td{padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:middle}
.arm-table .lift-bar{height:5px;border-radius:3px;background:#1F5C4A;margin-top:3px}

/* View more btn */
.view-more-btn{display:flex;align-items:center;justify-content:center;gap:8px;margin-top:14px;padding:10px 24px;border-radius:8px;border:1px dashed var(--border2);background:transparent;color:var(--muted2);font-size:.85rem;font-weight:600;cursor:pointer;width:100%;transition:border-color .15s,color .15s}
.view-more-btn:hover{border-color:var(--accent);color:var(--accent)}

/* Tooltip */
#tooltip{position:fixed;background:rgba(20,23,34,.97);border:1px solid var(--border2);border-radius:8px;padding:10px 14px;font-size:.78rem;line-height:1.6;pointer-events:none;z-index:9999;max-width:240px;display:none;box-shadow:0 8px 32px rgba(0,0,0,.5)}
#tooltip strong{color:var(--accent2);display:block;margin-bottom:3px}

/* cluster legend */
.cluster-legend{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:12px}
.cl-item{display:flex;align-items:center;gap:6px;font-size:.78rem;color:var(--muted2)}
.cl-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}

/* Regression insight card */
.reg-insight{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:20px}
.reg-insight h4{font-size:.88rem;font-weight:700;margin-bottom:8px}
.reg-insight p{font-size:.8rem;color:var(--muted2);line-height:1.6}
.coef-bar-wrap{margin-top:12px}
.coef-row{display:flex;align-items:center;gap:10px;margin-bottom:7px;font-size:.78rem}
.coef-row .cr-name{flex:0 0 160px;color:var(--muted2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.coef-row .cr-bar{flex:1;height:8px;border-radius:4px;background:var(--border)}
.coef-row .cr-fill{height:100%;border-radius:4px}
.coef-row .cr-val{flex:0 0 50px;text-align:right;font-weight:600}

.spinner{width:16px;height:16px;border:2px solid rgba(23,35,31,.25);border-top-color:#fff;border-radius:50%;animation:spin .6s linear infinite;display:inline-block}
@keyframes spin{to{transform:rotate(360deg)}}
</style>
</head>
<body>
<nav id="app-nav" aria-label="Sections">
  <div class="an-brand">Study café<br>feasibility</div>
  <div class="an-sub">Survey analysis</div>
  <button type="button" data-screen="s-nav" onclick="goTo('s-nav')">Overview</button>
  <button type="button" data-screen="s-summary" onclick="openSummary()">Data summary</button>
  <button type="button" data-screen="s-analytics" onclick="openAnalytics()">Exploratory analysis</button>
  <button type="button" data-screen="s-predictive" onclick="openPredictive()">Predictive analytics</button>
  <button type="button" data-screen="s-predictor" onclick="openPredictor()">Membership predictor</button>
  <button type="button" data-screen="s-whatif" onclick="openWhatIf()">What-if scenarios</button>
  <div class="an-gap"></div>
  <div class="an-facts" id="an-facts"></div>
  <div class="theme-switch" role="group" aria-label="Theme">
    <button type="button" data-theme-btn="light" onclick="setTheme('light')">Light</button>
    <button type="button" data-theme-btn="dark" onclick="setTheme('dark')">Dark</button>
  </div>
  <div class="theme-switch view-switch" role="group" aria-label="Level of detail">
    <button type="button" data-view-btn="surface" onclick="setView('surface')" title="What matters for running the café">Surface</button>
    <button type="button" data-view-btn="deep" onclick="setView('deep')" title="Every analysis technique">Deep</button>
  </div>
  <div class="an-rule"></div>
  <button type="button" onclick="downloadAnalyses()">Download report</button>
  <button type="button" onclick="analyseNewDataset()">Analyse a new dataset</button>
</nav>
<div id="theme-float" role="group" aria-label="Theme">
  <button type="button" data-theme-btn="light" onclick="setTheme('light')">Light</button>
  <button type="button" data-theme-btn="dark" onclick="setTheme('dark')">Dark</button>
</div>
<div id="tooltip"></div>

<!-- S1 UPLOAD -->
<div id="s-upload" class="screen active">
  <div class="upload-header" id="upload-header">
    <h1>Study café feasibility</h1>
    <p>Upload your raw Excel survey file to begin the automated cleaning &amp; analytics pipeline.</p>
  </div>

  <!-- Shown after Analyse New Dataset reload — waits for proxy to reconnect -->
  <div id="reconnect-overlay">
    <div class="reconnect-spinner"></div>
    <div class="reconnect-msg">
      <strong>Reconnecting to server…</strong>
      Waiting for the Colab proxy to re-establish the connection.<br>
      The upload will activate automatically once the server is ready.
    </div>
  </div>

  <div id="bundled-file" class="bundled-file" draggable="true" role="button" tabindex="0"
       aria-label="Survey dataset. Drag it into the box below, or press Enter, to start the analysis" style="display:none">
    <svg class="bf-icon" viewBox="0 0 48 56" width="48" height="56" aria-hidden="true">
      <path d="M6 2h26l14 14v36a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z" fill="#fff" stroke="#1D6F42" stroke-width="2"/>
      <path d="M32 2v12a2 2 0 0 0 2 2h12" fill="none" stroke="#1D6F42" stroke-width="2"/>
      <rect x="0" y="24" width="30" height="24" rx="3" fill="#1D6F42"/>
      <path d="M8 30l14 12M22 30L8 42" stroke="#fff" stroke-width="3.2" stroke-linecap="round"/>
    </svg>
    <span class="bf-name" id="bundled-name">study_cafe_survey_dataset.xlsx</span>
    <span class="bf-hint">Drag into the box below</span>
  </div>
  <div class="drop-card" id="drop-zone" onclick="if(!window._bundledMode)document.getElementById('file-input').click()">
    <div class="icon"></div>
    <h3>Drag &amp; drop your Excel file here</h3>
    <p>or <strong>click to browse</strong> — <em>.xlsx / .xls supported</em></p>
  </div>
  <input type="file" id="file-input" accept=".xlsx,.xls"/>
  <div id="file-chosen" class="file-chosen"></div>
  <div id="up-prog-wrap" class="up-prog-wrap">
    <div class="up-prog-track"><div id="up-prog-fill" class="up-prog-fill"></div></div>
    <div class="up-prog-row">
      <span id="up-prog-phase">Reading file...</span>
      <span id="up-prog-pct">0%</span>
    </div>
  </div>
  <div id="upload-err" style="margin-top:14px;color:var(--danger);font-size:.85rem;display:none"></div>
</div>

<!-- S2 PROGRESS -->
<div id="s-progress" class="screen">
  <div class="prog-title">Running Cleaning Pipeline</div>
  <div class="prog-sub">Applying all data-cleaning &amp; reliability checks…</div>
  <div class="big-prog-wrap">
    <div class="big-prog-track"><div class="big-prog-fill" id="prog-fill"></div></div>
    <div class="prog-pct" id="prog-pct">0%</div>
    <div class="prog-step-label" id="prog-label">Initialising…</div>
  </div>
  <div class="step-list" id="step-list"></div>
</div>

<!-- S3 NAVIGATION -->
<div id="s-nav" class="screen">
  <div class="nav-header">
    <h1>Overview</h1>
    <p>The survey has been cleaned. Open a section from the menu, or start with the data summary to see what was removed and why.</p>
  </div>
  <div class="nav-grid">
    <div class="nav-tile" onclick="openSummary()">
      <div class="t-icon"></div>
      <h3>Data summary</h3>
      <p>What was removed in cleaning and why, plus reliability checks on the answers.</p>
    </div>
    <div class="nav-tile" onclick="openAnalytics()">
      <div class="t-icon"></div>
      <h3>Exploratory analysis</h3>
      <p>Who the respondents are, how their answers relate, and the customer groups in the data.</p>
    </div>
    <div class="nav-tile" onclick="openPredictive()">
      <div class="t-icon"></div>
      <h3>Predictive analytics</h3>
      <p>Forecasts of footfall, revenue, seating demand and membership, with model accuracy.</p>
    </div>
    <div class="nav-tile" onclick="openPredictor()" style="border-color:var(--accent2)">
      <div class="t-icon"></div>
      <h3>Membership predictor</h3>
      <p>Enter one student's answers to see how likely they are to join a membership.</p>
    </div>
    <div class="nav-tile" onclick="openWhatIf()" style="border-color:var(--warn)">
      <div class="t-icon"></div>
      <h3>What-if scenarios</h3>
      <p>Change seats, prices, costs and catchment to see revenue, profit and break-even.</p>
    </div>

    <!-- 6th tile: utility actions — not clickable, contains buttons -->
    <div class="nav-tile-utility">
      <button class="util-btn danger" onclick="analyseNewDataset()">
        <span class="ub-icon"></span>
        <span class="ub-text">
          <span class="ub-title">Analyse a new dataset</span>
          <span class="ub-sub">Clear everything and restart from file upload</span>
        </span>
      </button>
      <button class="util-btn success" id="dl-analyses-btn" onclick="downloadAnalyses()">
        <span class="ub-icon"></span>
        <span class="ub-text">
          <span class="ub-title">Download report</span>
          <span class="ub-sub">Export all charts &amp; tables as PDF or interactive HTML</span>
        </span>
      </button>
    </div>
  </div>
</div>

<!-- S4 DATASET SUMMARY -->
<div id="s-summary" class="screen">
  <div class="topbar">
    <button class="back-btn" onclick="goTo('s-nav')">← Back</button>
    <h2>Data summary</h2>
    <button class="dl-btn" onclick="triggerDownload()">Download cleaned data</button>
  </div>
  <div id="sum-loading" style="position:fixed;inset:0;background:rgba(245,247,244,.85);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;z-index:200">
    <div class="spinner"></div><p style="color:var(--muted2);font-size:.9rem">Loading summary…</p>
  </div>
  <div class="sum-body" id="sum-body" style="display:none">
    <div class="sum-section"><h3>Overview</h3><div class="kpi-grid" id="kpi-grid"></div></div>
    <div class="sum-section"><h3>Relevant Data for Analytics</h3><div id="relevant-info"></div></div>
    <div class="sum-section"><h3>Cleaning Techniques Applied</h3>
      <div class="tbl-wrap"><table class="clean-table"><thead><tr><th>#</th><th>Technique</th><th style="text-align:center">Rows Affected</th><th>Method / Rationale</th></tr></thead><tbody id="clean-tbody"></tbody></table></div>
    </div>
    <div class="sum-section"><h3>Data Reliability Tests</h3><div class="rel-grid" id="rel-grid"></div></div>
    <div class="sum-section"><h3>Column Profiles</h3><div class="col-grid" id="col-grid"></div></div>
    <div class="sum-section"><h3>Cleaned Data — First 5 Rows</h3><div class="tbl-wrap" id="preview-tbl"></div></div>
  </div>
</div>

<!-- S6 PREDICTIVE ANALYTICS -->
<div id="s-predictive" class="screen">
  <div class="topbar">
    <button class="back-btn" onclick="goTo('s-nav')">← Back</button>
    <h2>Predictive analytics</h2>
  </div>
  <div id="pred-loading" style="position:fixed;inset:0;background:rgba(245,247,244,.85);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;z-index:200">
    <div class="spinner"></div><p style="color:var(--muted2);font-size:.9rem">Running predictive models…</p>
  </div>
  <div class="pred-body" id="pred-body" style="display:none">

    <!-- Levers -->
    <div class="dash-section" id="pred-levers">
      <div class="lev-head">
        <div>
          <div class="section-title">Levers</div>
          <p class="lev-intro">Drill down with <strong>Select</strong>, or by clicking a slice. Unticked options fade here, and every other chart in Predictive analytics then shows only the students in the options still ticked. The predictive models are always trained on all responses, so nothing is removed from the analysis.</p>
        </div>
        <button type="button" class="tool-btn lev-reset" id="lev-reset" onclick="levReset()" disabled>Reset all charts</button>
      </div>
      <p class="lev-status" id="lev-status" role="status" aria-live="polite"></p>
      <p class="lev-prep" id="lev-prep"></p>
      <div class="lev-grid">
        <div class="chart-card lev-card">
          <div class="lev-card-head"><div><div class="cc-title">Preferred Time Slot Distribution</div><div class="cc-sub">% of all responses, by preferred time</div></div>
            <div class="menu-wrap lev-menu-wrap" data-kind="time"><button type="button" class="tool-btn" aria-haspopup="true" aria-expanded="false">Select ▾</button><div class="lev-menu tool-menu" hidden></div></div></div>
          <canvas id="ch-lev-time" height="250"></canvas>
        </div>
        <div class="chart-card lev-card">
          <div class="lev-card-head"><div><div class="cc-title">Spend Tier Distribution</div><div class="cc-sub">% of all responses, by spend per visit</div></div>
            <div class="menu-wrap lev-menu-wrap" data-kind="spend"><button type="button" class="tool-btn" aria-haspopup="true" aria-expanded="false">Select ▾</button><div class="lev-menu tool-menu" hidden></div></div></div>
          <canvas id="ch-lev-spend" height="250"></canvas>
        </div>
        <div class="chart-card lev-card">
          <div class="lev-card-head"><div><div class="cc-title">Overall Seating Preference Mix</div><div class="cc-sub">% of all responses, by preferred seat</div></div>
            <div class="menu-wrap lev-menu-wrap" data-kind="seat"><button type="button" class="tool-btn" aria-haspopup="true" aria-expanded="false">Select ▾</button><div class="lev-menu tool-menu" hidden></div></div></div>
          <canvas id="ch-lev-seat" height="250"></canvas>
        </div>
      </div>
    </div>

    <!-- A. Peak Hours -->
    <div class="dash-section">
      <div class="section-title">Peak Hours Forecast</div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Preferred Time Slot Distribution</div>
          <div class="cc-sub">% of customers preferring each study time</div>
          <canvas id="ch-time-dist" height="240"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">Revenue Projection by Time Slot</div>
          <div class="cc-sub">Estimated revenue per 100 customers per slot</div>
          <canvas id="ch-slot-rev" height="240"></canvas>
        </div>
      </div>
      <div class="chart-card" style="margin-top:18px">
        <div class="cc-title">Average Spend per Visit by Time Slot (₹)</div>
        <div class="cc-sub" id="spend-slot-sub">Average spend per visit across each time slot. Use this for time-based pricing/offers.</div>
        <canvas id="ch-spend-by-slot" height="160"></canvas>
      </div>
      <div class="chart-card" style="margin-top:18px">
        <div class="cc-title">Peak Hours Predictor — RF Feature Importances</div>
        <div class="cc-sub">What drives which time slot a student prefers? (RF 5-fold CV accuracy shown)</div>
        <canvas id="ch-peak-fi" height="180"></canvas>
      </div>
      <div id="peak-insight" class="reg-insight" style="margin-top:14px"></div>
    </div>

    <!-- B. Footfall / Traffic -->
    <div class="dash-section">
      <div class="section-title">Footfall & Traffic Forecast</div>
      <div class="mini-kpi-row" id="foot-kpis"></div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Projected Daily Footfall by Time Slot</div>
          <div class="cc-sub">Estimated visitors per slot per day (1,000-student catchment)</div>
          <canvas id="ch-slot-foot" height="240"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">Weekly traffic pattern (illustrative)</div>
          <div class="cc-sub">Assumed weekday profiles applied to stated visit frequency. The survey did not ask which days students visit.</div>
          <canvas id="ch-weekly-pattern" height="240"></canvas>
        </div>
      </div>
      <div class="chart-card" style="margin-top:18px">
        <div class="cc-title">Avg Visits per Week by Distance from Campus</div>
        <div class="cc-sub">Closer students visit more frequently — proximity drives footfall</div>
        <canvas id="ch-foot-dist" height="160"></canvas>
      </div>
      <div id="foot-insight" class="reg-insight" style="margin-top:14px"></div>
    </div>

    <!-- C. Revenue -->
    <div class="dash-section">
      <div class="section-title">Revenue Forecast</div>
      <div class="mini-kpi-row" id="rev-kpis"></div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Spend Tier Distribution</div>
          <div class="cc-sub">How much customers spend per visit</div>
          <canvas id="ch-spend-dist" height="240"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">Revenue by Time Slot per Day</div>
          <div class="cc-sub">₹ revenue per slot (1,000-student catchment)</div>
          <canvas id="ch-rev-slot" height="240"></canvas>
        </div>
      </div>
      <div class="chart-card" style="margin-top:18px">
        <div class="cc-title">12-Month Revenue Projection (₹)</div>
        <div class="cc-sub">Assumes 5% monthly growth as the café builds a regular customer base</div>
        <canvas id="ch-rev-projection" height="180"></canvas>
      </div>
      <div class="chart-card" style="margin-top:18px">
        <div class="cc-title">Average Spend by University</div>
        <div class="cc-sub">Which campus has the highest spending potential?</div>
        <canvas id="ch-spend-univ" height="160"></canvas>
      </div>
      <div id="rev-insight" class="reg-insight" style="margin-top:14px"></div>
    </div>

    <!-- D. Seating Demand -->
    <div class="dash-section">
      <div class="section-title">Seating Demand Forecast</div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Overall Seating Preference Mix</div>
          <div class="cc-sub">Recommended allocation of seating types in the café</div>
          <canvas id="ch-seat-overall" height="260"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">Seating Demand by Time Slot (%)</div>
          <div class="cc-sub">How seating preference shifts across the day</div>
          <canvas id="ch-seat-by-time" height="260"></canvas>
        </div>
      </div>
      <div id="seat-insight" class="reg-insight" style="margin-top:14px"></div>
    </div>

    <!-- E. Membership -->
    <div class="dash-section">
      <div class="section-title">Membership Willingness Prediction</div>
      <div class="mini-kpi-row" id="mem-kpis"></div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Model Accuracy Comparison (Feature-Engineered)</div>
          <div class="cc-sub">5-fold CV accuracy with interaction features added</div>
          <canvas id="ch-mem-models" height="200"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title" id="mem-conf-title">Confusion matrix — best model</div>
          <div class="cc-sub">True vs Predicted membership willingness on 25% test set</div>
          <div class="conf-grid" id="mem-conf-matrix" style="margin:20px auto"></div>
        </div>
      </div>
      <div class="chart-grid-2" style="margin-top:18px">
        <div class="chart-card">
          <div class="cc-title">Membership Rate by Spend Tier (%)</div>
          <div class="cc-sub">Higher spenders are more willing — target them first</div>
          <canvas id="ch-mem-spend" height="220"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">Membership Rate by Distance from Campus (%)</div>
          <div class="cc-sub">Closer students are more likely to commit to membership</div>
          <canvas id="ch-mem-dist" height="220"></canvas>
        </div>
      </div>
      <div class="chart-card" style="margin-top:18px">
        <div class="cc-title">Feature Importances for Membership Prediction</div>
        <div class="cc-sub">Which features (incl. engineered interactions) matter most?</div>
        <canvas id="ch-mem-fi" height="200"></canvas>
      </div>
      <div id="mem-insight" class="reg-insight" style="margin-top:14px"></div>
    </div>

    <!-- F. Model Performance Summary -->
    <div class="dash-section">
      <div class="section-title">Model Performance Summary</div>
      <div class="reg-insight" style="margin-bottom:16px">
        <p>Every classification and regression algorithm tested on this dataset, ranked by performance.
        <strong>Classification</strong> models predict membership willingness and seating type (metric: Accuracy %).
        <strong>Regression</strong> models estimate visit frequency and avg spend (metric: R² score, where 1.0 = perfect, 0 = no better than mean, negative = worse than mean).
        All scores are 3-fold cross-validated on a representative sample of the cleaned dataset.</p>
      </div>

      <!-- Full comparison table -->
      <div class="chart-card" style="margin-top:18px">
        <div class="cc-title">Complete Model Scorecard</div>
        <div class="cc-sub">All models, all tasks — ranked by score</div>
        <div class="tbl-wrap" style="margin-top:12px" id="model-scorecard-wrap"></div>
      </div>
    </div>

  </div>
</div>

<!-- S5 ANALYTICS DASHBOARD -->
<div id="s-analytics" class="screen">
  <div class="topbar">
    <button class="back-btn" onclick="goTo('s-nav')">← Back</button>
    <h2>Exploratory analysis</h2>
  </div>
  <div class="dash-body">

    <!-- KPI strip -->
    <div class="mini-kpi-row" id="kpi-strip"></div>

    <!-- 1. Summary Charts -->
    <div class="dash-section">
      <div class="section-title">Data Summary — Distributions</div>
      <div class="chart-grid-3">
        <div class="chart-card"><div class="cc-title">Age Distribution</div><div class="cc-sub">Histogram of respondent ages</div><div class="chart-wrap" style="height:220px"><canvas id="ch-age"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Preferred Visit Time</div><div class="cc-sub">When students prefer to study</div><div class="chart-wrap" style="height:220px"><canvas id="ch-time"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Avg Spend per Visit</div><div class="cc-sub">Spending tier breakdown</div><div class="chart-wrap" style="height:220px"><canvas id="ch-spend"></canvas></div></div>
      </div>
      <div class="chart-grid-3" style="margin-top:18px">
        <div class="chart-card"><div class="cc-title">Beverage Preference</div><div class="cc-sub">Most popular drinks</div><div class="chart-wrap" style="height:220px"><canvas id="ch-bev"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Food Preference</div><div class="cc-sub">Snack vs meal preferences</div><div class="chart-wrap" style="height:220px"><canvas id="ch-food"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Noise Tolerance</div><div class="cc-sub">Preferred study environment</div><div class="chart-wrap" style="height:220px"><canvas id="ch-noise"></canvas></div></div>
      </div>
      <div class="chart-grid-3" style="margin-top:18px">
        <div class="chart-card"><div class="cc-title">Seating Preference</div><div class="cc-sub">Preferred seat type</div><div class="chart-wrap" style="height:220px"><canvas id="ch-seat"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Visit Frequency</div><div class="cc-sub">How often they'd visit per week</div><div class="chart-wrap" style="height:220px"><canvas id="ch-freq"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Membership Willingness</div><div class="cc-sub">Membership split &amp; price tolerance</div><div class="chart-wrap" style="height:220px"><canvas id="ch-membership"></canvas></div></div>
      </div>
      <!-- 3D bar chart: University × Spend -->
      <div style="margin-top:18px" class="chart-card">
        <div class="cc-title">Likert Scale Ratings (3D Bar) — WiFi · Power · Productivity</div>
        <div class="cc-sub">Average scores per rating level across all three Likert items</div>
        <canvas id="ch-likert3d" height="280"></canvas>
      </div>
    </div>

    <!-- 2. EDA Deep Dive -->
    <div class="dash-section">
      <div class="section-title">EDA — Cross-Tabulations &amp; Patterns</div>
      <div class="chart-grid-2">
        <div class="chart-card"><div class="cc-title">Study Hours vs Visit Frequency</div><div class="cc-sub">Stacked bar — more study hours → more visits?</div><div class="chart-wrap" style="height:260px"><canvas id="ch-hours-freq"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Distance vs Late-Night Preference</div><div class="cc-sub">Grouped — proximity correlates with night study?</div><div class="chart-wrap" style="height:260px"><canvas id="ch-dist-night"></canvas></div></div>
      </div>
      <div class="chart-grid-2" style="margin-top:18px">
        <div class="chart-card"><div class="cc-title">Group Study Frequency vs Seating Type</div><div class="cc-sub">Do group studiers prefer group tables?</div><div class="chart-wrap" style="height:260px"><canvas id="ch-grp-seat"></canvas></div></div>
        <div class="chart-card"><div class="cc-title">Budget Distribution by University</div><div class="cc-sub">Budget band chosen by respondents at each university</div><div class="chart-wrap" style="height:260px"><canvas id="ch-univ-budget"></canvas></div></div>
      </div>
    </div>

    <!-- 3. Correlation -->
    <div class="dash-section">
      <div class="section-title">Correlation Heatmap</div>
      <div class="chart-card">
        <div class="cc-title">Pearson Correlation Matrix — Numeric &amp; Ordinal Variables</div>
        <div class="cc-sub">How strongly each pair of answers moves together, in three shades per direction. Hover a square for the exact value.</div>
        <canvas id="ch-corr" height="420"></canvas>
      </div>
    </div>

    <!-- 4. Regression -->
    <div class="dash-section">
      <div class="section-title">Regression Analysis</div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Logistic Regression — Predicting Membership Willingness</div>
          <div class="cc-sub">Feature coefficients (standardised). Positive = increases probability of Yes.</div>
          <canvas id="ch-reg-coef" height="280"></canvas>
        </div>
        <div class="reg-insight">
          <h4>Regression Insights</h4>
          <p>A logistic regression model was trained to predict <strong>Willingness for Membership</strong> using all numeric and ordinal features (5-fold cross-validated accuracy: <strong id="reg-acc-val">—</strong> — near chance, consistent with weak linear relationships in this dataset).</p>
          <p style="margin-top:10px">Key drivers:</p>
          <div class="coef-bar-wrap" id="coef-bars"></div>
          <p style="margin-top:12px;color:var(--muted2)" id="reg-top-predictor">Loading top predictor…</p>
        </div>
      </div>
      <div style="margin-top:18px" class="chart-card">
        <div class="cc-title">Predicted Membership Probability by Budget Tier (2D Line)</div>
        <div class="cc-sub">Estimated membership probability increases with stated budget willingness</div>
        <canvas id="ch-reg-line" height="200"></canvas>
      </div>
    </div>

    <!-- 5. Clustering -->
    <div class="dash-section">
      <div class="section-title">K-Means Clustering</div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Elbow Curve — Optimal K Selection</div>
          <div class="cc-sub" id="elbow-sub">Optimal K selected via elbow method (knee point of inertia curve).</div>
          <canvas id="ch-elbow" height="220"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">Cluster Profiles — Radar Chart</div>
          <div class="cc-sub">Normalised feature means per cluster</div>
          <canvas id="ch-cluster-radar" height="220"></canvas>
        </div>
      </div>
      <div class="chart-stack" style="margin-top:18px">
        <div class="chart-card">
          <div class="cc-title">2D PCA Scatter — Cluster Separation</div>
          <div class="cc-sub">PC1 against PC2, the two strongest summary scores (explained in the table below). Hover for details.</div>
          <div class="cluster-legend" id="scatter-legend-2d"></div>
          <canvas id="ch-scatter2d" height="400"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">3D cluster plot</div>
          <div class="cc-sub">PC1, PC2 and PC3 (explained in the table below). Drag to rotate to any angle, scroll to zoom; hover a point for details.</div>
          <div class="cluster-legend" id="scatter-legend-3d" style="display:none"></div>
          <canvas id="ch-scatter3d" tabindex="0" role="img" aria-label="3D cluster plot of the first three principal components. Drag or use the arrow keys to rotate, scroll or plus and minus to zoom, double-click to reset." style="display:block;width:100%;height:440px;touch-action:none;cursor:grab"></canvas>
        </div>
      </div>
      <div style="margin-top:18px;display:none" class="chart-card" id="pca-explain">
        <div class="cc-title">What PC1, PC2 and PC3 mean</div>
        <div class="cc-sub" id="pca-explain-sub"></div>
        <div class="tbl-wrap"><table class="clean-table" id="pca-explain-tbl"><thead><tr><th>Axis</th><th>Share of variation</th><th>Higher scores mean more</th><th>Lower scores mean more</th></tr></thead><tbody></tbody></table></div>
      </div>
      <div style="margin-top:18px" class="chart-card">
        <div class="cc-title">Cluster Comparison — Grouped Bar</div>
        <div class="cc-sub">Key metrics averaged per cluster segment</div>
        <canvas id="ch-cluster-bar" height="240"></canvas>
      </div>
      <div class="chart-grid-2" style="margin-top:18px">
        <div class="chart-card">
          <div class="cc-title">Age vs Membership Rate — by Cluster</div>
          <div class="cc-sub">Bubble size = cluster size. Shows how age and membership rate vary across segments.</div>
          <canvas id="ch-cluster-age-mem" height="220"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title">Age Bracket Distribution — by Cluster</div>
          <div class="cc-sub">Proportion of each age group within each cluster segment.</div>
          <canvas id="ch-cluster-age-dist" height="220"></canvas>
        </div>
      </div>
    </div>

    <!-- 6. Association Rules -->
    <div class="dash-section" id="pricing-section">
      <div class="section-title">Pricing and willingness to pay</div>
      <div class="wi-note pricing-note" id="pricing-note">These charts use what students said they would pay, not what they actually paid. People tend to say they would pay more, and join more readily, than they really do, so read them as a guide to which price bands work better, not as a forecast. Each answer is a price band, so each curve has only a few points.</div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Membership: who would still join at each monthly fee</div>
          <div class="cc-sub">Students who want a membership, by the most they said they would pay</div>
          <canvas id="ch-wtp-mem" height="280"></canvas>
          <div class="tbl-wrap"><table class="sens-table pr-table" id="tbl-wtp-mem"></table></div>
        </div>
        <div class="chart-card">
          <div class="cc-title">One visit: whose budget reaches each price</div>
          <div class="cc-sub">All students, by the most they said they would spend on one visit</div>
          <canvas id="ch-wtp-visit" height="280"></canvas>
          <div class="tbl-wrap"><table class="sens-table pr-table" id="tbl-wtp-visit"></table></div>
        </div>
      </div>
      <div class="chart-grid-2">
        <div class="chart-card">
          <div class="cc-title">Room to raise prices</div>
          <div class="cc-sub">How far each student's budget sits above what they spend now</div>
          <p class="card-explain" data-title="What this shows" id="headroom-read"></p>
          <canvas id="ch-headroom" height="260"></canvas>
        </div>
        <div class="chart-card">
          <div class="cc-title" id="seg-title">Who would pay more for a membership</div>
          <div class="cc-sub" id="seg-sub"></div>
          <div class="seg-tabs" role="group" aria-label="Customer group" id="seg-tabs"></div>
          <p class="card-explain" data-title="What this shows" id="seg-read"></p>
          <canvas id="ch-wtp-seg" height="260"></canvas>
          <div class="tbl-wrap"><table class="sens-table pr-table" id="tbl-wtp-seg"></table></div>
        </div>
      </div>
    </div>

    <div class="dash-section">
      <div class="section-title">Association rules</div>
      <p id="arm-note" style="color:var(--muted2);font-size:14px;max-width:78ch;margin:-6px 0 14px"></p>

      <!-- Tab strip for chart types -->
      <div class="tab-strip">
        <button class="tab-btn active" onclick="armTab('table')">Rule Table</button>
        <button class="tab-btn" onclick="armTab('bubble')">Bubble Chart</button>
        <button class="tab-btn" onclick="armTab('scatter')">Lift Scatter</button>
        <button class="tab-btn" onclick="armTab('bar')">Confidence Bar</button>
        <button class="tab-btn" onclick="armTab('heatmap')">Support Heatmap</button>
      </div>

      <!-- Table view -->
      <div id="arm-table-view" class="chart-card">
        <div class="cc-title">Top Association Rules</div>
        <div class="cc-sub">Showing top 10 rules by lift — hover rows for details</div>
        <div class="card-explain arm-explain" id="arm-explain" data-title="About association rules"></div>
        <div class="tbl-wrap">
          <table class="arm-table" id="arm-main-table">
            <thead><tr>
              <th onclick="sortArm('antecedent')">Antecedent ↕</th>
              <th onclick="sortArm('consequent')">Consequent ↕</th>
              <th onclick="sortArm('support')">Support ↕</th>
              <th onclick="sortArm('confidence')">Confidence ↕</th>
              <th onclick="sortArm('lift')">Lift ↕</th>
              <th onclick="sortArm('n_both')">Students ↕</th>
              <th onclick="sortArm('q_value')">Adjusted p-value ↕</th>
              <th>Significant?</th>
              <th>Lift Bar</th>
            </tr></thead>
            <tbody id="arm-tbody-main"></tbody>
          </table>
        </div>
        <button class="view-more-btn" id="arm-view-more-btn" onclick="toggleViewMore()">
          ▼ View all rules
        </button>
        <div id="arm-all-rules" style="display:none;margin-top:14px">
          <div class="tbl-wrap">
            <table class="arm-table" id="arm-all-table">
              <thead><tr>
                <th onclick="sortArmAll('antecedent')">Antecedent ↕</th>
                <th onclick="sortArmAll('consequent')">Consequent ↕</th>
                <th onclick="sortArmAll('support')">Support ↕</th>
                <th onclick="sortArmAll('confidence')">Confidence ↕</th>
                <th onclick="sortArmAll('lift')">Lift ↕</th>
                <th onclick="sortArmAll('n_both')">Students ↕</th>
                <th onclick="sortArmAll('q_value')">Adjusted p-value ↕</th>
                <th>Significant?</th>
                <th>Lift Bar</th>
              </tr></thead>
              <tbody id="arm-tbody-all"></tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Bubble view -->
      <div id="arm-bubble-view" class="chart-card" style="display:none">
        <div class="cc-title">Support vs Confidence Bubble Chart (size = Lift)</div>
        <div class="cc-sub">Each bubble = one rule. Larger bubble = higher lift. Click a bubble to see the rule explained.</div>
        <canvas id="ch-arm-bubble" height="380"></canvas>
      </div>

      <!-- Scatter view -->
      <div id="arm-scatter-view" class="chart-card" style="display:none">
        <div class="cc-title">Lift Distribution Scatter — All Rules</div>
        <div class="cc-sub">Rules plotted by confidence vs lift. Hover for details.</div>
        <div style="position:relative;height:380px;width:100%"><canvas id="ch-arm-scatter"></canvas></div>
      </div>

      <!-- Bar view -->
      <div id="arm-bar-view" class="chart-card" style="display:none">
        <div class="cc-title">Top 20 Rules — Confidence Bar Chart</div>
        <div class="cc-sub">Horizontal bars showing confidence level for each rule</div>
        <div style="position:relative;height:500px;width:100%"><canvas id="ch-arm-bar"></canvas></div>
      </div>

      <!-- Heatmap view -->
      <div id="arm-heatmap-view" class="chart-card" style="display:none">
        <div class="cc-title">Support Heatmap — Antecedent vs Consequent</div>
        <div class="cc-sub" id="arm-heatmap-sub">The strongest rules by lift. Hover a square for details.</div>
        <div style="position:relative;height:420px;width:100%"><canvas id="ch-arm-heatmap"></canvas></div>
      </div>

    </div><!-- /ARM section -->

  </div><!-- /dash-body -->
</div><!-- /s-analytics -->

<script>
const BASE = window.location.origin;
let _summaryLoaded = false;
let _analyticsBuilt = false;

// ── Post-reload reconnect: if Analyse New Dataset triggered a reload,
//    poll GET / until Flask responds before enabling file upload ──────
(function(){
  if(!sessionStorage.getItem('_reloading')) return;
  sessionStorage.removeItem('_reloading');

  // Show the reconnecting overlay and hide the drop zone
  const overlay = document.getElementById('reconnect-overlay');
  const dz      = document.getElementById('drop-zone');
  const fi      = document.getElementById('file-input');
  if(overlay) overlay.classList.add('visible');
  if(dz)      dz.style.display = 'none';
  if(fi)      fi.disabled = true;

  // Poll GET / every 1.5s until we get a 200 response
  let attempts = 0;
  const MAX_ATTEMPTS = 20; // 30 seconds total
  const poll = setInterval(async ()=>{
    attempts++;
    try{
      const r = await fetch(BASE+'/', {method:'GET', cache:'no-store'});
      if(r.ok){
        clearInterval(poll);
        // Server is reachable — restore upload UI
        if(overlay) overlay.classList.remove('visible');
        if(dz)      dz.style.display = 'block';
        if(fi)      fi.disabled = false;
      }
    }catch(e){
      // Server not yet reachable — keep polling
    }
    if(attempts >= MAX_ATTEMPTS){
      clearInterval(poll);
      // Give up polling — let user try manually
      if(overlay){
        overlay.innerHTML = '<div class="reconnect-msg">'
          + '<strong style="color:var(--danger)">⚠ Could not reconnect</strong>'
          + 'The server did not respond after 30 seconds.<br>'
          + 'Please re-run the Colab cell to restart the server.</div>';
      }
    }
  }, 1500);
})();

// ══ Pre-computed datasets ══
// RULES is now computed live by /analytics_data
// ANALYTICS is now computed live by /analytics_data
// CLUSTERS is now computed live by /analytics_data
// CORR is now computed live by /analytics_data
// v41: names are filled in from the data after the analysis loads (never assumed in advance)
const CLUSTER_NAMES = ['Cluster 1','Cluster 2','Cluster 3','Cluster 4'];
const CLUSTER_COLORS = ['#2b6ea3','#d2692c','#2f8f6a','#cfae2a','#a8508e','#6a86d8','#c4534a','#7f9a24'];

// ══ Tooltip helper ══
const TT = document.getElementById('tooltip');
function showTip(e, html){ TT.innerHTML=html; TT.style.display='block'; moveTip(e); }
function moveTip(e){ TT.style.left=(e.clientX+14)+'px'; TT.style.top=(e.clientY-10)+'px'; }
function hideTip(){ TT.style.display='none'; }

// ══ Screen nav ══
function resetUploadScreen(){
  // Restore the upload screen to its original pristine state
  const uploadScreen = document.getElementById('s-upload');
  uploadScreen.style.opacity    = '';
  uploadScreen.style.minHeight  = '';
  uploadScreen.style.padding    = '';
  uploadScreen.style.overflow   = '';
  uploadScreen.style.transition = '';

  // Restore drop zone
  const dz = document.getElementById('drop-zone');
  dz.style.display    = 'block';
  dz.style.opacity    = '1';
  dz.style.transform  = 'none';
  dz.style.transition = '';
  dz.classList.remove('over');

  // Hide progress and file-chosen
  const fc   = document.getElementById('file-chosen');
  const wrap = document.getElementById('up-prog-wrap');
  const fill = document.getElementById('up-prog-fill');
  const phase= document.getElementById('up-prog-phase');
  const upPct= document.getElementById('up-prog-pct');
  if(fc)   { fc.style.display = 'none'; fc.textContent = ''; }
  if(wrap) { wrap.style.display = 'none'; }
  if(fill) { fill.style.width = '0%'; }
  if(phase){ phase.textContent = 'Reading file...'; }
  if(upPct){ upPct.textContent = '0%'; }

  // Restore upload header
  const hdr = document.getElementById('upload-header');
  if(hdr){
    hdr.style.opacity      = '1';
    hdr.style.maxHeight    = '';
    hdr.style.marginBottom = '';
  }

  // Clear error
  const err = document.getElementById('upload-err');
  if(err){ err.style.display = 'none'; err.textContent = ''; }

  // v54: the bundled icon belongs to a fresh start only
  const _bf=document.getElementById('bundled-file'); if(_bf && !window._bundledMode) _bf.style.display='none';

  // CRITICAL: clear the file input value so the same file can be selected again
  const fi = document.getElementById('file-input');
  if(fi) fi.value = '';
}

let _navFactsDone=false;
async function _navFacts(){
  if(_navFactsDone) return;
  try{
    const r=await fetch(BASE+'/summary_data'); const d=await r.json();
    if(d.error) return;
    _navFactsDone=true;
    const el=document.getElementById('an-facts');
    if(el) el.innerHTML='<b>'+d.n_final.toLocaleString()+'</b>of '+d.n_start.toLocaleString()+' responses kept after cleaning';
  }catch(e){}
}
function goTo(id){
  document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  // v42: persistent navigation — always shows where you are
  const _noNav=(id==='s-upload'||id==='s-progress');
  document.body.classList.toggle('has-nav',!_noNav);
  document.querySelectorAll('#app-nav button[data-screen]').forEach(b=>{
    const on=b.dataset.screen===id; b.classList.toggle('on',on);
    if(on) b.setAttribute('aria-current','page'); else b.removeAttribute('aria-current');
  });
  if(!_noNav) _navFacts();
  // If navigating back to upload, reset it fully so same file can be re-uploaded
  if(id === 's-upload') resetUploadScreen();
  // Force scroll to top using multiple methods for cross-browser reliability
  window.scrollTo({top:0, behavior:'instant'});
  document.documentElement.scrollTop = 0;
  document.body.scrollTop = 0;
}

// ══ Download helper — uses fetch+Blob so the page never navigates away ══
// This means clicking Download multiple times always works correctly.
function triggerDownload(){
  const btn = event && event.target ? event.target : document.querySelector('.dl-btn');
  const origText = btn ? btn.textContent : '';
  if(btn){ btn.disabled = true; btn.textContent = 'Preparing…'; }
  fetch(BASE + '/download')
    .then(r => {
      if(!r.ok) throw new Error('Server returned ' + r.status);
      return r.blob();
    })
    .then(blob => {
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = 'study_cafe_cleaned.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      // Release the object URL after a short delay
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      if(btn){ btn.disabled = false; btn.textContent = origText; }
    })
    .catch(e => {
      alert('Download failed: ' + e.message);
      if(btn){ btn.disabled = false; btn.textContent = origText; }
    });
}

// ══ v55 small behaviours ══
// One case style everywhere: headings and labels in sentence case. Words written
// with capitals inside them (PCA, K-Means, WiFi) are left as they are.
function _sentence(t,first){
  return t.replace(/(^|[\s(\/]+)([A-Z])([a-z]+)\b/g,(m,pre,a,b,off)=>{
    if(first&&off===0&&!pre.trim()) return m;
    if(first&&!t.slice(0,off).trim()) return m;
    return pre+a.toLowerCase()+b;
  });
}
function _sentenceCase(root){
  (root||document).querySelectorAll('.cc-title,.section-title,.pred-field > label,.impact-tag,.chart-card h3,.wi-tab-btn,.tab-btn').forEach(el=>{
    if(el.dataset.sc) return; el.dataset.sc='1';
    let first=true;
    el.childNodes.forEach(n=>{ if(n.nodeType!==3||!n.textContent.trim()) return; n.textContent=_sentence(n.textContent,first); first=false; });
    if(el.dataset.deepTitle) el.dataset.deepTitle=_sentence(el.dataset.deepTitle,true);
  });
}
document.addEventListener('DOMContentLoaded',()=>{
  _sentenceCase();
  let _scT=null; new MutationObserver(()=>{ clearTimeout(_scT); _scT=setTimeout(()=>_sentenceCase(),60); }).observe(document.body,{childList:true,subtree:true});
  // a slider's value sits next to its label ("3 / 5"), not after the scale's end
  document.querySelectorAll('.likert-row').forEach(r=>{
    const v=r.querySelector('.likert-val'), lab=r.closest('.pred-field')&&r.closest('.pred-field').querySelector('label');
    const inp=r.querySelector('input[type=range]');
    if(v&&lab&&inp){ lab.classList.add('has-val'); lab.appendChild(v); inp.setAttribute('aria-label',lab.firstChild.textContent.trim()); lab.setAttribute('for',inp.id); }
  });
  // on phones the section menu scrolls sideways: fade its edge while more is hidden
  const nav=document.getElementById('app-nav');
  if(nav){ const f=()=>nav.classList.toggle('more-right',nav.scrollWidth-nav.clientWidth-nav.scrollLeft>8);
    nav.addEventListener('scroll',f,{passive:true}); addEventListener('resize',f); setTimeout(f,300); }
});

// ══ Bundled dataset (v54) ══
// Shown on a fresh start when the app ships with the survey file. After
// "Analyse a new dataset" the tab is marked and only the plain drop box shows.
window._bundledMode=false;
const _BUNDLED_MIME='application/x-sc-bundled';
async function _bundledSetup(){
  let fresh=true; try{ fresh=!sessionStorage.getItem('sc-new-dataset'); }catch(e){}
  if(!fresh) return;
  let info=null; try{ info=await (await fetch(BASE+'/bundled_info')).json(); }catch(e){}
  if(!info||!info.available) return;
  const bf=document.getElementById('bundled-file'), dzb=document.getElementById('drop-zone');
  document.getElementById('bundled-name').textContent=info.name;
  bf.style.display='flex'; window._bundledMode=true;
  dzb.querySelector('h3').textContent='Drop the survey file here to start';
  dzb.querySelector('p').innerHTML='Drag the Excel file above into this box';
  dzb.style.cursor='default';
  const hp=document.querySelector('#upload-header p');
  if(hp) hp.textContent='The survey dataset is ready. Drag it into the box to run the cleaning and analysis.';
  bf.addEventListener('dragstart',e=>{
    e.dataTransfer.setData(_BUNDLED_MIME,'1'); e.dataTransfer.setData('text/plain',info.name);
    e.dataTransfer.effectAllowed='copy'; bf.classList.add('dragging');
  });
  bf.addEventListener('dragend',()=>bf.classList.remove('dragging'));
  bf.addEventListener('keydown',e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); _loadBundled(); } });
}
async function _loadBundled(){
  const bf=document.getElementById('bundled-file');
  try{
    const r=await fetch(BASE+'/bundled_dataset'); if(!r.ok) throw new Error('HTTP '+r.status);
    const blob=await r.blob();
    const name=document.getElementById('bundled-name').textContent||'study_cafe_survey_dataset.xlsx';
    bf.style.display='none'; window._bundledMode=false;
    handleFile(new File([blob],name,{type:blob.type}));
  }catch(e){
    const err=document.getElementById('upload-err');
    err.textContent='Could not open the bundled dataset ('+e.message+'). Click the box to choose the file yourself.';
    err.style.display='block'; window._bundledMode=false;
  }
}
document.addEventListener('DOMContentLoaded',_bundledSetup);

// ══ Drag & Drop ══
const dz = document.getElementById('drop-zone');
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over')});
dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');
  if(e.dataTransfer.types && [...e.dataTransfer.types].includes(_BUNDLED_MIME)){ _loadBundled(); return; }
  const f=e.dataTransfer.files[0];if(f)handleFile(f)});
document.getElementById('file-input').addEventListener('change',e=>{if(e.target.files[0])handleFile(e.target.files[0])});

function handleFile(file){
  // Clear file input so the same file can be selected again next time
  const fi = document.getElementById('file-input');
  if(fi) fi.value = '';

  const err   = document.getElementById('upload-err');   err.style.display='none';
  const fc    = document.getElementById('file-chosen');
  const wrap  = document.getElementById('up-prog-wrap');
  const fill  = document.getElementById('up-prog-fill');
  const phase = document.getElementById('up-prog-phase');
  const upPct = document.getElementById('up-prog-pct');
  if(!file.name.toLowerCase().endsWith(".xlsx") && !file.name.toLowerCase().endsWith(".xls")){
    err.textContent='⚠ Only .xlsx/.xls supported'; err.style.display='block'; return;
  }
  // Fade out drop zone
  const dz=document.getElementById('drop-zone');
  dz.style.transition='opacity 0.4s ease, transform 0.4s ease';
  dz.style.opacity='0'; dz.style.transform='scale(0.95)';
  setTimeout(()=>{ dz.style.display='none'; },400);

  fc.textContent=''+file.name+' — '+(file.size/1024).toFixed(1)+' KB';
  fc.style.display='block';
  wrap.style.display='flex';

  function setP(p, label){
    fill.style.width=p+'%';
    upPct.textContent=Math.round(p)+'%';
    if(label) phase.textContent=label;
  }
  setP(0,'Reading file...');

  // Phase 1: read file into FormData (0→40%)
  const reader = new FileReader();
  reader.onprogress = function(e){
    if(e.lengthComputable) setP((e.loaded/e.total)*40,'Reading file...');
  };
  reader.onload = function(){
    setP(40,'Uploading...');
    const fd=new FormData(); fd.append('file',file);
    const xhr=new XMLHttpRequest();
    xhr.open('POST', BASE+'/upload');
    xhr.timeout = 60000; // 60s timeout — prevents infinite hang at 95%
    xhr.ontimeout=function(){
      dz.style.display='block'; dz.style.opacity='1'; dz.style.transform='none';
      wrap.style.display='none'; fc.style.display='none';
      err.textContent='⚠ Upload timed out — the server took too long to respond. Please try again.';
      err.style.display='block';
      handleFile._retried=false;
    };
    xhr.upload.onprogress=function(e){
      if(e.lengthComputable) setP(40+(e.loaded/e.total)*55,'Uploading...');
    };
    xhr.onload=function(){
      // Colab's proxy sometimes returns an empty body on the first request
      // after a page reload. Detect this and auto-retry the upload once.
      if(!xhr.responseText || xhr.responseText.trim()===''){
        if(!handleFile._retried){
          handleFile._retried=true;
          setP(0,'Retrying…'); phase.textContent='Retrying upload…';
          setTimeout(()=>{
            handleFile._retried=false;
            // Re-send same file
            const fd2=new FormData(); fd2.append('file',file);
            const xhr2=new XMLHttpRequest();
            xhr2.open('POST',BASE+'/upload');
            xhr2.timeout=60000;
            xhr2.ontimeout=xhr.ontimeout;
            xhr2.upload.onprogress=function(e){
              if(e.lengthComputable) setP(40+(e.loaded/e.total)*55,'Uploading…');
            };
            xhr2.onload=xhr.onload;
            xhr2.onerror=xhr.onerror;
            xhr2.send(fd2);
          },1000);
        } else {
          dz.style.display='block'; dz.style.opacity='1'; dz.style.transform='none';
          wrap.style.display='none'; fc.style.display='none';
          err.textContent='⚠ Upload failed — server returned empty response. Please try again.';
          err.style.display='block';
          handleFile._retried=false;
        }
        return;
      }
      try{
        const data=JSON.parse(xhr.responseText);
        handleFile._retried=false;
        if(data.error){
          dz.style.display='block'; dz.style.opacity='1'; dz.style.transform='none';
          wrap.style.display='none'; fc.style.display='none';
          err.textContent='⚠ '+data.error; err.style.display='block'; return;
        }
        setP(100,'Done!');
        setTimeout(()=>{
          // Hide progress bar, show success line
          wrap.style.display='none';
          fc.textContent='✅ '+file.name+' — '+data.shape[0].toLocaleString()+' rows × '+data.shape[1]+' cols — starting pipeline…';
          // Also hide the header so only the success line is visible
          const hdr=document.getElementById('upload-header');
          if(hdr){ hdr.style.opacity='0'; hdr.style.maxHeight='0'; hdr.style.marginBottom='0'; }
          // After 3 seconds fade the entire upload screen out, then switch to progress
          setTimeout(()=>{
            const uploadScreen=document.getElementById('s-upload');
            uploadScreen.style.transition='opacity .4s ease, min-height .4s ease, padding .4s ease';
            uploadScreen.style.opacity='0';
            uploadScreen.style.minHeight='0';
            uploadScreen.style.padding='0';
            uploadScreen.style.overflow='hidden';
            setTimeout(()=>startCleaning(), 450);
          }, 3000);
        },300);
      }catch(e2){
        dz.style.display='block'; dz.style.opacity='1'; dz.style.transform='none';
        wrap.style.display='none'; fc.style.display='none';
        err.textContent='⚠ '+e2.message; err.style.display='block';
        handleFile._retried=false;
      }
    };
    xhr.onerror=function(){
      dz.style.display='block'; dz.style.opacity='1'; dz.style.transform='none';
      wrap.style.display='none'; fc.style.display='none';
      err.textContent='⚠ Upload failed — network error'; err.style.display='block';
    };
    xhr.send(fd);
  };
  reader.onerror=function(){
    dz.style.display='block'; dz.style.opacity='1'; dz.style.transform='none';
    wrap.style.display='none'; fc.style.display='none';
    err.textContent='⚠ Could not read file'; err.style.display='block';
  };
  reader.readAsDataURL(file);
}

const STEPS=['Removing duplicate responses','Detecting age & Likert outliers','✅ Validating categorical ranges','Removing straight-line responses','Checking late-night contradictions','Checking spend vs budget logic','Removing rows with excessive missing data','Imputing remaining missing values','Filtering target customers','Applying membership consistency rules'];

async function startCleaning(){
  _analyticsBuilt=false;
  window._ANALYTICS=null; window._CORR=null; window._CLUSTERS=null;
  window._RULES=null; window._REGRESSION=null;
  goTo('s-progress');
  const sl=document.getElementById('step-list');
  sl.innerHTML=STEPS.map((s,i)=>`<div class="step-row" id="step-${i}"><span class="step-dot"></span><span>${s}</span><span class="step-tick">✓</span></div>`).join('');
  let stepIdx=0;
  const fill=document.getElementById('prog-fill'),pct=document.getElementById('prog-pct'),lbl=document.getElementById('prog-label');
  const cleanProm=fetch(BASE+'/clean',{method:'POST'}).then(r=>r.json());
  const startT=Date.now();
  const iv=setInterval(()=>{
    const frac=Math.min((Date.now()-startT)/5000,.97);
    fill.style.width=(frac*100)+'%'; pct.textContent=Math.round(frac*100)+'%';
    const es=Math.min(Math.floor(frac*STEPS.length),STEPS.length-1);
    if(es>stepIdx){
      if(stepIdx>0){const p=document.getElementById(`step-${stepIdx-1}`);if(p){p.classList.remove('active');p.classList.add('done')}}
      const c=document.getElementById(`step-${es}`);if(c)c.classList.add('active');
      lbl.textContent=STEPS[es]; stepIdx=es;
    }
  },80);
  await Promise.all([cleanProm,new Promise(r=>setTimeout(r,5000))]);
  clearInterval(iv);
  STEPS.forEach((_,i)=>{const el=document.getElementById(`step-${i}`);if(el){el.classList.remove('active');el.classList.add('done')}});
  fill.style.width='100%'; pct.textContent='100%'; lbl.textContent='✅ Pipeline complete!';
  // Wait 2 seconds so user can read "Pipeline complete!"
  await new Promise(r=>setTimeout(r,2000));
  // Hide s-progress before switching — ensures layout recalculates without it
  // before s-nav is activated and scrollTo(top:0) fires.
  const progressScreen = document.getElementById('s-progress');
  progressScreen.style.display = 'none';
  goTo('s-nav');
  // No need to clear the inline style — #s-progress no longer has a permanent
  // display:flex rule, so .screen{display:none} already applies when inactive.
}

// ══ Summary screen ══
async function openSummary(){
  goTo('s-summary');
  if(_summaryLoaded){
    // Already have data — but Chart.js instances are destroyed when
    // the screen becomes display:none, so we must re-render them.
    // Re-fetch is cheap (Flask returns cached _store data instantly).
    try{
      const res=await fetch(BASE+'/summary_data');
      const d=await res.json();
      if(!d.error) renderSummary(d);
    }catch(e){}
    return;
  }
  try{
    const res=await fetch(BASE+'/summary_data');
    const d=await res.json();
    if(d.error){alert(d.error);goTo('s-nav');return}
    renderSummary(d); _summaryLoaded=true;
  }catch(e){alert('Failed: '+e.message);goTo('s-nav')}
}

function renderSummary(d){
  const kpis=[
    {label:'Raw Responses',val:d.raw_shape[0].toLocaleString(),sub:'before cleaning',cls:'c-info'},
    {label:'Cleaned Responses',val:d.n_final.toLocaleString(),sub:'target customers',cls:'c-green'},
    {label:'Total Removed',val:d.total_removed.toLocaleString(),sub:'all techniques',cls:'c-danger'},
    {label:'Columns',val:d.clean_columns.length,sub:'in cleaned dataset',cls:'c-accent'},
    {label:'Retention Rate',val:(d.n_final/d.raw_shape[0]*100).toFixed(1)+'%',sub:'responses kept',cls:'c-orange'},
    {label:'Completeness',val:d.reliability_pre.completeness.completeness_pct+'%',sub:'non-missing (raw)',cls:'c-warn'},
  ];
  document.getElementById('kpi-grid').innerHTML=kpis.map(k=>`<div class="kpi"><div class="k-label">${k.label}</div><div class="k-val ${k.cls}">${k.val}</div><div class="k-sub">${k.sub}</div></div>`).join('');
  document.getElementById('relevant-info').innerHTML=`<div style="background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:20px 22px;font-size:.86rem;line-height:1.75"><p>The survey collected <strong>${d.raw_shape[0].toLocaleString()}</strong> total responses. Only rows where <code style="background:rgba(47,143,106,.15);color:var(--accent2);padding:1px 6px;border-radius:4px">Prefer_Study_Cafe = "Yes"</code> are analytically relevant. After cleaning, <strong class="c-green">${d.n_final.toLocaleString()} respondents</strong> form the target-customer dataset used for all analysis.</p></div>`;
  const tb=document.getElementById('clean-tbody');
  tb.innerHTML='';
  let i=1;
  for(const[k,v]of Object.entries(d.clean_log)){
    tb.innerHTML+=`<tr><td style="color:var(--muted2);text-align:center">${i}</td><td><strong>${v.label}</strong></td><td><span style="font-size:1.05rem;font-weight:700;color:var(--warn)">${v.count}</span></td><td style="color:var(--muted2)">${v.method}</td></tr>`;i++;
  }
  const rel=d.reliability_pre;
  const rg=document.getElementById('rel-grid');
  rg.innerHTML='';
  const aC=rel.cronbach_alpha.value>=0.7?'c-green':rel.cronbach_alpha.value>=0.5?'c-warn':'c-danger';
  rg.innerHTML+=`<div class="rel-card"><div class="rc-title">Cronbach's Alpha <span class="pill pill-blue">Internal Consistency</span></div><div class="rc-desc">${rel.cronbach_alpha.description}</div><div class="rc-val ${aC}">${rel.cronbach_alpha.value??'N/A'}</div><div class="rc-rows"><div class="rc-row"><span class="rk">Interpretation</span><span class="rv">${rel.cronbach_alpha.interpretation}</span></div><div class="rc-row"><span class="rk">n</span><span class="rv">${rel.cronbach_alpha.n}</span></div></div></div>`;
  rg.innerHTML+=`<div class="rel-card"><div class="rc-title">Inter-Item Correlations <span class="pill pill-blue">Construct Validity</span></div><div class="rc-desc">${rel.inter_item_correlation.description}</div><div class="rc-val c-info">${rel.inter_item_correlation.average}</div><div class="rc-rows">${Object.entries(rel.inter_item_correlation.pairs).map(([k,v])=>`<div class="rc-row"><span class="rk">${k}</span><span class="rv">${v}</span></div>`).join('')}</div></div>`;
  const nC=rel.normality_age.is_normal?'c-green':'c-warn';
  rg.innerHTML+=`<div class="rel-card"><div class="rc-title">Shapiro-Wilk (Age) <span class="pill pill-green">Normality Test</span></div><div class="rc-desc">${rel.normality_age.description}</div><div class="rc-val ${nC}">${rel.normality_age.shapiro_stat??'N/A'}</div><div class="rc-rows"><div class="rc-row"><span class="rk">p-value</span><span class="rv">${rel.normality_age.p_value}</span></div><div class="rc-row"><span class="rk">Normal?</span><span class="rv">${rel.normality_age.is_normal?'✅ Yes':'⚠ No'}</span></div></div></div>`;
  const cpC=rel.completeness.completeness_pct>=85?'c-green':rel.completeness.completeness_pct>=70?'c-warn':'c-danger';
  rg.innerHTML+=`<div class="rel-card"><div class="rc-title">Completeness <span class="pill pill-orange">Data Integrity</span></div><div class="rc-desc">${rel.completeness.description}</div><div class="rc-val ${cpC}">${rel.completeness.completeness_pct}%</div><div class="rc-rows"><div class="rc-row"><span class="rk">Missing cells</span><span class="rv">${rel.completeness.missing_cells.toLocaleString()}</span></div><div class="rc-row"><span class="rk">Complete rows</span><span class="rv">${rel.completeness.complete_rows.toLocaleString()}</span></div></div></div>`;
  rg.innerHTML+=`<div class="rel-card"><div class="rc-title">⚠ Response Bias <span class="pill pill-orange">Survey Quality</span></div><div class="rc-desc">${rel.response_bias.description}</div><div class="rc-rows"><div class="rc-row"><span class="rk">Straight-line</span><span class="rv c-danger">${rel.response_bias.straight_line_count}</span></div><div class="rc-row"><span class="rk">Acquiescence</span><span class="rv c-warn">${rel.response_bias.acquiescence_bias_count}</span></div><div class="rc-row"><span class="rk">Disacquiescence</span><span class="rv c-warn">${rel.response_bias.disacquiescence_bias_count}</span></div></div></div>`;
  rg.innerHTML+=`<div class="rel-card"><div class="rc-title">Spend &gt; Budget <span class="pill pill-blue">Logic Check</span></div><div class="rc-desc">${rel.spend_budget_contradiction.description}</div><div class="rc-val c-danger">${rel.spend_budget_contradiction.count}</div><div class="rc-rows"><div class="rc-row"><span class="rk">Removed</span><span class="rv">✅ Yes</span></div></div></div>`;
  const hmE=Object.entries(rel.high_missing_columns.columns);
  rg.innerHTML+=`<div class="rel-card"><div class="rc-title">High-Missing Cols <span class="pill pill-orange">&gt;30% missing</span></div><div class="rc-desc">${rel.high_missing_columns.description}</div><div class="rc-rows">${hmE.length===0?'<div class="rc-row"><span class="rk">None</span><span class="rv c-green">✅</span></div>':hmE.map(([c,p])=>`<div class="rc-row"><span class="rk">${c}</span><span class="rv c-danger">${p}%</span></div>`).join('')}</div></div>`;
  const cg=document.getElementById('col-grid');
  cg.innerHTML='';
  for(const[col,s]of Object.entries(d.col_stats)){
    let ex='';
    if(s.mean!==undefined)ex=`<div class="cc-row"><span class="cc-k">Mean</span><span class="cc-v">${s.mean}</span></div><div class="cc-row"><span class="cc-k">Min/Max</span><span class="cc-v">${s.min}/${s.max}</span></div>`;
    else if(s.top_values)ex=Object.entries(s.top_values).map(([k,v])=>`<div class="cc-row"><span class="cc-k">"${k}"</span><span class="cc-v">${v}</span></div>`).join('');
    cg.innerHTML+=`<div class="col-card"><div class="cc-name">${col}</div><div class="cc-row"><span class="cc-k">Type</span><span class="cc-v">${s.dtype}</span></div><div class="cc-row"><span class="cc-k">Unique</span><span class="cc-v">${s.unique}</span></div><div class="cc-row"><span class="cc-k">Missing</span><span class="cc-v ${s.missing>0?'c-warn':'c-green'}">${s.missing}</span></div>${ex}</div>`;
  }
  const cols=d.clean_columns,rows=d.clean_preview;
  let tbl='<table class="clean-table"><thead><tr>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr></thead><tbody>';
  rows.forEach(r=>{tbl+='<tr>'+cols.map(c=>`<td>${r[c]??''}</td>`).join('')+'</tr>'});
  document.getElementById('preview-tbl').innerHTML=tbl+'</tbody></table>';
  document.getElementById('sum-loading').style.display='none';
  document.getElementById('sum-body').style.display='block';
}

// ══ Analytics Dashboard ══
async function openAnalytics(){
  goTo('s-analytics');
  if(_analyticsBuilt) return;
  _analyticsBuilt=true;
  // Show a loading overlay while computing
  const loadDiv = document.getElementById('analytics-loading');
  if(loadDiv) loadDiv.style.display='flex';
  try{
    const res = await fetch(BASE+'/analytics_data');
    const d   = await res.json();
    if(d.error){ alert('EDA error: '+d.error); if(loadDiv) loadDiv.style.display='none'; return; }
    // Set globals from live data
    window._ANALYTICS = d.analytics;
    window._CORR      = d.corr;
    window._CLUSTERS  = d.clusters;
    CLUSTER_NAMES.length=0;
    ((d.clusters&&d.clusters.cluster_means)||[]).forEach((c,i)=>CLUSTER_NAMES.push(c.name||('Cluster '+(i+1))));
    // v44: show rules in plain words ("Avg spend per visit: >600"), not as column=value codes
    const _ruleText=s=>{const i=String(s).indexOf('='); if(i<0) return s; const c=s.slice(0,i).replace(/_/g,' '); return c.charAt(0).toUpperCase()+c.slice(1).toLowerCase()+': '+s.slice(i+1);};
    (d.rules||[]).forEach(r=>{r.antecedent=_ruleText(r.antecedent); r.consequent=_ruleText(r.consequent);});
    window._RULES     = d.rules;
    window._RULES_META= d.rules_meta||{};
    window._PRICING   = d.pricing||null;
    window._REGRESSION= d.regression;
    window._VC        = d.analytics.value_counts;
    window._nFinal    = d.analytics.n_rows;
    if(loadDiv) loadDiv.style.display='none';
    buildAnalytics();
  }catch(e){
    if(loadDiv) loadDiv.style.display='none';
    alert('Failed to load EDA: '+e.message);
  }
}

const PALETTE=['#2b6ea3','#d2692c','#2f8f6a','#cfae2a','#a8508e','#6a86d8','#c4534a','#7f9a24'];
// VC is set dynamically after analytics_data fetch

Chart.defaults.maintainAspectRatio=false;
// v42 design layer: one typographic and colour voice for every chart (DESIGN.md)
Chart.defaults.font.family="'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif";
Chart.defaults.font.size=12;
Chart.defaults.color='#55635D';
Chart.defaults.borderColor='rgba(23,35,31,.08)';
Chart.defaults.elements.line.borderWidth=2;
Chart.defaults.plugins.legend.labels.boxWidth=10;
Chart.defaults.plugins.legend.labels.boxHeight=10;
Chart.defaults.plugins.tooltip.backgroundColor='#17231F';
Chart.defaults.plugins.tooltip.titleColor='#FFFFFF';
Chart.defaults.plugins.tooltip.bodyColor='#FFFFFF';
Chart.defaults.plugins.tooltip.padding=10;
Chart.defaults.plugins.tooltip.cornerRadius=6;
Chart.defaults.plugins.tooltip.displayColors=true;
Chart.defaults.animation.duration=200;
function _fitCanvas(el,cfg){
  // v41: one sizing rule for all charts — the canvas fills a box of known height
  let h=parseInt(el.dataset.h||el.getAttribute('height'),10)||260;
  if(cfg&&(cfg.type==='radar'||cfg.type==='polarArea')) h=Math.max(h,340);   // radars need room for labels + legend
  el.dataset.h=h;
  let box=el.parentElement;
  const sized=box&&(box.classList.contains('cv-box')||(box.style&&box.style.height));
  if(sized){ if(!box.style.position) box.style.position='relative'; if(parseInt(box.style.height,10)<h) box.style.height=h+'px'; }
  else{
    box=document.createElement('div'); box.className='cv-box';
    box.style.cssText='position:relative;width:100%;height:'+h+'px';
    el.parentNode.insertBefore(box,el); box.appendChild(el);
  }
  if(cfg){ cfg.options=cfg.options||{}; cfg.options.responsive=true; cfg.options.maintainAspectRatio=false; }
}
// ════════ v43 THEMES ════════
// Charts are drawn on canvas, which cannot read CSS variables, so each chart's
// colours are translated between the light and dark sets and the chart redrawn.
const TH_HEX_L2D={'#55635d':'#a9b5ae','#17231f':'#e9efea','#ffffff':'#18211e','#f5f7f4':'#101614','#dce3dd':'#2a3631','#c5d0c8':'#3b4b44','#6e7b75':'#8c9992',
  '#1f5c4a':'#6fcfae','#1f7a4d':'#5cc08a','#b3261e':'#f0766c','#a8741a':'#d9a441',
  '#2b6ea3':'#5495cf','#d2692c':'#c6610a','#2f8f6a':'#47a67d','#cfae2a':'#967400','#a8508e':'#a34991','#6a86d8':'#5a7bd8','#c4534a':'#d96d67','#7f9a24':'#657905','#7fbd9f':'#2f5e4b','#4f9a78':'#3f8467','#2c7656':'#5eab8a','#1a4c39':'#9ad6bb','#86c2a6':'#2e5c4a','#5ea485':'#3b7a61','#3f8667':'#4f9a7b','#2a6a4f':'#73b898','#184534':'#a6dcc3'};
const TH_RGB_L2D={'23,35,31':'233,239,234','245,247,244':'16,22,20','43,110,163':'84,149,207','210,105,44':'198,97,10','47,143,106':'71,166,125','207,174,42':'150,116,0','168,80,142':'163,73,145','106,134,216':'90,123,216','196,83,74':'217,109,103','127,154,36':'101,121,5'};
const _inv=o=>Object.fromEntries(Object.entries(o).map(([k,v])=>[v,k]));
const TH_HEX_D2L=_inv(TH_HEX_L2D), TH_RGB_D2L=_inv(TH_RGB_L2D);
function _theme(){ return document.documentElement.getAttribute('data-theme')==='dark'?'dark':'light'; }
function _mapColor(s,toDark){
  if(typeof s!=='string'||s.length>40) return s;
  let m=s.match(/^#([0-9a-fA-F]{6})([0-9a-fA-F]{2})?$/);
  if(m){ const t=(toDark?TH_HEX_L2D:TH_HEX_D2L)['#'+m[1].toLowerCase()]; return t?t+(m[2]||''):s; }
  m=s.match(/^rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,(.*)\)$/);
  if(m){ const t=(toDark?TH_RGB_L2D:TH_RGB_D2L)[m[1]+','+m[2]+','+m[3]]; return t?'rgba('+t+','+m[4]+')':s; }
  return s;
}
function _walkColors(o,toDark,depth){
  if(!o||typeof o!=='object'||depth>7) return;
  for(const k of Object.keys(o)){
    if(k==='data'||k==='labels') continue;                 // values and names, never colours
    const v=o[k];
    if(typeof v==='string') o[k]=_mapColor(v,toDark);
    else if(Array.isArray(v)){ for(let i=0;i<v.length;i++){ if(typeof v[i]==='string') v[i]=_mapColor(v[i],toDark); else _walkColors(v[i],toDark,depth+1); } }
    else if(v&&typeof v==='object'&&!(v instanceof Node)) _walkColors(v,toDark,depth+1);
  }
}
function _themeConfig(cfg,toDark){
  if(!cfg) return;
  if(cfg.options) _walkColors(cfg.options,toDark,0);
  if(cfg.data&&cfg.data.datasets) cfg.data.datasets.forEach(ds=>_walkColors(ds,toDark,0));
}
function _themeDefaults(){
  const dark=_theme()==='dark';
  Chart.defaults.color=dark?'#A9B5AE':'#55635D';
  Chart.defaults.borderColor=dark?'rgba(233,239,234,.10)':'rgba(23,35,31,.08)';
  Chart.defaults.plugins.tooltip.backgroundColor=dark?'#E9EFEA':'#17231F';
  Chart.defaults.plugins.tooltip.titleColor=dark?'#101614':'#FFFFFF';
  Chart.defaults.plugins.tooltip.bodyColor=dark?'#101614':'#FFFFFF';
}
function _syncChartThemes(){
  const want=_theme();
  Object.values(Chart.instances).forEach(ch=>{
    const have=ch.$scTheme||'light';
    if(have===want) return;
    _themeConfig(ch.config,want==='dark'); ch.$scTheme=want;
    // Chart.js keeps each element's resolved colours; clear them so bars repaint too
    try{ ch.getSortedVisibleDatasetMetas().forEach(m=>{ if(m.controller){ m.controller._cachedDataOpts={}; m.controller._sharedOptions=undefined; } }); }catch(e){}
    try{ ch.update('none'); }catch(e){}
  });
}
let _syncTimer=null;
function _syncSoon(){ clearTimeout(_syncTimer); _syncTimer=setTimeout(_syncChartThemes,40); }
function _markThemeButtons(){
  const t=_theme();
  document.querySelectorAll('[data-theme-btn]').forEach(b=>b.setAttribute('aria-pressed', b.dataset.themeBtn===t?'true':'false'));
}
function _sc3dTheme(){
  const dark=_theme()==='dark';
  return dark ? {ink:'#E9EFEA',muted:'#A9B5AE',grid:'rgba(233,239,234,.26)',surface:'#18211E',
                 colors:['#5495cf','#c6610a','#47a67d','#967400','#a34991','#5a7bd8','#d96d67','#657905']}
              : {ink:'#17231F',muted:'#55635D',grid:'rgba(23,35,31,.22)',surface:'#FFFFFF',colors:CLUSTER_COLORS};
}
function _applyTheme(t,persist){
  t=(t==='dark')?'dark':'light';
  document.documentElement.setAttribute('data-theme',t);
  if(persist){ try{ localStorage.setItem('sc-theme',t); }catch(e){} }
  _themeDefaults(); _syncChartThemes(); _markThemeButtons();
  if(window._sc3d){ try{ window._sc3d.draw(); }catch(e){} }
}
function setTheme(t){ _applyTheme(t,true); }
// ════════════════════════════════════════════════════════════
//  v47 — PLAIN-LANGUAGE GUIDES
//  Every chart gets an "About this chart" note: why the analysis is done and
//  what each axis (or slice, colour, bubble size) means, written for readers
//  with no statistics background. The same notes appear in the dashboard and
//  in every downloaded report, because the reports read them from the page.
//  Every what-if input gets a short explanation the same way.
// ════════════════════════════════════════════════════════════
const _H='Horizontal axis', _V='Vertical axis', _SL='Slices', _CL='Colours';
function _pcText(n){
  const p=(window._CLUSTERS&&window._CLUSTERS.pca&&window._CLUSTERS.pca.axes)||[], a=p[n-1];
  if(!a) return 'PC'+n+', a summary score that combines several survey answers into one number (see "What PC1, PC2 and PC3 mean").';
  const and=l=>l.length>1?l.slice(0,-1).join(', ')+' and '+l[l.length-1]:l[0];
  const hi=a.high.length?' A high score means above-average '+and(a.high)+(a.low.length?' and below-average '+and(a.low):'')+'.':'';
  return 'PC'+n+', a summary score that combines several survey answers into one number. It captures '+a.share_pct+'% of the differences between students.'+hi;
}
const CHART_GUIDE={
  // ── Exploratory analysis ─────────────────────────────────
  'ch-age':{p:'Shows the age mix of the students who would use a study café, so its look, menu and prices can suit the people most likely to come.',
    a:[[_H,'Age groups, in years.'],[_V,'Number of students in each age group.']]},
  'ch-time':{p:'Shows when in the day students would rather study at the café. This tells you when to have more staff and when to run offers.',
    a:[[_SL,'Each slice is a time of day. A bigger slice means more students prefer that time. Hover a slice for the number of students.']]},
  'ch-spend':{p:'Shows how much students usually spend on one visit, which helps set realistic menu prices.',
    a:[[_H,'Spend per visit, in rupee ranges.'],[_V,'Number of students in each spending range.']]},
  'ch-bev':{p:'Shows which drinks students want most, so the menu stocks what will sell.',
    a:[[_SL,'Each slice is a drink. A bigger slice means more students chose it.']]},
  'ch-food':{p:'Shows what kind of food students want with their study session.',
    a:[[_V,'Types of food.'],[_H,'Number of students who chose each type. A longer bar means more popular.']]},
  'ch-noise':{p:'Shows how quiet students want the café to be. This shapes the layout, for example silent zones versus areas where talking is fine.',
    a:[[_SL,'Each slice is a noise level. A bigger slice means more students prefer it.']]},
  'ch-seat':{p:'Shows which kinds of seats students prefer, to decide how many of each type to provide.',
    a:[['Wedges','Each wedge is a seat type. The further a wedge reaches from the centre, the more students chose it. The rings mark the number of students.']]},
  'ch-freq':{p:'Shows how often students say they would come in a week. This is the basis for estimating how many visitors to expect each day.',
    a:[[_H,'Visits per week, as answered in the survey (from rarely to daily).'],[_V,'Number of students giving each answer.']]},
  'ch-membership':{p:'Shows how many students would sign up for a paid monthly membership.',
    a:[[_SL,'Yes = would take a membership. No = would not.']]},
  'ch-likert3d':{p:'Compares how important three facilities are to students: power outlets, WiFi and productivity tools. Each was rated from 1 (very low) to 5 (very high).',
    a:[[_H,'The rating given, from 1 = very low importance to 5 = very high importance.'],[_V,'Number of students who gave that rating.'],[_CL,'Each colour is one facility, so the three can be compared side by side. Most bars on the right means the facility matters a lot.']]},
  'ch-hours-freq':{p:'Checks whether students who study for longer also plan to visit more often.',
    a:[[_H,'Study hours, as answered in the survey.'],[_V,'Number of students.'],[_CL,'Each bar is split by how often those students would visit. Larger "Daily" or "4-5" bands mean more frequent visitors in that group.']]},
  'ch-dist-night':{p:'Checks whether students who live closer to campus are more interested in studying late at night, which matters for opening hours.',
    a:[[_H,'Distance from campus.'],[_V,'Number of students.'],[_CL,'Yes = wants late-night study. No = does not.']]},
  'ch-grp-seat':{p:'Checks whether students who often study in groups prefer group tables, to plan the right mix of seats.',
    a:[[_H,'How often the student studies in a group.'],[_V,'Number of students.'],[_CL,'Each colour is the type of seat they prefer.']]},
  'ch-univ-budget':{p:'Compares how much students at each university are willing to spend per visit, to see which campus offers the most spending power.',
    a:[[_H,'Budget per visit, in rupee ranges.'],[_V,'Number of students.'],[_CL,'Each colour is one university.']]},
  'ch-corr':{p:'Shows which survey answers tend to move together, for example whether students who visit more often also spend more. It is a quick way to spot links worth looking into.',
    a:[['Rows and columns','The same list of survey questions runs down the side and along the bottom. Each square is where two questions meet.'],
       [_CL,'Green means that when one answer goes up, the other tends to go up too. Red means that when one goes up, the other tends to go down. Each colour has three shades: light for a weak link (0.1 to 0.3), medium for moderate (0.3 to 0.5) and dark for strong (0.5 or more). Grey means little or no link. Hover a square for the exact value, from −1 (perfectly opposite) to +1 (perfectly together).']]},
  'ch-reg-coef':{p:'A statistical model (logistic regression) estimates which student traits make someone more or less likely to want a membership. This shows the effect of each trait.',
    a:[[_V,'The student traits the model used.'],[_H,'Effect on the chance of wanting a membership. Bars to the right (green) raise the chance and bars to the left (red) lower it. Longer bars have a bigger effect. All traits are put on a common scale so their bars can be compared fairly.']]},
  'ch-reg-line':{p:'Shows the model\'s estimate of how likely a student is to want a membership at each budget level.',
    a:[[_H,'Budget per visit, in rupee ranges.'],[_V,'Estimated chance of wanting a membership, in %. The axis stops 3 points above the highest value so differences between budget levels are easier to see.']]},
  'ch-elbow':{p:'Students are split into customer groups (clusters) of people with similar habits. This chart helps choose how many groups to use. More groups always fit the data better, but after a point the gain is small. That bend, the "elbow", is the number chosen.',
    a:[[_H,'Number of customer groups tried.'],[_V,'How spread out students are within their groups (called inertia). Lower means tighter, more alike groups. Look for where the line stops falling steeply.']]},
  'ch-cluster-radar':{p:'Compares the customer groups on several traits at once, so each group has a recognisable shape.',
    a:[['Spokes','Each spoke is one trait, such as WiFi importance or spend. The further out a point is, the higher that group\'s average on the trait, compared with the full range seen in the data.'],[_CL,'Each colour is one customer group.']]},
  'ch-scatter2d':{p:'Places every student on a flat map so you can check that the customer groups really are different. Each dot is one student, coloured by group. Groups forming separate clouds are clearly different customers. Overlapping clouds are similar.',
    a:[[_H,()=>_pcText(1)],[_V,()=>_pcText(2)]]},
  'ch-scatter3d':{p:'The same idea as the flat map, with a third summary score added as depth. Drag the plot to turn it and see the groups from any angle. Groups that stay apart from every angle are clearly different customers.',
    a:[['PC1 axis',()=>_pcText(1)],['PC2 axis',()=>_pcText(2)],['PC3 axis',()=>_pcText(3)]]},
  'ch-cluster-bar':{p:'Compares the customer groups side by side on a few key measures.',
    a:[[_H,'The customer groups.'],[_V,'For age, spend and visit frequency: where the group\'s average sits between the lowest (0) and highest (100) value seen in the data. For membership willingness: the % of the group who said Yes.'],[_CL,'Each colour is one measure.']]},
  'ch-cluster-age-mem':{p:'Shows each customer group\'s average age against how keen it is on membership, to see which group to target with memberships.',
    a:[[_H,'Average age of the group, in years.'],[_V,'% of the group who would take a membership.'],['Bubble size','Number of students in the group. Bigger bubbles are bigger groups.']]},
  'ch-cluster-age-dist':{p:'Shows the age make-up of each customer group.',
    a:[[_H,'Age groups, in years.'],[_V,'% of that customer group falling in each age band. The bars of one colour add up to 100%.'],[_CL,'Each colour is one customer group.']]},
  'ch-arm-bubble':{p:'Association rules find answers that tend to appear together, such as "students who spend over ₹600 tend to want a membership". This chart places each rule by how common and how reliable it is. Click a bubble to have its rule explained in plain words.',
    a:[[_H,'Support: the share of all students the rule applies to. For example, 12% means 12 in every 100 students.'],[_V,'Confidence: of the students matching the first part of the rule, the share who also match the second part. Higher means more reliable.'],['Bubble size','Lift: how many times more likely the second part is when the first part is true, compared with pure chance. Bigger bubbles are stronger links.']]},
  'ch-arm-scatter':{p:'Shows every rule by how reliable it is and how much stronger than chance it is. Click a dot to have its rule explained in plain words.',
    a:[[_H,'Confidence: of students matching the "if" part, the share who also match the "then" part.'],[_V,'Lift: 1 means no better than chance. 1.5 means 50% more likely than chance.']]},
  'ch-arm-bar':{p:'Lists the 20 most reliable rules.',
    a:[[_V,'Each rule, written as "if this → then that".'],[_H,'Confidence: of students matching the "if" part, the share who also match the "then" part.']]},
  'ch-arm-heatmap':{p:'Shows at a glance which answers are linked to which. It starts with the 25 strongest rules (those with the highest lift); the Expand button shows every rule.',
    a:[['Rows and columns','Down the side: the "if" part of each rule. Along the bottom: the "then" part. Each square is one if-then pair.'],[_CL,'Four clearly different greens, from light (fewest students) to dark (most students), split at the quarter points of support; the key under the chart gives the ranges. The number in each square is its support. Grey squares mean no rule links the two.']]},
  // ── Predictive analytics ─────────────────────────────────
  'ch-lev-time':{p:'A lever for drilling down: untick a time slot under Select, or click its slice, and every other predictive chart shows only the students who prefer the slots still ticked. For example, untick Late Night to see the café through the eyes of its daytime students.',
    a:[[_SL,'Each slice is a time slot. Its size is that slot\'s share of all responses. Faded slices are not part of the current drill-down.']]},
  'ch-lev-spend':{p:'A lever for drilling down: untick a spending range under Select, or click its slice, and every other predictive chart shows only the students in the ranges still ticked, for example only the higher spenders.',
    a:[[_SL,'Each slice is a spending range. Its size is that range\'s share of all responses. Faded slices are not part of the current drill-down.']]},
  'ch-lev-seat':{p:'A lever for drilling down: untick a seat type under Select, or click its slice, and every other predictive chart shows only the students who prefer the seat types still ticked.',
    a:[[_SL,'Each slice is a seat type. Its size is that seat type\'s share of all responses. Faded slices are not part of the current drill-down.']]},
  'ch-time-dist':{p:'Shows how expected customers split across the day\'s time slots. It is the starting point for staffing and opening hours.',
    a:[[_SL,'Each slice is a time slot. A bigger slice means more customers prefer it.']]},
  'ch-slot-rev':{p:'Estimates which time slots would bring in the most money.',
    a:[[_H,'Time slots of the day.'],[_V,'Estimated revenue in rupees for every 100 customers.']]},
  'ch-spend-by-slot':{p:'Shows whether students who come at different times spend different amounts. This is useful for time-based pricing or offers.',
    a:[[_V,'Time slots of the day.'],[_H,'Average spend per visit, in rupees.']]},
  'ch-peak-fi':{p:'A prediction model (random forest) was trained to guess which time slot a student prefers. This shows which answers it relied on most.',
    a:[[_V,'Student traits the model used.'],[_H,'Importance: how much the model relied on each trait. Longer bars were more useful for the prediction. All bars add up to 100%.']]},
  'ch-slot-foot':{p:'Estimates how many visitors to expect in each time slot on a typical day.',
    a:[[_H,'Time slots of the day.'],[_V,'Expected visitors per day in that slot.']]},
  'ch-weekly-pattern':{p:'Illustrates how visits might spread across the week. The survey did not ask about days, so this uses assumed patterns and is a guide only.',
    a:[[_H,'Day of the week.'],[_V,'Expected visits on that day.']]},
  'ch-foot-dist':{p:'Checks whether students living closer to campus would visit more often.',
    a:[[_H,'Distance from campus.'],[_V,'Average number of visits per week.']]},
  'ch-spend-dist':{p:'Shows how customers split by how much they spend on one visit.',
    a:[[_SL,'Each slice is a spending range. A bigger slice means more customers in that range.']]},
  'ch-rev-slot':{p:'Shows which time slots bring in the most money in a day.',
    a:[[_H,'Time slots of the day.'],[_V,'Revenue per day from that slot, in rupees.']]},
  'ch-rev-projection':{p:'Projects monthly revenue for the first year if the customer base grows steadily.',
    a:[[_H,'Months after opening (M1 = first month).'],[_V,'Revenue for the month, in rupees.']]},
  'ch-spend-univ':{p:'Compares how much students from each university spend per visit, to see which campus to focus marketing on.',
    a:[[_H,'University.'],[_V,'Average spend per visit, in rupees.']]},
  'ch-seat-overall':{p:'Recommends how to split the café\'s seats between seat types, based on what students prefer.',
    a:[[_SL,'Each slice is a seat type. Its size is the recommended share of seats.']]},
  'ch-seat-by-time':{p:'Shows how the preferred type of seat changes through the day.',
    a:[[_H,'Time slots of the day.'],[_V,'% of that slot\'s visitors who prefer each seat type.'],[_CL,'Each colour is a seat type.']]},
  'ch-mem-models':{p:'Several prediction models were tested on how well they predict who wants a membership. This compares their accuracy so the best one can be used.',
    a:[[_H,'The prediction models tested.'],[_V,'Accuracy: the share of students whose answer the model predicted correctly. It is measured on students the model had not seen (5-fold cross-validation).']]},
  'ch-mem-spend':{p:'Shows whether students who spend more per visit are keener on membership, to decide whom to target first.',
    a:[[_H,'Spend per visit, in rupee ranges.'],[_V,'% of students in that range who would take a membership.']]},
  'ch-mem-dist':{p:'Shows whether students living closer to campus are keener on membership.',
    a:[[_H,'Distance from campus.'],[_V,'% of students at that distance who would take a membership.']]},
  'ch-mem-fi':{p:'Shows which student traits the best model relied on most to predict who wants a membership. Items with "×" combine two traits, for example spend × visits is roughly total weekly spend.',
    a:[[_V,'Student traits the model used.'],[_H,'Importance: how much the model relied on each trait. Longer bars were more useful. All bars add up to 100%.']]},
  'mem-conf-matrix':{p:'Checks the best model\'s predictions against what students actually answered, using a quarter of the students held back for testing.',
    a:[['Rows and columns','Rows show what the student actually said. Columns show what the model predicted.'],['Cells','"True" cells are correct predictions and "False" cells are mistakes. The bigger the True numbers compared with the False ones, the better the model.']]},
  // ── What-if scenarios ────────────────────────────────────
  'wi-ch-projection':{p:'Projects monthly revenue for the first year under your scenario settings.',
    a:[[_H,'Months after opening (M1 = first month).'],[_V,'Revenue for the month, in rupees. The solid line is what the seats can earn. A dashed line, when shown, is what you would earn if every interested student could get a seat.']]},
  'wi-ch-slots':{p:'Shows how a day\'s revenue splits across the time slots.',
    a:[[_H,'Time slots of the day.'],[_V,'Revenue per day from that slot, in rupees.']]},
  'wi-ch-sensitivity':{p:'Tests what happens to revenue if the real number of visitors turns out higher or lower than estimated. If all bars are equal, the café is already full, so extra visitors cannot add revenue.',
    a:[[_H,'Change in visitors compared with the estimate. "Base" is the estimate itself.'],[_V,'Monthly revenue, in rupees.']]},
  'wi-ch-membership':{p:'Compares money from walk-in customers with money from membership fees.',
    a:[[_H,'Type of revenue, per month and per year.'],[_V,'Amount in rupees.']]},
  'wi-ch-seating':{p:'Compares how many visits students want to make with how many the seats can host. If demand is higher than capacity, the café turns people away.',
    a:[[_H,'Daily demand = visits students want to make per day. Daily seat capacity = visits the seats can host per day. Visitors served = the smaller of the two.'],[_V,'Visits per day.']]},
  'wi-ch-profit':{p:'Projects profit or loss for each month of the first year, after all costs.',
    a:[[_H,'Months after opening (M1 = first month).'],[_V,'Profit for the month, in rupees. Green bars are a profit and red bars are a loss.']]},
  'wi-ch-breakeven-bar':{p:'Puts the first month\'s costs next to its income, to see whether the café covers its costs.',
    a:[[_H,'Cost items (fixed costs, staff, food and drink costs) and income items (walk-in revenue, membership fees).'],[_V,'Amount for the first month, in rupees.']]},
  'wi-ch-price-sweep-rev':{p:'Tests what happens to revenue if prices are raised or cut. Higher prices earn more per visit, but students whose budget is exceeded may stop coming. It shows the whole year by default; use Annual, Quarterly or Monthly above the charts to compare a single quarter or month. The highlighted bar earns the most in the period shown.',
    a:[[_H,'Price change compared with what students spend today.'],[_V,'Revenue in rupees for the period chosen above the charts: the year, a quarter or a month.']]},
  'wi-ch-price-sweep-profit':{p:'The same price test, measured as profit after all costs: fixed costs, food and drink, and staff (who scale with the visitors each price brings). The recommended price is the one with the most profit over the whole year.',
    a:[[_H,'Price change compared with what students spend today.'],[_V,'Profit in rupees for the period chosen above the charts. Red bars are a loss.']]},
  'wi-ch-comp-rev':{p:'Shows how annual revenue falls if a competitor opens nearby and takes away some of the students.',
    a:[[_H,'Share of potential customers lost to a competitor.'],[_V,'Annual revenue, in rupees.']]},
  'wi-ch-comp-loss':{p:'Shows how much revenue a competitor would cost you. If the café is full, losing some students may cost nothing, because others fill the seats.',
    a:[[_H,'Share of potential customers lost to a competitor.'],[_V,'Revenue lost per year compared with no competitor, in rupees.']]},
  'wi-ch-staff-count':{p:'Estimates the minimum number of staff needed in each time slot, based on the busiest moment (one staff member for every 20 customers present at once).',
    a:[[_H,'Time slots of the day.'],[_V,'Staff members needed.']]},
  'wi-ch-staff-profit':{p:'Checks whether each time slot earns enough to pay for its staff.',
    a:[[_H,'Time slots of the day.'],[_V,'Amount per day, in rupees.'],[_CL,'Slot revenue, staff cost, and what is left after paying staff.']]},
  'wi-ch-tier-split':{p:'Shows how monthly membership income splits between the Basic and Premium levels.',
    a:[[_SL,'Each slice is a membership level. Its size is that level\'s share of membership income.']]},
  'wi-ch-tier-compare':{p:'Checks whether offering two membership levels (Basic and Premium) earns more than one single price.',
    a:[[_H,'Single price versus two levels combined.'],[_V,'Monthly membership revenue, in rupees.']]},
  'wi-ch-compare':{p:'Compares the scenarios you saved, side by side.',
    a:[[_H,'Your saved scenarios, in the order you saved them.'],[_V,'Annual revenue, in rupees.'],[_CL,'Revenue from visits and revenue from memberships.']]},
};
function _guideHtml(g){
  const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;');
  const val=v=>typeof v==='function'?v():v;
  return '<div class="cg-head">About this chart</div><p class="cg-p">'+esc(val(g.p))+'</p>'
    +(g.a&&g.a.length?'<dl class="cg-axes">'+g.a.map(r=>'<div><dt>'+esc(r[0])+'</dt><dd>'+esc(val(r[1]))+'</dd></div>').join('')+'</dl>':'');
}
function _addChartGuide(id){
  const g=CHART_GUIDE[id], el=document.getElementById(id); if(!g||!el) return;
  const card=el.closest('.chart-card'); if(!card) return;
  let box=card.querySelector('.chart-guide[data-for="'+id+'"]');
  if(!box){ box=document.createElement('div'); box.className='chart-guide'; box.dataset.for=id; card.appendChild(box); }
  box.innerHTML=_guideHtml(g);
}
function _addChartGuides(){ Object.keys(CHART_GUIDE).forEach(_addChartGuide); }

// ── What-if inputs ─────────────────────────────────────────
const WI_PARAM_HELP={
  'wi-catchment':'How many students live or study close enough to be potential customers. Almost every result scales with this number.',
  'wi-growth':'How much demand grows each month as the café becomes known. 5 means 5% more each month. Use a negative number for decline.',
  'wi-price':'Raise or cut prices compared with what students said they spend today. +10 means 10% higher. Higher prices earn more per visit, but students whose budget is exceeded may stop coming.',
  'wi-hours':'How many hours a day the café is open. Longer hours let each seat be used by more visitors per day.',
  'wi-memtarget':'Of the students who said they would want a membership, the share you expect to actually sign up.',
  'wi-occupancy':'How full the seats get on average while open. 70 means 70% of seats in use. 100% is rarely realistic.',
  'wi-seats':'Number of seats in the café. With opening hours and occupancy, it sets the most visitors the café can serve in a day.',
  'wi-fixedcosts':'Monthly costs that do not change with the number of customers, such as rent, utilities, internet and maintenance. Staff wages are set separately.',
  'wi-competitor':'Share of potential customers you expect to lose if a competing café opens nearby. 20 means 20% of students go elsewhere.',
  'wi-staffcost':'What one staff member costs for one shift. The number of staff per time slot is worked out automatically.',
  'wi-tier1price':'Monthly fee for the Basic membership.',
  'wi-tier1uptake':'Of all members, the share who choose Basic. The rest choose Premium.',
  'wi-tier2price':'Monthly fee for the Premium membership.',
};
const WI_ASSUMPTIONS='Built into every scenario: food and drink cost 35% of walk-in revenue; one staff member is needed for every 20 customers present at once; each visit lasts as long as the average study session reported in the survey; and spend per visit comes from the survey answers.';
function _addParamHelp(){
  Object.keys(WI_PARAM_HELP).forEach(id=>{
    const inp=document.getElementById(id); if(!inp) return;
    const f=inp.closest('.wi-field'); if(!f||f.querySelector('.wi-help')) return;
    const h=document.createElement('div'); h.className='wi-help'; h.id=id+'-help'; h.textContent=WI_PARAM_HELP[id];
    f.appendChild(h); inp.setAttribute('aria-describedby', h.id);
  });
  const ctl=document.querySelector('.wi-controls');
  if(ctl && !document.getElementById('wi-assumptions')){
    const n=document.createElement('p'); n.id='wi-assumptions'; n.className='wi-assumptions'; n.textContent=WI_ASSUMPTIONS;
    const gs=ctl.querySelectorAll('.wi-group'), grid=gs.length?gs[gs.length-1]:ctl.querySelector('.wi-grid'); (grid||ctl).insertAdjacentElement(grid?'afterend':'beforeend', n);
  }
}
document.addEventListener('DOMContentLoaded', function(){ _addChartGuides(); _addParamHelp(); });

// ════════════════════════════════════════════════════════════
//  v45 — Scatter3D: a real, rotatable 3D scatter plot on a canvas.
//  No external library. Drag (or arrow keys) to rotate, double-click
//  to reset. The same function is embedded in the HTML report.
// ════════════════════════════════════════════════════════════
function Scatter3D(canvas, opts){
  var self=this;
  self.canvas=canvas; self.o=opts||{};
  self.yaw0=-0.65; self.pitch0=0.32; self.yaw=self.yaw0; self.pitch=self.pitch0; self.zoom=1; self.hover=null;
  var pts=self.o.points||[];
  var mn=[Infinity,Infinity,Infinity], mx=[-Infinity,-Infinity,-Infinity];
  pts.forEach(function(p){ [p.x,p.y,p.z].forEach(function(v,i){ if(v<mn[i])mn[i]=v; if(v>mx[i])mx[i]=v; }); });
  self.mn=mn; self.mx=mx;
  function nrm(v,i){ var r=mx[i]-mn[i]; return r>0 ? ((v-mn[i])/r)*2-1 : 0; }
  self.P=pts.map(function(p){ return {v:[nrm(p.x,0),nrm(p.y,1),nrm(p.z,2)], c:p.c||0, raw:p}; });
  function theme(){
    var t=(typeof self.o.theme==='function')?self.o.theme():(self.o.theme||{});
    return {ink:t.ink||'#17231F', muted:t.muted||'#55635D', grid:t.grid||'rgba(23,35,31,.22)', surface:t.surface||'#FFFFFF',
            colors:t.colors||['#2b6ea3','#d2692c','#2f8f6a','#cfae2a','#a8508e','#6a86d8','#c4534a','#7f9a24'],
            font:t.font||"'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif"};
  }
  function rot(v){
    var cy=Math.cos(self.yaw), sy=Math.sin(self.yaw), cp=Math.cos(self.pitch), sp=Math.sin(self.pitch);
    var x=v[0]*cy+v[2]*sy, z=-v[0]*sy+v[2]*cy, y=v[1];
    return [x, y*cp-z*sp, y*sp+z*cp];
  }
  self.project=function(v,w,h){
    var r=rot(v), d=4.2, f=d/(d-r[2]), s=Math.min(w,h)*0.25*self.zoom;
    return {x:w/2+r[0]*f*s, y:h/2+10-r[1]*f*s, z:r[2], f:f};
  };
  self.draw=function(){
    var cv=self.canvas, dpr=window.devicePixelRatio||1;
    var w=cv.clientWidth||cv.width/dpr||600, h=cv.clientHeight||cv.height/dpr||380;
    if(cv.width!==Math.round(w*dpr)||cv.height!==Math.round(h*dpr)){ cv.width=Math.round(w*dpr); cv.height=Math.round(h*dpr); }
    var ctx=cv.getContext('2d'), T=theme(); ctx.setTransform(dpr,0,0,dpr,0,0); ctx.clearRect(0,0,w,h);
    if(self.o.fill){ ctx.fillStyle=T.surface; ctx.fillRect(0,0,w,h); }
    // legend
    var names=self.o.names||[], lx=8, ly=14; ctx.font='12px '+T.font; ctx.textBaseline='middle'; ctx.textAlign='left';
    names.forEach(function(n,i){
      var tw=ctx.measureText(n).width+26; if(lx+tw>w-8){ lx=8; ly+=18; }
      ctx.fillStyle=T.colors[i%T.colors.length]; ctx.beginPath(); ctx.arc(lx+5,ly,5,0,Math.PI*2); ctx.fill();
      ctx.fillStyle=T.muted; ctx.fillText(n,lx+15,ly); lx+=tw;
    });
    // cube
    var C=[[-1,-1,-1],[1,-1,-1],[1,1,-1],[-1,1,-1],[-1,-1,1],[1,-1,1],[1,1,1],[-1,1,1]].map(function(v){return self.project(v,w,h);});
    ctx.strokeStyle=T.grid; ctx.lineWidth=1;
    [[0,1],[1,2],[2,3],[3,0],[4,5],[5,6],[6,7],[7,4],[0,4],[1,5],[2,6],[3,7]].forEach(function(e){
      ctx.beginPath(); ctx.moveTo(C[e[0]].x,C[e[0]].y); ctx.lineTo(C[e[1]].x,C[e[1]].y); ctx.stroke(); });
    // axes from the (-1,-1,-1) corner, with names
    var ax=self.o.axes||['X','Y','Z'], ends=[[1.18,-1,-1],[-1,1.18,-1],[-1,-1,1.18]];
    ctx.strokeStyle=T.muted; ctx.lineWidth=1.5; ctx.fillStyle=T.ink; ctx.font='600 12px '+T.font; ctx.textAlign='center';
    ends.forEach(function(e,i){ var a=C[0], b=self.project(e,w,h);
      ctx.beginPath(); ctx.moveTo(a.x,a.y); ctx.lineTo(b.x,b.y); ctx.stroke(); ctx.fillText(ax[i],b.x,b.y-9); });
    // points, far to near
    var Q=self.P.map(function(p){ var q=self.project(p.v,w,h); q.p=p; return q; }).sort(function(a,b){return a.z-b.z;});
    self._screen=Q;
    Q.forEach(function(q){
      var r=3.4*q.f*Math.sqrt(self.zoom);
      ctx.globalAlpha=0.55+0.4*((q.z+1.75)/3.5); ctx.fillStyle=T.colors[q.p.c%T.colors.length];
      ctx.beginPath(); ctx.arc(q.x,q.y,r,0,Math.PI*2); ctx.fill();
      ctx.globalAlpha=1; ctx.strokeStyle=T.surface; ctx.lineWidth=0.6; ctx.stroke();
    });
    ctx.globalAlpha=1;
    if(self.hover){ var q=self.hover, nm=(names[q.p.c]||('Cluster '+(q.p.c+1)));
      var t1=nm, t2=ax[0]+' '+q.p.raw.x.toFixed(2)+'   '+ax[1]+' '+q.p.raw.y.toFixed(2)+'   '+ax[2]+' '+q.p.raw.z.toFixed(2);
      ctx.font='12px '+T.font; var bw=Math.max(ctx.measureText(t1).width,ctx.measureText(t2).width)+18, bh=40;
      var bx=Math.min(Math.max(q.x+12,4),w-bw-4), by=Math.min(Math.max(q.y-bh-8,4),h-bh-4);
      ctx.fillStyle=T.ink; ctx.fillRect(bx,by,bw,bh); ctx.fillStyle=T.surface; ctx.textAlign='left';
      ctx.font='600 12px '+T.font; ctx.fillText(t1,bx+9,by+13); ctx.font='12px '+T.font; ctx.fillText(t2,bx+9,by+28);
      ctx.strokeStyle=T.ink; ctx.lineWidth=2; ctx.beginPath(); ctx.arc(q.x,q.y,6,0,Math.PI*2); ctx.stroke();
    }
    if(self.o.hint!==false){ ctx.fillStyle=T.muted; ctx.font='12px '+T.font; ctx.textAlign='left'; ctx.fillText(self.o.hint||'Drag to rotate · scroll to zoom · double-click to reset',8,h-10); }
  };
  self.reset=function(){ self.yaw=self.yaw0; self.pitch=self.pitch0; self.zoom=1; self.hover=null; self.draw(); };
  if(self.o.interactive!==false){
    var drag=null; canvas.style.touchAction='none'; canvas.style.cursor='grab';
    canvas.addEventListener('pointerdown',function(e){ drag={x:e.clientX,y:e.clientY}; self.hover=null; canvas.style.cursor='grabbing'; try{canvas.setPointerCapture(e.pointerId);}catch(_){ } });
    canvas.addEventListener('pointermove',function(e){
      if(drag){ self.yaw+=(e.clientX-drag.x)*0.01; self.pitch=Math.max(-1.5,Math.min(1.5,self.pitch+(e.clientY-drag.y)*0.01)); drag={x:e.clientX,y:e.clientY}; self.draw(); return; }
      var r=canvas.getBoundingClientRect(), mx_=e.clientX-r.left, my_=e.clientY-r.top, best=null, bd=100;
      (self._screen||[]).forEach(function(q){ var d=(q.x-mx_)*(q.x-mx_)+(q.y-my_)*(q.y-my_); if(d<bd){bd=d;best=q;} });
      if(best!==self.hover){ self.hover=best; self.draw(); }
    });
    var end=function(){ drag=null; canvas.style.cursor='grab'; };
    canvas.addEventListener('pointerup',end); canvas.addEventListener('pointercancel',end);
    canvas.addEventListener('pointerleave',function(){ if(!drag&&self.hover){ self.hover=null; self.draw(); } });
    canvas.addEventListener('dblclick',function(){ self.reset(); });
    canvas.addEventListener('wheel',function(e){ e.preventDefault(); var f=Math.exp(-e.deltaY*(e.ctrlKey?0.01:0.0015)); self.zoom=Math.max(0.5,Math.min(3,self.zoom*Math.min(1.5,Math.max(0.67,f)))); self.hover=null; self.draw(); },{passive:false});
    canvas.addEventListener('keydown',function(e){ var k=e.key, s=0.12, used=true;
      if(k==='ArrowLeft') self.yaw-=s; else if(k==='ArrowRight') self.yaw+=s;
      else if(k==='ArrowUp') self.pitch=Math.max(-1.5,self.pitch-s); else if(k==='ArrowDown') self.pitch=Math.min(1.5,self.pitch+s);
      else if(k==='+'||k==='=') self.zoom=Math.min(3,self.zoom*1.1); else if(k==='-') self.zoom=Math.max(0.5,self.zoom/1.1);
      else if(k==='Home'||k==='0'){ self.reset(); } else used=false;
      if(used){ e.preventDefault(); self.draw(); } });
    if(typeof ResizeObserver!=='undefined'){ try{ new ResizeObserver(function(){ self.draw(); }).observe(canvas); }catch(_){ } }
  }
  self.draw();
}

_themeDefaults();
document.addEventListener('DOMContentLoaded', _markThemeButtons);

const FEAT_LABELS={visits_pw:'Visit frequency',spend_v:'Spend per visit',study_h:'Study hours',dist_n:'Distance from campus',gs_n:'Group study',time_n:'Preferred time',late_n:'Late-night preference',mem_n:'Membership willingness',Power_Outlet_Requirement:'Power outlet need',WiFi_Importance:'WiFi importance',Productivity_Tools_Interest:'Productivity tools',spend_x_visit:'Spend × visits',dist_x_late:'Distance × late night',study_x_visit:'Study hours × visits'};
function _dvNormalize(cfg){
  if(!cfg||!cfg.data) return;
  if(Array.isArray(cfg.data.labels)) cfg.data.labels=cfg.data.labels.map(l=>FEAT_LABELS[l]||l);
  // colour carries identity: a single series is one colour, not one colour per bar
  if(!cfg.__drilled&&!cfg.__semantic&&cfg.type==='bar'&&cfg.data.datasets&&cfg.data.datasets.length===1){
    const ds=cfg.data.datasets[0];
    if(Array.isArray(ds.backgroundColor)&&new Set(ds.backgroundColor).size>=3) ds.backgroundColor='#2b6ea3';
  }
}
// v55: a category keeps the same colour on every chart; ordered scales use one green ramp
const _ENT={
  time:{Morning:'#cfae2a',Afternoon:'#d2692c',Evening:'#a8508e','Late Night':'#2b6ea3'},
  seat:{'Individual Desk':'#2b6ea3','Group Table':'#d2692c','Sofa Lounge':'#2f8f6a','Private Booth':'#cfae2a','Standing Desk':'#a8508e'},
  yesno:{Yes:'#2f8f6a',No:'#2b6ea3'}
};
const _ORD={spend:['100-200','200-400','400-600','>600'],freq:['Never','Rarely','Sometimes','Occasionally','Frequently']};
const _RAMP4=['#7fbd9f','#4f9a78','#2c7656','#1a4c39'], _RAMP5=['#86c2a6','#5ea485','#3f8667','#2a6a4f','#184534'];
function _normCat(l){ return String(l).replace(/[₹\s,]/g,'').replace(/[–—]/g,'-').replace(/^<=?200$/,'100-200'); }
function _entLookup(labels){
  if(!Array.isArray(labels)||labels.length<2) return null;
  for(const k in _ENT){ const t=_ENT[k]; if(labels.every(l=>t[l])) return labels.map(l=>t[l]); }
  const n=labels.map(_normCat);
  if(n.every(l=>_ORD.spend.includes(l))) return n.map(l=>_RAMP4[_ORD.spend.indexOf(l)]);
  if(labels.every(l=>_ORD.freq.includes(l))) return labels.map(l=>_RAMP5[_ORD.freq.indexOf(l)]);
  return null;
}
function _entityColors(cfg){
  if(!cfg||!cfg.data||cfg.__semantic||cfg.__drilled) return;
  const ds=cfg.data.datasets||[], t=cfg.type;
  if(['pie','doughnut','polarArea'].includes(t)&&ds.length===1){
    const c=_entLookup(cfg.data.labels); if(c) ds[0].backgroundColor=c;
  }else if(ds.length>=2){
    const c=_entLookup(ds.map(d=>d.label)); if(!c) return;
    ds.forEach((d,i)=>{
      const a=(typeof d.backgroundColor==='string'&&/^#[0-9a-f]{8}$/i.test(d.backgroundColor))?d.backgroundColor.slice(7):'';
      d.backgroundColor=c[i]+a; if(t==='line'||d.type==='line'||d.borderColor) d.borderColor=c[i];
    });
  }
}
function mkChart(id,cfg){
  const el=document.getElementById(id);
  if(!el) return null;
  _fitCanvas(el,cfg);
  _dvNormalize(cfg);
  _entityColors(cfg);
  const _dark=_theme()==='dark';
  if(_dark) _themeConfig(cfg,true);
  // Destroy existing chart on this canvas before creating a new one
  const existing=Chart.getChart(el);
  if(existing) existing.destroy();
  const _ch=new Chart(el, cfg);
  _ch.$scTheme=_dark?'dark':'light';
  _syncSoon();            // also themes charts created outside mkChart
  try{ _addChartGuide(id); }catch(e){}
  try{ _afterChart(id,cfg); }catch(e){ console.warn(e); }
  return _ch;
}

function buildAnalytics(){
  const _vcK=col=>{const o=(window._ANALYTICS&&window._ANALYTICS.value_count_order||{})[col]; const d=(window._VC||{})[col]||{}; return (o&&o.length)?o:Object.keys(d);};
  const _vcV=col=>{const d=(window._VC||{})[col]||{}; return _vcK(col).map(k=>d[k]||0);};
  const VC=window._VC||{};
  const ANALYTICS=window._ANALYTICS||{};
  const CORR=window._CORR||{labels:[],matrix:[]};
  const CLUSTERS=window._CLUSTERS||{scatter2d:[],scatter3d:[],cluster_means:[],elbow:[]};
  const RULES=window._RULES||[];
  const n_rows = ANALYTICS.n_rows||0;
  const mbYes  = (VC.Willingness_for_Membership||{}).Yes||0;
  const mbPct  = n_rows>0 ? (mbYes/n_rows*100).toFixed(1)+'%' : '—';
  const spendVC= VC.Avg_Spend_per_visit||{};
  const topSpend= Object.entries(spendVC).sort((a,b)=>b[1]-a[1])[0];
  const topSpendLabel = topSpend ? '₹'+String(topSpend[0]).replace('-','–') : '—';
  const topSpendN = topSpend ? (+topSpend[1]).toLocaleString() : '';
  // ── KPI strip
  const kd=[
    {val:n_rows.toLocaleString(),lab:'Students in the clean data',col:'var(--accent)'},
    {val:mbPct,lab:'Would take a membership',col:'var(--accent2)'},
    {val:topSpendLabel,lab:'Most common spend per visit'+(topSpendN?' ('+topSpendN+' students)':''),col:'var(--accent3)'},
    {val:ANALYTICS.age_mean?(+ANALYTICS.age_mean).toFixed(1):'—',lab:'Average age, in years',col:'var(--warn)'},
    {val:RULES.length,lab:'Association rules found',col:'var(--info)'},
    {val:(window._CLUSTERS&&window._CLUSTERS.best_k||4),lab:'Customer groups (k-means)',col:'var(--purple)'},
  ];
  document.getElementById('kpi-strip').innerHTML=kd.map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  // helper
  function hoverPlugin(labelsFn){
    return {id:'hover',afterEvent(chart,args){
      const e=args.event;
      if(e.type==='mousemove'){
        const pts=chart.getElementsAtEventForMode(e.native,'nearest',{intersect:true},true);
        if(pts.length){
          const pt=pts[0];const ds=chart.data.datasets[pt.datasetIndex];
          const lbl=chart.data.labels?chart.data.labels[pt.index]:'';
          const val=ds.data[pt.index];
          const info=labelsFn?labelsFn(pt,lbl,val,ds):null;
          if(info) showTip(e.native,info);
        } else hideTip();
      } else if(e.type==='mouseout') hideTip();
    }};
  }

  const baseFont={color:'#55635D',size:10};
  const gridCol='rgba(23,35,31,.05)';

  // 1. Age histogram — LIVE from analytics_data
  const ageBucketLabels=ANALYTICS.age_bucket_labels||Object.keys(ANALYTICS.age_histogram||{});
  const ageBucketMap=ANALYTICS.age_histogram||{};
  const ageData=ageBucketLabels.map(k=>ageBucketMap[k]||0);
  mkChart('ch-age',{type:'bar',data:{labels:ageBucketLabels,datasets:[{data:ageData,backgroundColor:PALETTE,borderRadius:6}]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`Count: ${c.parsed.y}`}}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // 2. Preferred time donut
  const pt=VC.Preferred_Time;
  mkChart('ch-time',{type:'doughnut',data:{labels:_vcK('Preferred_Time'),datasets:[{data:_vcV('Preferred_Time'),backgroundColor:PALETTE,borderWidth:2,borderColor:'#FFFFFF',hoverOffset:8}]},options:{plugins:{legend:{position:'bottom',labels:{color:'#55635D',boxWidth:10,font:{size:10}}},tooltip:{callbacks:{label:c=>`${c.label}: ${c.parsed} (${(c.parsed/(window._ANALYTICS&&window._ANALYTICS.n_rows||697)*100).toFixed(1)}%)`}}},cutout:'55%'}});

  // 3. Spend bar
  const sp=VC.Avg_Spend_per_visit;
  mkChart('ch-spend',{type:'bar',data:{labels:_vcK('Avg_Spend_per_visit'),datasets:[{data:_vcV('Avg_Spend_per_visit'),backgroundColor:['#2f8f6a','#2b6ea3','#d2692c','#7f9a24'],borderRadius:6}]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y} respondents`}}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // 4. Beverage pie
  const bv=VC.Beverage_Preference;
  mkChart('ch-bev',{type:'pie',data:{labels:_vcK('Beverage_Preference'),datasets:[{data:_vcV('Beverage_Preference'),backgroundColor:PALETTE,borderWidth:2,borderColor:'#FFFFFF',hoverOffset:6}]},options:{plugins:{legend:{position:'bottom',labels:{color:'#55635D',boxWidth:10,font:{size:10}}},tooltip:{callbacks:{label:c=>`${c.label}: ${c.parsed} (${(c.parsed/(window._ANALYTICS&&window._ANALYTICS.n_rows||697)*100).toFixed(1)}%)`}}}}});

  // 5. Food bar
  const fo=VC.Food_Preference;
  mkChart('ch-food',{type:'bar',data:{labels:_vcK('Food_Preference'),datasets:[{data:_vcV('Food_Preference'),backgroundColor:PALETTE,borderRadius:6}]},options:{indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // 6. Noise doughnut
  const no=VC.Noise_Tolerance;
  mkChart('ch-noise',{type:'doughnut',data:{labels:_vcK('Noise_Tolerance'),datasets:[{data:_vcV('Noise_Tolerance'),backgroundColor:PALETTE,borderWidth:2,borderColor:'#FFFFFF',hoverOffset:8}]},options:{plugins:{legend:{position:'bottom',labels:{color:'#55635D',boxWidth:10,font:{size:10}}}},cutout:'50%'}});

  // 7. Seating polar area
  const se=VC.Preferred_Seating_Type;
  mkChart('ch-seat',{type:'polarArea',data:{labels:_vcK('Preferred_Seating_Type'),datasets:[{data:_vcV('Preferred_Seating_Type'),backgroundColor:PALETTE.map(c=>c+'cc'),borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{color:'#55635D',boxWidth:10,font:{size:10}}}},scales:{r:{ticks:{color:'#55635D',backdropColor:'transparent'},grid:{color:gridCol}}}}});

  // 8. Visit frequency
  const vf=VC.Visit_Frequency_per_week;
  mkChart('ch-freq',{type:'bar',data:{labels:_vcK('Visit_Frequency_per_week'),datasets:[{data:_vcV('Visit_Frequency_per_week'),backgroundColor:PALETTE,borderRadius:6}]},options:{plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // 9. Membership doughnut
  const mb=VC.Willingness_for_Membership;
  mkChart('ch-membership',{type:'doughnut',data:{labels:['Yes','No'],datasets:[{data:[mb.Yes||0,mb.No||0],backgroundColor:['#2f8f6a','#2b6ea3'],borderWidth:2,borderColor:'#FFFFFF',hoverOffset:8}]},options:{plugins:{legend:{position:'bottom',labels:{color:'#55635D',boxWidth:10}}},cutout:'60%'}});

  // 10. Likert grouped bar — LIVE from analytics_data
  const likertDist=ANALYTICS.likert_dist||{};
  const ld_power=(likertDist['Power_Outlet_Requirement']||[0,0,0,0,0]);
  const ld_wifi =(likertDist['WiFi_Importance']||[0,0,0,0,0]);
  const ld_prod =(likertDist['Productivity_Tools_Interest']||[0,0,0,0,0]);
  mkChart('ch-likert3d',{type:'bar',data:{labels:['1 — Very Low','2 — Low','3 — Neutral','4 — High','5 — Very High'],datasets:[
    {label:'Power Outlet Req.',data:ld_power,backgroundColor:'rgba(43,110,163,.8)',borderRadius:4},
    {label:'WiFi Importance',data:ld_wifi,backgroundColor:'rgba(47,143,106,.8)',borderRadius:4},
    {label:'Productivity Tools',data:ld_prod,backgroundColor:'rgba(210,105,44,.8)',borderRadius:4},
  ]},options:{plugins:{legend:{labels:{color:'#55635D'}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.parsed.y} respondents`}}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // EDA 1: Study hours vs visit freq — LIVE
  const XL=ANALYTICS.xtab_labels||{};
  const shLabels=XL.sh||[];
  const vfLabels=XL.vf||Object.keys(ANALYTICS.xtab_sh_vf||{});
  const xtab_sh_vf=ANALYTICS.xtab_sh_vf||{};
  mkChart('ch-hours-freq',{type:'bar',data:{labels:shLabels,datasets:vfLabels.map((lbl,i)=>({label:lbl,data:xtab_sh_vf[lbl]||shLabels.map(()=>0),backgroundColor:PALETTE[i%PALETTE.length]+'cc',borderRadius:2}))},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}}},scales:{x:{stacked:true,ticks:{color:'#55635D'},grid:{color:gridCol}},y:{stacked:true,ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // EDA 2: Distance vs Late Night — LIVE
  const distLabels=XL.dist||[];
  const xtab_dist_ln=ANALYTICS.xtab_dist_ln||{};
  const lnCols=['#2b6ea3cc','#2f8f6acc','#d2692ccc','#7f9a24cc'];
  mkChart('ch-dist-night',{type:'bar',data:{labels:distLabels,datasets:(XL.ln||Object.keys(xtab_dist_ln)).map((l,i)=>(
    {label:'Late Night: '+l,data:xtab_dist_ln[l]||distLabels.map(()=>0),backgroundColor:lnCols[i%lnCols.length],borderRadius:4}))},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // EDA 3: Group study vs seating — LIVE
  const gsLabels=XL.gs||[];
  const xtab_gs_seat=ANALYTICS.xtab_gs_seat||{};
  mkChart('ch-grp-seat',{type:'bar',data:{labels:gsLabels,datasets:(XL.seat||Object.keys(xtab_gs_seat)).map((s,i)=>(
    {label:s,data:xtab_gs_seat[s]||gsLabels.map(()=>0),backgroundColor:PALETTE[i%PALETTE.length]+'cc',borderRadius:3}))},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});

  // EDA 4: University × Budget radar — LIVE
  const budgetCats=XL.budget||[];
  const univLabels=budgetCats;
  const xtab_univ=ANALYTICS.xtab_univ_budget||{};
  const univs=ANALYTICS.univs||Object.keys(xtab_univ);
  const univColors=['#2b6ea3','#2f8f6a','#d2692c','#7f9a24','#c4534a','#cfae2a'];
  const univBgColors=['rgba(43,110,163,.15)','rgba(47,143,106,.15)','rgba(210,105,44,.15)','rgba(127,154,36,.15)','rgba(196,83,74,.15)','rgba(207,174,42,.15)'];
  const univDS=univs.map((u,i)=>({label:u,data:xtab_univ[u]||budgetCats.map(()=>0),borderColor:univColors[i%univColors.length],backgroundColor:univBgColors[i%univBgColors.length],pointRadius:3}));
  mkChart('ch-univ-budget',{type:'bar',data:{labels:univLabels,datasets:univDS.map((d,i)=>({label:d.label,data:d.data,backgroundColor:PALETTE[i%PALETTE.length],borderRadius:3}))},options:{plugins:{legend:{labels:{boxWidth:10}}},scales:{x:{title:{display:true,text:'Budget per visit (₹)'},grid:{display:false}},y:{title:{display:true,text:'Respondents'}}}}});

  // Correlation heatmap
  buildCorrHeatmap(CORR);

  // Regression charts
  buildRegression(window._REGRESSION||{});

  // Clustering
  buildClustering(CLUSTERS);

  // ARM
  buildARM(RULES);
  try{ buildPricing(); }catch(e){ console.warn('pricing',e); }
}

function buildCorrHeatmap(CORR){
  const labels=(CORR&&CORR.labels)||[];
  const matrix=(CORR&&CORR.matrix)||[];
  const n=labels.length;
  const heatData=[];
  for(let i=0;i<n;i++) for(let j=0;j<n;j++) heatData.push({x:j,y:i,v:matrix[i][j]});
  const el=document.getElementById('ch-corr');
  if(!el) return;
  const getColor=_corrColor;
  _fitCanvas(el,null);
  _hmLegend('ch-corr',_corrLegendHtml());
  new Chart(el,{type:'scatter',data:{datasets:[{data:heatData,
    pointStyle:'rect',
    pointRadius:ctx=>{const c=ctx.chart;const w=c.width,h=c.height;return Math.min(w,h)/(n*2.2)},
    backgroundColor:ctx=>getColor(ctx.raw.v),
    borderWidth:0
  }]},options:{
    plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const d=ctx.raw;const lbl1=labels[d.y],lbl2=labels[d.x];return `${lbl1} ↔ ${lbl2}: ${d.v!==null?d.v.toFixed(3):'N/A'}`}}}},
    scales:{
      x:{type:'linear',min:-0.5,max:n-0.5,afterBuildTicks:a=>{a.ticks=labels.map((_,i)=>({value:i}));},ticks:{stepSize:1,autoSkip:false,callback:v=>Number.isInteger(v)?(labels[v]||''):'',color:'#55635D',maxRotation:45},grid:{display:false}},
      y:{type:'linear',min:-0.5,max:n-0.5,reverse:true,afterBuildTicks:a=>{a.ticks=labels.map((_,i)=>({value:i}));},ticks:{stepSize:1,autoSkip:false,callback:v=>Number.isInteger(v)?(labels[v]||''):'',color:'#55635D'},grid:{display:false}}
    }
  }});
}

function buildRegression(reg){
  const gC='rgba(23,35,31,.05)';
  const hasReg = reg && reg.features && reg.features.length > 0;
  const features = hasReg ? reg.features : [];
  const coefs    = hasReg ? reg.coefs    : [];
  const bProbs   = (reg && reg.budget_probs && reg.budget_probs.length > 0) ? reg.budget_probs : [];

  if(!hasReg){
    // No real regression data — show a clear unavailable message instead of fake data
    document.getElementById('coef-bars').innerHTML =
      '<p style="color:var(--muted2);font-size:.8rem;padding:12px 0">Regression data not available — run cleaning first.</p>';
    const _accEl2 = document.getElementById('reg-acc-val');
    if(_accEl2) _accEl2.textContent = 'N/A';
    const _predEl2 = document.getElementById('reg-top-predictor');
    if(_predEl2) _predEl2.textContent = 'No regression data available.';
    return;
  }

  const colors=coefs.map(c=>c>0?'#2f8f6a':'#c4534a');
  mkChart('ch-reg-coef',{type:'bar',data:{labels:features,datasets:[{data:coefs,backgroundColor:colors,borderRadius:4}]},options:{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`Coefficient: ${c.parsed.x.toFixed(3)}`}}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',font:{size:10}},grid:{color:gC}}}}});
  const maxC=Math.max(...coefs.map(Math.abs));
  document.getElementById('coef-bars').innerHTML=features.map((f,i)=>{
    const pct=Math.abs(coefs[i])/maxC*100;
    const col=coefs[i]>0?'#2f8f6a':'#c4534a';
    return `<div class="coef-row"><span class="cr-name">${f}</span><div class="cr-bar"><div class="cr-fill" style="width:${pct}%;background:${col}"></div></div><span class="cr-val" style="color:var(--text)">${coefs[i].toFixed(3)}</span></div>`;
  }).join('');

  if(bProbs.length > 0){
    const budgetLabels = (reg.budget_labels&&reg.budget_labels.length===bProbs.length)?reg.budget_labels:bProbs.map((_,i)=>'Level '+(i+1));
    mkChart('ch-reg-line',{type:'line',data:{labels:budgetLabels,datasets:[{label:'P(Membership=Yes)',data:bProbs,borderColor:'#2b6ea3',backgroundColor:'rgba(43,110,163,.1)',fill:true,tension:.4,pointRadius:5,pointBackgroundColor:'#2b6ea3'}]},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}}},scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{min:0,max:Math.min(1,Math.ceil((Math.max(...bProbs)+0.03)*100)/100),ticks:{color:'#55635D',callback:v=>(v*100).toFixed(0)+'%'},grid:{color:gC}}}}});
  }

  // ── Dynamically update accuracy % and strongest predictor ──────────────
  const _accEl = document.getElementById('reg-acc-val');
  if(_accEl && reg && reg.accuracy){
    _accEl.textContent = (reg.accuracy * 100).toFixed(1) + '%';
  }
  const _predEl = document.getElementById('reg-top-predictor');
  if(_predEl && features && features.length > 0){
    // The feature with the largest absolute coefficient is the strongest predictor
    const _maxIdx = coefs.reduce((best,c,i) => Math.abs(c) > Math.abs(coefs[best]) ? i : best, 0);
    const _topFeat = features[_maxIdx];
    const _topCoef = coefs[_maxIdx];
    const _dir = _topCoef > 0 ? 'increases' : 'decreases';
    _predEl.textContent = `${_topFeat} is the strongest predictor — higher values ${_dir} the probability of membership willingness.`;
  }
}

function buildClustering(CLUSTERS){
  const elbow=(CLUSTERS&&CLUSTERS.elbow)||[];
  const bestK=(CLUSTERS&&CLUSTERS.best_k)||4;
  const _elbowSubEl=document.getElementById('elbow-sub');
  if(_elbowSubEl) _elbowSubEl.textContent=`K=${bestK} selected as optimal cluster count (knee point of inertia curve).`;
  mkChart('ch-elbow',{type:'line',data:{labels:elbow.map(e=>e.k),datasets:[{label:'Inertia',data:elbow.map(e=>e.inertia),borderColor:'#2b6ea3',backgroundColor:'rgba(43,110,163,.1)',fill:true,tension:.4,pointRadius:5,pointBackgroundColor:elbow.map(e=>e.k===bestK?'#7f9a24':'#2b6ea3'),pointBorderWidth:elbow.map(e=>e.k===bestK?3:1),pointBorderColor:elbow.map(e=>e.k===bestK?'#fff':'#2b6ea3')}]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`K=${c.label}: Inertia=${c.parsed.y.toFixed(0)}`}}},scales:{x:{title:{display:true,text:'Number of Clusters (K)',color:'#55635D'},ticks:{color:'#55635D'},grid:{color:gridCol}},y:{ticks:{color:'#55635D'},grid:{color:gridCol}}}}});
  const cm=(CLUSTERS&&CLUSTERS.cluster_means)||[];
  // Radar: 7 features only — Age removed (scale ~22 vs 1-5 distorts the chart)
  const radarLabels=['WiFi','Power','Productivity','Spend','Budget','Study Hours','Visit Freq'];
  function norm(arr){const mn=Math.min(...arr),mx=Math.max(...arr);return arr.map(v=>mx===mn?0:((v-mn)/(mx-mn)*100).toFixed(1))}
  const rKeys=['wifi','power','productivity','spend','budget','study_hours','visit_freq'];
  const rData=cm.map((c,i)=>({label:CLUSTER_NAMES[i],data:rKeys.map(k=>(c.norm&&c.norm[k]!==undefined)?c.norm[k]:0),borderColor:CLUSTER_COLORS[i],backgroundColor:CLUSTER_COLORS[i]+'22',pointRadius:3}));
  mkChart('ch-cluster-radar',{type:'radar',data:{labels:radarLabels,datasets:rData},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}}},scales:{r:{min:0,max:100,ticks:{color:'#55635D',backdropColor:'transparent',font:{size:8}},grid:{color:gridCol},pointLabels:{color:'#55635D',font:{size:9}}}}}});
  // Chart A: Age vs Membership Rate bubble (per cluster)
  const maxClusterSize = Math.max(...cm.map(c=>c.size||1));
  mkChart('ch-cluster-age-mem',{type:'bubble',data:{datasets:cm.map((c,i)=>({
    label:CLUSTER_NAMES[i],
    data:[{x:+(c.age).toFixed(1), y:+(c.membership*100).toFixed(1), r:Math.max(8,Math.round((c.size/maxClusterSize)*32))}],
    backgroundColor:CLUSTER_COLORS[i]+'bb',
    borderColor:CLUSTER_COLORS[i],
    borderWidth:2
  }))},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10,font:{size:9}}},
    tooltip:{callbacks:{label:c=>`${c.dataset.label}: Age ${c.raw.x}y | Mem ${c.raw.y}% | n=${cm[c.datasetIndex]?.size||''}`}}},
    scales:{
      x:{title:{display:true,text:'Mean Age (years)',color:'#55635D'},ticks:{color:'#55635D'},grid:{color:gridCol}},
      y:{title:{display:true,text:'Membership Rate (%)',color:'#55635D'},min:0,max:100,ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gridCol}}
    }}});

  // Chart B: Age bracket breakdown per cluster (using scatter2d point ages)
  const ageBrackets=(window._ANALYTICS&&window._ANALYTICS.age_bucket_labels)||[];
  const clusterAgeCounts=cm.map(c=>c.age_dist||ageBrackets.map(()=>0));
  mkChart('ch-cluster-age-dist',{type:'bar',data:{labels:ageBrackets,
    datasets:cm.map((c,i)=>({
      label:CLUSTER_NAMES[i],
      data:clusterAgeCounts[i],
      backgroundColor:CLUSTER_COLORS[i]+'cc',
      borderRadius:3
    }))},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:9,font:{size:9}}}},
      scales:{
        x:{ticks:{color:'#55635D',font:{size:9}},grid:{color:gridCol}},
        y:{ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gridCol},title:{display:true,text:'% of cluster',color:'#55635D',font:{size:9}}}
      }}});
  const legend2d=document.getElementById('scatter-legend-2d');
  const legend3d=document.getElementById('scatter-legend-3d');
  [legend2d,legend3d].forEach(el=>{el.innerHTML=CLUSTER_NAMES.map((n,i)=>`<span class="cl-item"><span class="cl-dot" style="background:${CLUSTER_COLORS[i]}"></span>${n}</span>`).join('')});
  const s2=(CLUSTERS&&CLUSTERS.scatter2d)||[];
  const scatter2dDatasets=CLUSTER_COLORS.slice(0,cm.length).map((col,ci)=>({
    label:CLUSTER_NAMES[ci],
    data:s2.filter(p=>p.cluster===ci).map(p=>({x:p.x,y:p.y,age:p.age,wifi:p.wifi,spend:p.spend})),
    backgroundColor:col+'99',pointRadius:4,pointHoverRadius:7
  }));
  mkChart('ch-scatter2d',{type:'scatter',data:{datasets:scatter2dDatasets},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10,font:{size:9}}},tooltip:{callbacks:{label:c=>`${c.dataset.label} | Age:${c.raw.age} WiFi:${c.raw.wifi} Spend:${c.raw.spend}`}}},scales:{x:{title:{display:true,text:'PC1',color:'#55635D'},ticks:{color:'#55635D'},grid:{color:gridCol}},y:{title:{display:true,text:'PC2',color:'#55635D'},ticks:{color:'#55635D'},grid:{color:gridCol}}}}});
  // Real 3D: the three principal components drawn in perspective, rotatable by the user
  const s3=(CLUSTERS&&CLUSTERS.scatter3d)||[];
  const cv3=document.getElementById('ch-scatter3d');
  if(cv3){
    const fresh=cv3.cloneNode(false); cv3.parentNode.replaceChild(fresh,cv3);   // drop listeners from an earlier build
    window._sc3d=new Scatter3D(fresh,{points:s3.map(p=>({x:p.x,y:p.y,z:p.z,c:p.cluster,age:p.age})),
      names:CLUSTER_NAMES.slice(0,cm.length), axes:['PC1','PC2','PC3'], theme:_sc3dTheme});
  }
  // Plain-language key for the axes of the two cluster plots
  const pcaBox=document.getElementById('pca-explain'), pca=CLUSTERS&&CLUSTERS.pca;
  if(pcaBox){
    if(pca&&pca.axes&&pca.axes.length){
      const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;');
      document.getElementById('pca-explain-sub').textContent='The clusters are formed on '+pca.n_variables+' variables ('+pca.variables.join(', ')+'). '
        +'That many dimensions cannot be drawn, so principal component analysis combines them into a few summary scores, the principal components (PC). '
        +'Each PC is a weighted mix of all the variables, and PC1 captures the most variation between respondents. '
        +'Together the axes shown hold '+pca.total_share_pct+'% of the variation, so the plots are a faithful but not complete picture. '
        +'The scores have no units: what matters is distance. Respondents close together have similar profiles.';
      document.querySelector('#pca-explain-tbl tbody').innerHTML=pca.axes.map(a=>'<tr><td><strong>'+a.name+'</strong></td><td>'+a.share_pct+'%</td><td>'
        +esc(a.high.join(', ')||'no single variable dominates')+'</td><td>'+esc(a.low.join(', ')||'(nothing stands out; lower simply means less of those)')+'</td></tr>').join('');
      pcaBox.style.display='block';
      ['ch-scatter2d','ch-scatter3d'].forEach(_addChartGuide);
    } else pcaBox.style.display='none';
  }
  const gridC=gridCol;
  const _nv=(c,k)=>(c.norm&&c.norm[k]!==undefined)?c.norm[k]:0;
  mkChart('ch-cluster-bar',{type:'bar',data:{labels:cm.map((c,i)=>CLUSTER_NAMES[i]),datasets:[
    {label:'Age (index)',data:cm.map(c=>_nv(c,'age')),raw:cm.map(c=>c.age+' yrs'),backgroundColor:'#2b6ea3cc',borderRadius:4},
    {label:'Spend level (index)',data:cm.map(c=>_nv(c,'spend')),raw:cm.map(c=>'level '+c.spend),backgroundColor:'#2f8f6acc',borderRadius:4},
    {label:'Visit frequency (index)',data:cm.map(c=>_nv(c,'visit_freq')),raw:cm.map(c=>'level '+c.visit_freq),backgroundColor:'#d2692ccc',borderRadius:4},
    {label:'Membership willingness (%)',data:cm.map(c=>+(c.membership*100).toFixed(1)),raw:cm.map(c=>(c.membership*100).toFixed(0)+'% willing'),backgroundColor:'#7f9a24cc',borderRadius:4},
  ]},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.parsed.y} (actual: ${c.dataset.raw[c.dataIndex]})`}}},scales:{x:{ticks:{color:'#55635D',font:{size:10}},grid:{color:gridCol}},y:{min:0,max:100,title:{display:true,text:'0–100 = position within the range seen in the data',color:'#55635D',font:{size:9}},ticks:{color:'#55635D'},grid:{color:gridCol}}}}});
}

// ════════════════════════════════════════════════════════════
//  v48 — INTERACTIONS
//  Drill-downs, radar isolate/compare with callouts, wheel zoom, rule
//  selection with a descriptive callout, stepped heatmap colours, and the
//  predictive-analytics levers. Every interaction can be undone with a
//  reset button, and all of it is cleared before a report is exported.
// ════════════════════════════════════════════════════════════
const _RM=()=>window.matchMedia&&window.matchMedia('(prefers-reduced-motion: reduce)').matches;
const _ink=()=>_theme()==='dark'?'#E9EFEA':'#17231F';
const _mut=()=>_theme()==='dark'?'#A9B5AE':'#55635D';
const _surf=()=>_theme()==='dark'?'#18211E':'#FFFFFF';
const _escH=s=>String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
function _fade(c,a){ try{ return Chart.helpers.color(c).alpha(a).rgbString(); }catch(e){ return c; } }
function _boxOf(el){ if(!el) return null; const b=el.closest('.cv-box')||el.parentElement; if(b&&!b.style.position) b.style.position='relative'; return b; }
const _CFG={};            // the config each chart was last built from (mkChart records it)
const _INIT={};           // chart id → function run after the chart is (re)built

// ── Card toolbar: buttons that sit in the top-right corner of a chart card ──
function _cardTools(card){
  let t=card.querySelector(':scope > .card-tools');
  if(!t){ t=document.createElement('div'); t.className='card-tools'; card.prepend(t); card.classList.add('has-tools'); }
  return t;
}
function _syncToolPad(card){
  const t=card.querySelector(':scope > .card-tools'), ti=card.querySelector('.cc-title');
  if(t&&ti) ti.style.paddingRight=(t.offsetWidth+12)+'px';
}

// ════════ 1. DRILL-DOWN ════════
// Click an x-axis value (or a bar, or a chip) to swap the chart for a horizontal
// bar breakdown of that value, with labels on each bar and the value axis on top.
const _DRILL_OPTS={
  'ch-hours-freq':{cat:'Study hours',value:'Number of students',part:'visit-frequency group',share:true,fmt:v=>Math.round(v).toLocaleString()},
  'ch-grp-seat':{cat:'Group study frequency',value:'Number of students',part:'seat type',share:true,fmt:v=>Math.round(v).toLocaleString()},
  'ch-univ-budget':{cat:'Budget per visit',value:'Number of students',part:'university',share:true,fmt:v=>Math.round(v).toLocaleString()},
  'ch-cluster-bar':{cat:'Customer group',value:'Score from 0 to 100 (membership: % willing)',part:'measure',share:false,fmt:v=>(+v).toFixed(0)},
  'ch-cluster-age-dist':{cat:'Age group',value:'% of the customer group',part:'customer group',share:false,fmt:v=>(+v).toFixed(1)+'%'},
  'ch-seat-by-time':{cat:'Time slot',value:'% of that slot\'s visitors',part:'seat type',share:false,fmt:v=>(+v).toFixed(1)+'%'},
  'ch-arm-bubble':{cat:'Support range',value:'Confidence (%)',part:'rule',share:false,fmt:v=>(+v).toFixed(0)+'%',custom:true},
};
const _DRILL={};
function _drillCats(id){
  const st=_DRILL[id];
  if(!st.o.custom) return (st.cfg.data.labels||[]).map(String);
  return st.bands.map(b=>b.label);
}
function _armBands(cfg){
  const xs=cfg.data.datasets[0].data.map(p=>p.x).sort((a,b)=>a-b);
  if(!xs.length) return [];
  const q=f=>xs[Math.min(xs.length-1,Math.floor(f*(xs.length-1)))];
  let edges=[xs[0],q(.25),q(.5),q(.75),xs[xs.length-1]].map(v=>Math.round(v*200)/200);   // 0.5-point steps
  edges=[...new Set(edges)];
  if(edges.length<2) edges=[xs[0],xs[xs.length-1]+0.005];
  const bands=[];
  for(let k=0;k<edges.length-1;k++){
    const lo=edges[k], hi=edges[k+1], last=k===edges.length-2;
    bands.push({lo,hi,last,label:(lo*100).toFixed(1)+'–'+(hi*100).toFixed(1)+'%'});
  }
  return bands;
}
function _inBand(b,x){ return x>=b.lo-1e-9 && (b.last? x<=b.hi+1e-9 : x<b.hi-1e-9); }
function _drillInit(id,cfg){
  const card=document.getElementById(id).closest('.chart-card'); if(!card) return;
  const st=_DRILL[id]=_DRILL[id]||{};
  st.o=_DRILL_OPTS[id]; st.cfg=cfg; st.active=null;
  if(st.o.custom) st.bands=_armBands(cfg);
  let bar=card.querySelector('.drill-bar');
  if(!bar){
    bar=document.createElement('div'); bar.className='drill-bar';
    const box=_boxOf(document.getElementById(id));
    box.parentNode.insertBefore(bar,box);
  }
  const cats=_drillCats(id);
  bar.innerHTML='<span class="drill-label">'+_escH(st.o.cat)+':</span>'
    +cats.map((c,i)=>'<button type="button" class="drill-chip" data-i="'+i+'" aria-pressed="false">'+_escH(c)+'</button>').join('')
    +'<span class="drill-hint">Click a value to see its breakdown</span>';
  bar.querySelectorAll('.drill-chip').forEach(b=>b.addEventListener('click',()=>{ const i=+b.dataset.i; _drillTo(id, st.active===i?null:i); }));
  _attachClick(id);
}
function _drillCfg(id,i){
  const st=_DRILL[id], o=st.o, base=st.cfg;
  let labels, vals, cols, extra=[];
  if(o.custom){
    const b=st.bands[i], pts=base.data.datasets[0].data, bg=base.data.datasets[0].backgroundColor;
    const sel=pts.map((p,k)=>({p,c:Array.isArray(bg)?bg[k]:bg})).filter(r=>_inBand(b,r.p.x)).sort((a,b2)=>b2.p.y-a.p.y).slice(0,15);
    labels=sel.map(r=>{ const s='If '+r.p.antecedent+' → '+r.p.consequent; return s.length>58?s.slice(0,57)+'…':s; });
    vals=sel.map(r=>+(r.p.y*100).toFixed(1)); cols=sel.map(()=>'#2b6ea3'); extra=sel.map(r=>'lift '+(+r.p.lift).toFixed(2));
  } else {
    const ds=base.data.datasets.filter(d=>!d.hidden);
    labels=ds.map(d=>d.label);
    vals=ds.map(d=>{ const v=d.data[i]; return (v&&typeof v==='object')?+v.y:(+v||0); });
    cols=ds.map(d=>Array.isArray(d.backgroundColor)?d.backgroundColor[i]:d.backgroundColor);
    extra=ds.map(d=>Array.isArray(d.raw)?d.raw[i]:'');
  }
  const total=vals.reduce((a,b)=>a+b,0);
  const txt=(v,k)=>o.fmt(v)+(o.share&&total?' ('+(v/total*100).toFixed(0)+'%)':'')+(extra[k]?' · '+extra[k]:'');
  const longest=Math.max(0,...vals.map((v,k)=>txt(v,k).length));
  return {type:'bar',__drilled:true,plugins:window.ChartDataLabels?[ChartDataLabels]:[],
    data:{labels,datasets:[{label:_drillCats(id)[i],data:vals,backgroundColor:cols,borderRadius:4,barPercentage:.78,categoryPercentage:.9}]},
    options:{indexAxis:'y',layout:{padding:{right:Math.min(170,16+longest*6.4)}},
      animation:{duration:_RM()?0:650,easing:'easeOutCubic'},
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>txt(c.parsed.x,c.dataIndex)}},
        datalabels:{anchor:'end',align:'end',offset:6,clamp:true,clip:false,color:()=>_ink(),font:{weight:'600',size:12},formatter:(v,c)=>txt(v,c.dataIndex)}},
      scales:{x:{position:'top',beginAtZero:true,title:{display:true,text:o.value,color:_mut()},ticks:{color:_mut()},grid:{color:'rgba(23,35,31,.06)'}},
              y:{ticks:{color:_mut(),autoSkip:false},grid:{display:false}}}}};
}
function _drillTo(id,i){
  const st=_DRILL[id]; if(!st) return;
  const el=document.getElementById(id), box=_boxOf(el), card=el.closest('.chart-card');
  if(i===st.active && i!==null) i=null;
  const swap=()=>{
    if(i===null){ st.active=null; const cfg=st.cfg; _drillRestoring=true; try{ mkChart(id,cfg); } finally{ _drillRestoring=false; } }
    else { st.active=i; const dc=_drillCfg(id,i);
      // give the breakdown enough height for its bars
      const need=Math.max(parseInt(el.dataset.h||'240',10), dc.data.labels.length*34+70);
      box.style.height=need+'px'; mkChart(id,dc); }
    if(i===null){ box.style.height=el.dataset.h?el.dataset.h+'px':box.style.height; }
    card.querySelectorAll('.drill-chip').forEach(b=>{ const on=+b.dataset.i===st.active; b.classList.toggle('on',on); b.setAttribute('aria-pressed',on?'true':'false'); });
    const hint=card.querySelector('.drill-hint');
    if(hint) hint.textContent=st.active===null?'Click a value to see its breakdown'
      :'Showing each '+st.o.part+' for '+st.o.cat.toLowerCase()+' "'+_drillCats(id)[st.active]+'". Click it again to go back.';
    _attachClick(id);
  };
  if(_RM()){ swap(); return; }
  box.classList.add('drill-out');
  setTimeout(()=>{ swap(); box.classList.remove('drill-out'); },180);
}
let _drillRestoring=false;
function _drillResetAll(){ Object.keys(_DRILL).forEach(id=>{ if(_DRILL[id].active!==null){ const s=_RM; _DRILL[id].active=null; const cfg=_DRILL[id].cfg; _drillRestoring=true; try{ mkChart(id,cfg); }finally{ _drillRestoring=false; }
  const card=document.getElementById(id).closest('.chart-card'); const el=document.getElementById(id); const box=_boxOf(el); if(box) box.style.height=(el.dataset.h||240)+'px';
  card.querySelectorAll('.drill-chip').forEach(b=>{ b.classList.remove('on'); b.setAttribute('aria-pressed','false'); });
  const h=card.querySelector('.drill-hint'); if(h) h.textContent='Click a value to see its breakdown'; _attachClick(id); } }); }

// One click handler per chart: x-axis labels drill down; points select a rule.
function _attachClick(id){
  const ch=Chart.getChart(id); if(!ch) return;
  const drilled=_DRILL[id]&&_DRILL[id].active!==null;
  ch.options.onHover=(e,els)=>{
    const c=e.native&&e.native.target; if(!c) return;
    c.style.cursor=(els&&els.length>0&&!drilled)?'pointer':'default';
  };
  // Chart.js only reports clicks inside the plot area, so the x-axis labels get their own listener
  const cv=ch.canvas;
  if(_DRILL[id]&&!cv.$axisClick){ cv.$axisClick=true;
    const axisHit=(ev)=>{ const c=Chart.getChart(cv); if(!c||!_DRILL[id]||_DRILL[id].active!==null) return null;
      const r=cv.getBoundingClientRect(), x=ev.clientX-r.left, y=ev.clientY-r.top, sx=c.scales.x;
      if(!sx||y<=c.chartArea.bottom||y>sx.bottom+4||x<c.chartArea.left-20||x>c.chartArea.right+20) return null;
      if(_DRILL[id].o.custom){ const v=sx.getValueForPixel(x); const k=_DRILL[id].bands.findIndex(b=>_inBand(b,v)); return k>=0?k:null; }
      const k=Math.round(sx.getValueForPixel(x)); return (k>=0&&k<(c.data.labels||[]).length)?k:null; };
    cv.addEventListener('click',ev=>{ const k=axisHit(ev); if(k!==null) _drillTo(id,k); });
    cv.addEventListener('mousemove',ev=>{ if(axisHit(ev)!==null) cv.style.cursor='pointer'; });
  }
  ch.options.onClick=(e,els)=>{
    if(drilled) return;
    if(_SEL_CHARTS[id]){ _ruleClick(id,els); return; }
    if(_DRILL[id]&&!_DRILL[id].o.custom&&els&&els.length) _drillTo(id,els[0].index);
  };
}

// ════════ 2. WHEEL ZOOM (scatter and bubble charts) ════════
function _zoomable(id){
  const cv=document.getElementById(id); if(!cv) return;
  const card=cv.closest('.chart-card');
  if(card&&!card.querySelector('.zoom-reset')){
    const b=document.createElement('button'); b.type='button'; b.className='tool-btn zoom-reset'; b.textContent='Reset zoom'; b.hidden=true;
    b.addEventListener('click',()=>_zoomReset(id)); _cardTools(card).appendChild(b); _syncToolPad(card);
  }
  const hint=card&&card.querySelector('.zoom-hint');
  if(card&&!hint){ const h=document.createElement('div'); h.className='zoom-hint'; h.textContent='Scroll or pinch over the chart to zoom. Double-click to reset.';
    const box=_boxOf(cv)||cv; box.insertAdjacentElement('afterend',h); }
  if(cv.$zoomBound) return; cv.$zoomBound=true;
  cv.addEventListener('wheel',e=>{
    const ch=Chart.getChart(cv); if(!ch||!ch.scales.x||ch.scales.x.type!=='linear'||(_DRILL[id]&&_DRILL[id].active!==null)) return;
    e.preventDefault();
    if(!ch.$zb) ch.$zb={x:[ch.scales.x.min,ch.scales.x.max],y:[ch.scales.y.min,ch.scales.y.max]};
    const f=Math.min(1.5,Math.max(0.67,Math.exp(e.deltaY*(e.ctrlKey?0.01:0.0015))));
    const r=cv.getBoundingClientRect(), px=e.clientX-r.left, py=e.clientY-r.top;
    ['x','y'].forEach(ax=>{
      const s=ch.scales[ax], o=ch.$zb[ax], c=s.getValueForPixel(ax==='x'?px:py);
      let mn=c-(c-s.min)*f, mx=c+(s.max-c)*f;
      const full=o[1]-o[0];
      if(mx-mn>=full){ mn=o[0]; mx=o[1]; }
      else { if(mn<o[0]){ mx+=o[0]-mn; mn=o[0]; } if(mx>o[1]){ mn-=mx-o[1]; mx=o[1]; } }
      ch.options.scales[ax].min=mn; ch.options.scales[ax].max=mx;
    });
    ch.update('none');
    const zoomed=ch.options.scales.x.min>ch.$zb.x[0]+1e-9||ch.options.scales.x.max<ch.$zb.x[1]-1e-9;
    const zb=card&&card.querySelector('.zoom-reset'); if(zb){ zb.hidden=!zoomed; _syncToolPad(card); }
  },{passive:false});
  cv.addEventListener('dblclick',()=>_zoomReset(id));
}
function _zoomReset(id){
  const ch=Chart.getChart(id); if(!ch||!ch.$zb) return;
  ['x','y'].forEach(ax=>{ ch.options.scales[ax].min=ch.$zb[ax][0]; ch.options.scales[ax].max=ch.$zb[ax][1]; });
  ch.update(_RM()?'none':undefined);
  const card=document.getElementById(id).closest('.chart-card'), zb=card&&card.querySelector('.zoom-reset'); if(zb){ zb.hidden=true; _syncToolPad(card); }
}

// ════════ 3. RULE SELECTION (bubble chart and lift scatter) ════════
const _SEL_CHARTS={'ch-arm-bubble':{ruleOf:p=>({a:p.antecedent,c:p.consequent})},'ch-arm-scatter':{ruleOf:p=>({a:p.ant,c:p.con})}};
const _SEL={};
function _ruleFind(a,c){ return (window._RULES||[]).find(r=>r.antecedent===a&&r.consequent===c)||null; }
function _pct(v,d){ return (v*100).toFixed(d==null?1:d)+'%'; }
function _ruleCalloutHtml(r){
  const m=window._RULES_META||{}, n=m.n_tx||window._nFinal||0;
  const q=r.q_value, sig=r.significant;
  const qs=q==null?'not available':(q<0.001?'below 0.001':q.toFixed(3));
  const more=r.lift>=1?Math.round((r.lift-1)*100)+'% more likely':Math.round((1-r.lift)*100)+'% less likely';
  return '<button type="button" class="rc-close" aria-label="Close">×</button>'
    +'<div class="rc-rule"><span>If</span> '+_escH(r.antecedent)+'<br><span>then</span> '+_escH(r.consequent)+'</div>'
    +'<p class="rc-plain">Of the '+(r.n_antecedent||Math.round(r.support/r.confidence*n)).toLocaleString()+' students who answered "'+_escH(r.antecedent)+'", '
    +Math.round(r.confidence*100)+' in 100 also answered "'+_escH(r.consequent)+'". That is '+more+' than among students in general.</p>'
    +'<dl class="rc-dl">'
    +'<div><dt>Support</dt><dd>'+_pct(r.support)+(r.n_both?' · '+r.n_both.toLocaleString()+' of '+n.toLocaleString()+' students give both answers':'')+'</dd></div>'
    +'<div><dt>Confidence</dt><dd>'+_pct(r.confidence)+' of students with the "if" answer also give the "then" answer</dd></div>'
    +'<div><dt>Lift</dt><dd>'+(+r.lift).toFixed(2)+' · '+more+' than chance</dd></div>'
    +'<div><dt>Significance</dt><dd>'+(sig==null?'Not tested':(sig?'Significant':'Not significant'))+' (adjusted p-value '+qs+')</dd></div>'
    +'</dl>';
}
function _ruleClick(id,els){
  const ch=Chart.getChart(id); if(!ch) return;
  const k=els&&els.length?els[0].index:null;
  if(k===null||k===(_SEL[id]&&_SEL[id].i)){ _ruleClear(id); return; }
  _ruleSelect(id,k);
}
function _ruleSelect(id,k){
  const ch=Chart.getChart(id), ds=ch.data.datasets[0];
  if(!_SEL[id]) _SEL[id]={bg:Array.isArray(ds.backgroundColor)?ds.backgroundColor.slice():ds.backgroundColor,bw:ds.borderWidth,bc:ds.borderColor};
  const s=_SEL[id]; s.i=k;
  if(s.tt===undefined) s.tt=ch.options.plugins.tooltip.enabled; ch.options.plugins.tooltip.enabled=false; ch.tooltip&&ch.tooltip.setActiveElements([],{x:0,y:0});
  const n=ds.data.length, base=s.bg;
  ds.backgroundColor=Array.from({length:n},(_,j)=>{ const c=Array.isArray(base)?base[j]:base; return j===k?c:_fade(c,0.12); });
  ds.borderColor=Array.from({length:n},(_,j)=>j===k?_ink():'rgba(0,0,0,0)');
  ds.borderWidth=Array.from({length:n},(_,j)=>j===k?2.5:0);
  ch.update(_RM()?'none':undefined);
  const p=ds.data[k], ro=_SEL_CHARTS[id].ruleOf(p), r=_ruleFind(ro.a,ro.c)||{antecedent:ro.a,consequent:ro.c,support:p.support!=null?p.support:p.x,confidence:p.y<=1?p.y:p.x,lift:p.lift!=null?p.lift:p.y};
  const box=_boxOf(document.getElementById(id));
  let co=box.querySelector('.rule-callout');
  if(!co){ co=document.createElement('div'); co.className='rule-callout'; co.setAttribute('role','dialog'); box.appendChild(co); }
  co.innerHTML=_ruleCalloutHtml(r);
  co.querySelector('.rc-close').addEventListener('click',()=>_ruleClear(id));
  co.classList.remove('show'); void co.offsetWidth; co.classList.add('show');
  _rulePlace(id);
}
function _rulePlace(id){
  const s=_SEL[id]; const ch=Chart.getChart(id); if(!s||s.i==null||!ch) return;
  const box=_boxOf(document.getElementById(id)), co=box&&box.querySelector('.rule-callout'); if(!co) return;
  const el=ch.getDatasetMeta(0).data[s.i]; if(!el) return;
  const a=ch.chartArea, inside=el.x>=a.left-2&&el.x<=a.right+2&&el.y>=a.top-2&&el.y<=a.bottom+2;
  co.style.visibility=inside?'visible':'hidden';
  const W=box.clientWidth, H=box.clientHeight, w=co.offsetWidth, h=co.offsetHeight, r=(el.options&&el.options.radius)||6;
  let left, side;
  if(W<520){ left=Math.max(4,Math.min(W-w-4,el.x-w/2)); side='below'; }
  else if(el.x+r+16+w<=W-4){ left=el.x+r+16; side='right'; }
  else { left=Math.max(4,el.x-r-16-w); side='left'; }
  let top=side==='below'?Math.min(H-h-4,el.y+r+12):Math.max(4,Math.min(H-h-4,el.y-h/2));
  co.style.left=left+'px'; co.style.top=top+'px'; co.dataset.side=side;
  co.style.setProperty('--arrow-y',Math.max(14,Math.min(h-14,el.y-top))+'px');
  co.style.setProperty('--arrow-x',Math.max(14,Math.min(w-14,el.x-left))+'px');
}
function _ruleClear(id){
  const s=_SEL[id], ch=Chart.getChart(id);
  if(s&&ch){ const ds=ch.data.datasets[0]; ds.backgroundColor=s.bg; ds.borderWidth=s.bw; ds.borderColor=s.bc; ch.options.plugins.tooltip.enabled=s.tt!==false; ch.update(_RM()?'none':undefined); }
  delete _SEL[id];
  const box=_boxOf(document.getElementById(id)), co=box&&box.querySelector('.rule-callout'); if(co) co.remove();
}

// ════════ 4. RADAR: isolate one cluster, compare several, callout lines ════════
let _radarSel=new Set();
function _radarInit(){
  const ch=Chart.getChart('ch-cluster-radar'); if(!ch) return;
  const card=document.getElementById('ch-cluster-radar').closest('.chart-card');
  ch.options.plugins.legend.onClick=(e,item)=>_radarPick(item.datasetIndex);
  ch.options.plugins.legend.labels.generateLabels=c=>{
    const items=Chart.defaults.plugins.legend.labels.generateLabels(c);
    items.forEach(it=>{ const on=!_radarSel.size||_radarSel.has(it.datasetIndex); it.hidden=false;
      if(!on){ it.fillStyle=_fade(it.fillStyle,0.25); it.strokeStyle=_fade(it.strokeStyle,0.25); it.fontColor=_fade(_mut(),0.55); } });
    return items;
  };
  ch.$radarBasePad=(ch.options.scales.r.pointLabels&&ch.options.scales.r.pointLabels.padding)||5;
  if(!card.querySelector('.radar-compare')){
    const wrap=document.createElement('div'); wrap.className='menu-wrap radar-compare';
    wrap.innerHTML='<button type="button" class="tool-btn" aria-haspopup="true" aria-expanded="false">Compare ▾</button><div class="tool-menu" hidden></div>';
    const tools=_cardTools(card); tools.prepend(wrap);
    const rb=document.createElement('button'); rb.type='button'; rb.className='tool-btn radar-reset'; rb.textContent='Show all'; rb.disabled=true;
    rb.addEventListener('click',()=>_radarApply(new Set())); tools.appendChild(rb);
    wrap.querySelector('button').addEventListener('click',ev=>{ ev.stopPropagation(); _menuToggle(wrap); });
    _syncToolPad(card);
    const hint=document.createElement('div'); hint.className='zoom-hint'; hint.textContent='Click a group in the legend to see it alone, or use Compare to pick several.';
    _boxOf(document.getElementById('ch-cluster-radar')).insertAdjacentElement('afterend',hint);
  }
  _radarMenu();
  _radarApply(_radarSel,true);
}
function _radarMenu(){
  const ch=Chart.getChart('ch-cluster-radar'); if(!ch) return;
  const m=document.querySelector('.radar-compare .tool-menu');
  m.innerHTML='<div class="tm-head">Show these groups</div>'+ch.data.datasets.map((d,i)=>'<label class="tm-row"><input type="checkbox" value="'+i+'"'+(_radarSel.has(i)?' checked':'')+'>'
    +'<span class="tm-sw" style="background:'+_escH(d.borderColor)+'"></span>'+_escH(d.label)+'</label>').join('')
    +'<button type="button" class="tm-link">Show all groups</button>';
  m.querySelectorAll('input').forEach(cb=>cb.addEventListener('change',()=>{
    const s=new Set([...m.querySelectorAll('input:checked')].map(x=>+x.value)); _radarApply(s); }));
  m.querySelector('.tm-link').addEventListener('click',()=>{ _radarApply(new Set()); _menuClose(); });
}
function _radarPick(i){ _radarApply(_radarSel.size===1&&_radarSel.has(i)?new Set():new Set([i])); }
function _radarApply(sel,quiet){
  const ch=Chart.getChart('ch-cluster-radar'); if(!ch) return;
  if(sel.size===ch.data.datasets.length) sel=new Set();
  _radarSel=new Set(sel);
  ch.data.datasets.forEach((d,i)=>ch.setDatasetVisibility(i,!_radarSel.size||_radarSel.has(i)));
  const n=_radarSel.size;
  ch.options.scales.r.pointLabels.padding=n?ch.$radarBasePad+18+n*20:ch.$radarBasePad;
  ch.update(quiet||_RM()?'none':undefined);
  document.querySelectorAll('.radar-compare input').forEach(cb=>cb.checked=_radarSel.has(+cb.value));
  const rb=document.querySelector('.radar-reset'); if(rb) rb.disabled=!n;
}
// draws a line from every corner of each selected polygon to its value
const _radarCallouts={id:'scRadarCallouts',afterDatasetsDraw(chart){
  if(chart.canvas.id!=='ch-cluster-radar'||!_radarSel.size) return;
  const r=chart.scales.r, ctx=chart.ctx, sel=[..._radarSel].sort((a,b)=>a-b);
  ctx.save(); ctx.font='600 11px '+(Chart.defaults.font.family||'sans-serif'); ctx.lineWidth=1.25;
  sel.forEach((di,k)=>{
    const meta=chart.getDatasetMeta(di); if(!meta||meta.hidden) return;
    const ds=chart.data.datasets[di], col=ds.borderColor;
    meta.data.forEach((pt,j)=>{
      const ang=r.getIndexAngle(j)-Math.PI/2, R=r.drawingArea+12+k*20;
      const ex=r.xCenter+Math.cos(ang)*R, ey=r.yCenter+Math.sin(ang)*R;
      ctx.strokeStyle=col; ctx.beginPath(); ctx.moveTo(pt.x,pt.y); ctx.lineTo(ex,ey); ctx.stroke();
      ctx.fillStyle=col; ctx.beginPath(); ctx.arc(ex,ey,2,0,Math.PI*2); ctx.fill();
      const c=Math.cos(ang), s=Math.sin(ang), v=String(Math.round(+ds.data[j]));
      ctx.textAlign=c>0.3?'left':(c<-0.3?'right':'center'); ctx.textBaseline=s>0.3?'top':(s<-0.3?'bottom':'middle');
      const tx=ex+c*5, ty=ey+s*5;
      ctx.lineWidth=3; ctx.strokeStyle=_surf(); ctx.strokeText(v,tx,ty); ctx.lineWidth=1.25;
      ctx.fillStyle=col; ctx.fillText(v,tx,ty);
    });
  });
  ctx.restore();
}};
const _rulePlacer={id:'scRulePlace',afterDraw(chart){ if(_SEL[chart.canvas.id]) _rulePlace(chart.canvas.id); }};

// ── small dropdown menus ──
let _openMenu=null;
function _menuToggle(wrap){
  if(_openMenu&&_openMenu!==wrap) _menuClose();
  const m=wrap.querySelector('.tool-menu,.lev-menu'), b=wrap.querySelector('button');
  const open=m.hidden; m.hidden=!open; b.setAttribute('aria-expanded',open?'true':'false'); _openMenu=open?wrap:null;
  if(open){ const f=m.querySelector('input'); if(f) f.focus(); }
}
function _menuClose(){ if(!_openMenu) return; const m=_openMenu.querySelector('.tool-menu,.lev-menu'); m.hidden=true; _openMenu.querySelector('button').setAttribute('aria-expanded','false'); _openMenu=null; }
document.addEventListener('click',e=>{ if(_openMenu&&!_openMenu.contains(e.target)) _menuClose(); });
document.addEventListener('keydown',e=>{ if(e.key==='Escape'){ if(_openMenu){ const b=_openMenu.querySelector('button'); _menuClose(); b.focus(); } else { Object.keys(_SEL).forEach(_ruleClear); } } });

// ════════ 5. STEPPED HEATMAP COLOURS ════════
// Correlations: three shades per direction, never a continuous gradient.
const _CORR_STEPS=[{lo:0.1,label:'weak'},{lo:0.3,label:'moderate'},{lo:0.5,label:'strong'}];
const _CORR_POS=['#bfe6d3','#5bbf96','#147a55'], _CORR_NEG=['#f5c8c4','#e47f78','#b3322b'], _CORR_NONE='rgba(120,130,125,.12)';
function _corrColor(v){
  if(v===null||v===undefined||isNaN(v)) return 'rgba(120,130,125,.05)';
  const a=Math.abs(v); if(a<0.1) return _CORR_NONE;
  const b=a<0.3?0:(a<0.5?1:2); return (v>0?_CORR_POS:_CORR_NEG)[b];
}
function _corrLegendHtml(){
  const sw=(c,t)=>'<span class="hm-key"><span class="hm-sw" style="background:'+c+'"></span>'+t+'</span>';
  return sw(_CORR_NEG[2],'Strong opposite (−0.5 or less)')+sw(_CORR_NEG[1],'Moderate opposite')+sw(_CORR_NEG[0],'Weak opposite')
    +sw(_CORR_NONE,'Little or no link (−0.1 to 0.1)')+sw(_CORR_POS[0],'Weak together')+sw(_CORR_POS[1],'Moderate together')+sw(_CORR_POS[2],'Strong together (0.5 or more)');
}
// Rule heatmap: four clearly different greens, cut at the quartiles of support.
const _ARM_STEPS=['#cfeedf','#7ccfa8','#2b9a6c','#0b5a3d'];
let _armCuts=[0,0,0];
function _armSetCuts(vals){
  const v=vals.filter(x=>x>0).sort((a,b)=>a-b);
  if(!v.length){ _armCuts=[0,0,0]; return; }
  const q=f=>v[Math.min(v.length-1,Math.floor(f*(v.length-1)))];
  _armCuts=[q(.25),q(.5),q(.75)];
}
function _armColor(v){ if(!(v>0)) return 'rgba(120,130,125,.07)'; const k=v<_armCuts[0]?0:(v<_armCuts[1]?1:(v<_armCuts[2]?2:3)); return _ARM_STEPS[k]; }
function _armLegendHtml(){
  const f=x=>(x*100).toFixed(1)+'%';
  const sw=(c,t)=>'<span class="hm-key"><span class="hm-sw" style="background:'+c+'"></span>'+t+'</span>';
  return '<span class="hm-cap">Support:</span>'+sw(_ARM_STEPS[0],'below '+f(_armCuts[0]))+sw(_ARM_STEPS[1],f(_armCuts[0])+'–'+f(_armCuts[1]))
    +sw(_ARM_STEPS[2],f(_armCuts[1])+'–'+f(_armCuts[2]))+sw(_ARM_STEPS[3],f(_armCuts[2])+' or more')+sw('rgba(120,130,125,.07)','no rule');
}
function _hmLegend(id,html){
  const el=document.getElementById(id); if(!el) return; const box=_boxOf(el)||el;
  let lg=box.parentNode.querySelector('.hm-legend[data-for="'+id+'"]');
  if(!lg){ lg=document.createElement('div'); lg.className='hm-legend'; lg.dataset.for=id; box.insertAdjacentElement('afterend',lg); }
  lg.innerHTML=html;
}

// ════════ 6. ASSOCIATION RULES EXPLAINED ════════
function _armExplain(){
  const el=document.getElementById('arm-explain'); if(!el) return;
  const m=window._RULES_META||{}, R=window._RULES||[], n=m.n_tx||window._nFinal||0;
  const nsig=R.filter(r=>r.significant).length;
  const f=x=>Math.round(x*100)+'%';
  el.innerHTML='<h4>What are association rules?</h4>'
   +'<p>An association rule says that students who gave one answer also tend to give another: "if A, then B". For example, "if a student spends over ₹600 a visit, then they want a membership". A rule describes a pattern in this survey. It does not prove that A causes B.</p>'
   +'<h4>What the three numbers mean</h4>'
   +'<dl class="ax-dl">'
   +'<div><dt>Support</dt><dd>The share of all '+n.toLocaleString()+' students who gave both answers. 12% means 12 in every 100 students. Only pairs given by at least '+f(m.min_support||0.1)+' of students were considered, because rarer patterns are too likely to be coincidence.</dd></div>'
   +'<div><dt>Confidence</dt><dd>Of the students who gave the "if" answer, the share who also gave the "then" answer. 80% means 8 in 10. Higher is more reliable. Rules here need at least '+f(m.min_confidence||0.45)+'.</dd></div>'
   +'<div><dt>Lift</dt><dd>How much more likely the "then" answer is among students with the "if" answer than among students in general. 1 means no link at all. 1.5 means 50% more likely; below 1 means less likely. Rules here need a lift of at least '+(m.min_lift||1.1).toFixed(2)+'.</dd></div>'
   +'</dl>'
   +'<h4>How to read lift</h4>'
   +'<table class="clean-table ax-tbl"><thead><tr><th>Lift</th><th>Meaning (a common rule of thumb)</th></tr></thead><tbody>'
   +'<tr><td>Below 1.1</td><td>Too close to chance to matter; not shown</td></tr><tr><td>1.1 to 1.3</td><td>Weak link</td></tr><tr><td>1.3 to 1.6</td><td>Moderate link</td></tr><tr><td>Above 1.6</td><td>Strong link</td></tr></tbody></table>'
   +'<h4>When is a rule statistically significant?</h4>'
   +'<p>Each pair of answers is tested with a chi-square test of independence. It asks: if the two answers were completely unrelated, how likely is a link at least this strong just by chance? That probability is the p-value. '
   +(m.n_tests?m.n_tests+' pairs were tested at once, so the p-values are adjusted (Benjamini–Hochberg) to keep false alarms to about 5%. ':'Because many pairs are tested at once, the p-values are adjusted (Benjamini–Hochberg) to keep false alarms to about 5%. ')
   +'A rule is marked significant when its adjusted p-value is below 0.05.</p>'
   +'<p>'+(R.length?nsig+' of the '+R.length+' rules shown are significant. ':'')
   +'With a survey this large, even weak links pass the test, so significance only says a link is real, not that it is large. Judge how much a rule matters by its lift and support.</p>';
}

// ════════ v56 PRICING ════════
Object.assign(CHART_GUIDE,{
  'ch-wtp-mem':{p:'Shows how many of the students who want a membership would still join as the monthly fee goes up. Use it to pick a fee band: the table under it shows which fee brings in the most money from the students who said they would join.',
    a:[[_H,'Monthly membership fee, at the start of each price band the survey asked about.'],[_V,'Share of interested students whose stated maximum is at or above that fee. The shaded band is the 95% range for that share.']]},
  'ch-wtp-visit':{p:'Shows how many students could pay each price for one visit, from the most they said they would spend. It shows where a price rise starts to lose many students.',
    a:[[_H,'Price of one visit, at the start of each budget band.'],[_V,'Share of all students whose budget reaches that price. The shaded band is the 95% range.']]},
  'ch-headroom':{p:'Compares what students spend now with the most they said they would spend. Students at their limit are the first to stop coming if prices rise.',
    a:[[_V,'What students spend on a visit now.'],[_H,'Share of those students, split by how much room their budget leaves above what they spend.']]},
  'ch-wtp-seg':{p:'Compares groups of interested students by how many would pay the middle membership fee or more. A big gap between groups points to a higher-priced tier for the group that pays more.',
    a:[[_V,'Groups of students, for the measure chosen above the chart.'],[_H,'Share of each group who would pay that fee or more. The thin line through each bar is its 95% range; the dashed line is the average for all interested students.']]}
});
const _INR=v=>'₹'+Math.round(v).toLocaleString('en-IN');
const _PCT=v=>(v*100).toFixed(1)+'%';
const _MINUS=v=>(v<0?'−':'')+Math.abs(v).toFixed(2);
function _elWords(e){
  if(e==null) return '';
  if(e>-0.8) return 'few students drop out, so revenue rises';
  if(e>=-1.2) return 'revenue stays about the same';
  return 'many students drop out, so revenue falls';
}
// thin 95% range lines on bar or point charts, and an optional dashed reference line
const scRangeBars={id:'scRangeBars',afterDatasetsDraw(ch,args,opts){
  const r=ch.$ranges; if(!r) return; const ctx=ch.ctx, ds=ch.getDatasetMeta(0); if(!ds||ds.hidden) return;
  const horiz=ch.options.indexAxis==='y', sc=horiz?ch.scales.x:ch.scales.y;
  ctx.save(); ctx.strokeStyle=_theme()==='dark'?'#e9efea':'#17231f'; ctx.lineWidth=1.5;
  ds.data.forEach((el,i)=>{ const rg=r[i]; if(!rg) return; const a=sc.getPixelForValue(rg[0]), b=sc.getPixelForValue(rg[1]);
    ctx.beginPath();
    if(horiz){ ctx.moveTo(a,el.y); ctx.lineTo(b,el.y); ctx.moveTo(a,el.y-5); ctx.lineTo(a,el.y+5); ctx.moveTo(b,el.y-5); ctx.lineTo(b,el.y+5); }
    else{ ctx.moveTo(el.x,a); ctx.lineTo(el.x,b); ctx.moveTo(el.x-5,a); ctx.lineTo(el.x+5,a); ctx.moveTo(el.x-5,b); ctx.lineTo(el.x+5,b); }
    ctx.stroke(); });
  if(ch.$refLine!=null){ const x=sc.getPixelForValue(ch.$refLine); ctx.setLineDash([5,4]); ctx.strokeStyle=_theme()==='dark'?'#a9b5ae':'#55635d';
    ctx.beginPath(); if(horiz){ ctx.moveTo(x,ch.chartArea.top); ctx.lineTo(x,ch.chartArea.bottom); } else { ctx.moveTo(ch.chartArea.left,x); ctx.lineTo(ch.chartArea.right,x); } ctx.stroke(); }
  ctx.restore(); }};

function _curveChart(id, cv, unit){
  const pts=cv.points, labels=pts.map(p=>_INR(p.price)+(p===pts[pts.length-1]&&cv.open_top?'+':''));
  mkChart(id,{type:'line',data:{labels,datasets:[
    {label:'95% range, low',data:pts.map(p=>+(p.lo*100).toFixed(1)),borderColor:'rgba(43,110,163,0)',backgroundColor:'rgba(43,110,163,0.16)',pointRadius:0,fill:false},
    {label:'95% range, high',data:pts.map(p=>+(p.hi*100).toFixed(1)),borderColor:'rgba(43,110,163,0)',backgroundColor:'rgba(43,110,163,0.16)',pointRadius:0,fill:'-1'},
    {label:'Share still willing',data:pts.map(p=>+(p.share*100).toFixed(1)),borderColor:'#2b6ea3',backgroundColor:'#2b6ea3',pointRadius:5,pointHoverRadius:7,borderWidth:2,fill:false}
  ]},plugins:[ChartDataLabels],options:{
    plugins:{legend:{display:false},
      datalabels:{display:c=>c.datasetIndex===2,align:c=>c.dataIndex===0?'right':'top',offset:6,color:'#17231F',font:{weight:'600'},formatter:v=>v.toFixed(0)+'%'},
      tooltip:{filter:i=>i.datasetIndex===2,callbacks:{label:c=>{const p=pts[c.dataIndex];return [`${(p.share*100).toFixed(1)}% (95% range ${(p.lo*100).toFixed(1)}–${(p.hi*100).toFixed(1)}%)`,`${p.students.toLocaleString()} of ${cv.n.toLocaleString()} students`];}}}},
    scales:{y:{min:0,max:110,ticks:{stepSize:20,callback:v=>v<=100?v+'%':''}},x:{grid:{display:false}}}}});
}
function _curveTable(id, cv, revLabel){
  const t=document.getElementById(id); if(!t) return;
  const best=cv.best_price, n=cv.points.length;
  const head='<thead><tr><th>Price</th><th>Still willing</th><th>'+revLabel+'</th><th>Elasticity to the next price</th><th>What that means</th></tr></thead>';
  const body=cv.points.map((p,i)=>{
    const e=cv.elasticity[i]; const isBest=p.price===best;
    const eTxt=e?`${_MINUS(e.e)}`+(e.lo!=null?`<span class="pr-sub"> 95% range ${_MINUS(e.lo)} to ${_MINUS(e.hi)}</span>`:''):'—';
    return `<tr${isBest?' class="pr-best"':''}><td>${_INR(p.price)}${i===n-1&&cv.open_top?' or more':''}</td>`
      +`<td>${_PCT(p.share)}<span class="pr-sub"> 95% range ${(p.lo*100).toFixed(1)}–${(p.hi*100).toFixed(1)}%</span></td>`
      +`<td>${_INR(p.rev100)}${isBest?'<span class="pr-sub"> highest</span>':''}</td><td>${eTxt}</td><td>${e?_elWords(e.e):'—'}</td></tr>`;
  }).join('');
  t.innerHTML=head+'<tbody>'+body+'</tbody>';
}
function _headroomChart(h){
  const step=Math.round(h.step), names=['At their limit',`Room for about ${_INR(step)} more`,`Room for about ${_INR(step*2)} or more`,'Both answers in the top band (can\'t tell)'];
  const cols=[_RAMP4[0],_RAMP4[1],_RAMP4[3],'#c5d0c8'];
  const labels=h.rows.map(r=>{ const t=String(r.spend), nums=t.match(/\d+/g)||[]; return 'Spending '+(/^>/.test(t)?'over ₹'+nums[0]:/^<|^≤/.test(t)?'up to ₹'+nums[0]:'₹'+nums.join('–'))+' now'; });
  mkChart('ch-headroom',{type:'bar',__semantic:true,data:{labels,datasets:names.map((nm,j)=>({label:nm,
      data:h.rows.map(r=>r.n?+(r.counts[j]/r.n*100).toFixed(1):0),raw:h.rows.map(r=>r.counts[j]),backgroundColor:cols[j],borderColor:'#ffffff',borderWidth:1}))},
    options:{indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.parsed.x}% (${c.dataset.raw[c.dataIndex].toLocaleString()} students)`}}},
      scales:{x:{stacked:true,min:0,max:100,ticks:{callback:v=>v+'%'}},y:{stacked:true,grid:{display:false}}}}});
  const s=h.shares;
  document.getElementById('headroom-read').textContent=
    `Of the students whose answers can be compared, ${(100*(s[1]+s[2])).toFixed(0)}% said they could spend more than they do now, so there is some room to raise prices. `
    +`${(100*s[0]).toFixed(0)}% are already at their limit: they are the first who would stop coming after a price rise.`
    +(h.unknown?` ${h.unknown.toLocaleString()} students spend and budget in the top band (over ₹${String(h.budget_bands[h.budget_bands.length-1]).replace(/\D/g,'')}), so the survey cannot show their room.`:'');
}
let _segKey=null;
function _segRender(){
  const S=(window._PRICING||{}).segments; if(!S||!S.groups.length) return;
  const g=S.groups.find(x=>x.key===_segKey)||S.groups[0]; _segKey=g.key;
  document.querySelectorAll('#seg-tabs button').forEach(b=>b.setAttribute('aria-pressed',b.dataset.key===g.key?'true':'false'));
  const thr=_INR(S.threshold)+' or more a month';
  document.getElementById('seg-sub').textContent=`Share of interested students who would pay ${thr}, by ${g.title.toLowerCase()}`;
  const ch=mkChart('ch-wtp-seg',{type:'bar',data:{labels:g.rows.map(r=>r.label),datasets:[{label:'Would pay '+thr,data:g.rows.map(r=>+(r.share*100).toFixed(1)),backgroundColor:'#2b6ea3',borderRadius:4,maxBarThickness:28}]},
    plugins:[scRangeBars],options:{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>{const r=g.rows[c.dataIndex];return `${(r.share*100).toFixed(1)}% of ${r.n.toLocaleString()} students (95% range ${(r.lo*100).toFixed(1)}–${(r.hi*100).toFixed(1)}%)`;}}}},
      scales:{x:{min:0,max:100,ticks:{callback:v=>v+'%'}},y:{grid:{display:false}}}}});
  if(ch){ ch.$ranges=g.rows.map(r=>[r.lo*100,r.hi*100]); ch.$refLine=S.overall*100; ch.update('none'); }
  const sorted=g.rows.slice().sort((a,b)=>b.share-a.share), hi=sorted[0], lo=sorted[sorted.length-1];
  const sig=g.p==null?'':(g.p<0.001?' The difference between the groups is statistically significant (p below 0.001).':
    g.p<0.05?` The difference between the groups is statistically significant (p = ${g.p.toFixed(3)}).`:` The difference between the groups could be chance (p = ${g.p.toFixed(2)}).`);
  document.getElementById('seg-read').textContent=`${hi.label}: ${(hi.share*100).toFixed(0)}% would pay ${thr}, against ${(lo.share*100).toFixed(0)}% for ${lo.label}. `
    +`On average, ${(S.overall*100).toFixed(0)}% of interested students would.`+sig;
  document.getElementById('tbl-wtp-seg').innerHTML='<thead><tr><th>'+g.title+'</th><th>Students</th><th>Would pay '+thr+'</th><th>95% range</th></tr></thead><tbody>'
    +g.rows.map(r=>`<tr><td>${r.label}</td><td>${r.n.toLocaleString()}</td><td>${_PCT(r.share)}</td><td>${(r.lo*100).toFixed(1)}–${(r.hi*100).toFixed(1)}%</td></tr>`).join('')+'</tbody>';
}
function buildPricing(){
  const P=window._PRICING, sec=document.getElementById('pricing-section'); if(!sec) return;
  if(!P||P.error||(!P.membership&&!P.visit)){ sec.style.display='none'; return; }
  sec.style.display='';
  const card=id=>{ const el=document.getElementById(id); return el&&el.closest('.chart-card'); };
  if(P.membership){ _curveChart('ch-wtp-mem',P.membership); _curveTable('tbl-wtp-mem',P.membership,'Income per 100 interested students'); }
  else if(card('ch-wtp-mem')) card('ch-wtp-mem').style.display='none';
  if(P.visit){ _curveChart('ch-wtp-visit',P.visit); _curveTable('tbl-wtp-visit',P.visit,'Takings per 100 students'); }
  else if(card('ch-wtp-visit')) card('ch-wtp-visit').style.display='none';
  if(P.headroom) _headroomChart(P.headroom); else if(card('ch-headroom')) card('ch-headroom').style.display='none';
  const S=P.segments, tabs=document.getElementById('seg-tabs');
  if(S&&S.groups.length&&tabs){
    tabs.innerHTML=S.groups.map(g=>`<button type="button" class="tool-btn" data-key="${g.key}" aria-pressed="false">${g.title}</button>`).join('');
    tabs.querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>{ _segKey=b.dataset.key; _segRender(); }));
    _segRender();
  } else if(card('ch-wtp-seg')) card('ch-wtp-seg').style.display='none';
}

// ── What-if: profit for the year, price against one other setting ──
let _sgKey='catchment', _sgData=null;
function _inrShort(v){ const a=Math.abs(v), s=v<0?'−':'';
  if(a>=1e7) return s+'₹'+(a/1e7).toFixed(2)+' cr'; if(a>=1e5) return s+'₹'+(a/1e5).toFixed(1)+' L'; return s+'₹'+Math.round(a).toLocaleString('en-IN'); }
function _sgRender(d){
  if(d) _sgData=d.sens_grid; const G=_sgData; if(!G) return;
  const tabs=document.getElementById('wi-sgrid-tabs');
  if(tabs&&!tabs.children.length){
    tabs.innerHTML=Object.entries(G.grids).map(([k,g])=>`<button type="button" class="tool-btn" data-key="${k}" aria-pressed="false">Price and ${g.title.toLowerCase()}</button>`).join('');
    tabs.querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>{ _sgKey=b.dataset.key; _sgRender(); }));
  }
  const g=G.grids[_sgKey]||G.grids.catchment;
  tabs&&tabs.querySelectorAll('button').forEach(b=>b.setAttribute('aria-pressed',b.dataset.key===_sgKey?'true':'false'));
  const all=g.cells.flat(), pos=all.filter(v=>v>=0).sort((a,b)=>a-b), neg=all.filter(v=>v<0).map(Math.abs).sort((a,b)=>a-b);
  const tier=(arr,v)=>{ if(arr.length<3) return 2; const r=arr.indexOf(v)/(arr.length-1); return r>0.66?3:r>0.33?2:1; };
  const cls=v=>v>=0?'sg-p'+tier(pos,v):'sg-n'+tier(neg,Math.abs(v));
  const colBest=g.cols.map((_,j)=>g.cells.reduce((b,row,i)=>row[j]>g.cells[b][j]?i:b,0));
  let h=`<caption>Rows: price change. Columns: ${g.title.toLowerCase()}.</caption>`+'<thead><tr><th scope="col">Price change</th>'+g.cols.map((c,j)=>`<th scope="col">${c}${j===g.current_col?' (now)':''}</th>`).join('')+'</tr></thead><tbody>';
  g.cells.forEach((row,i)=>{
    h+=`<tr><th scope="row">${G.rows[i]}${i===G.current_row?' (now)':''}</th>`+row.map((v,j)=>{
      const now=i===G.current_row&&j===g.current_col, best=colBest[j]===i;
      return `<td class="${cls(v)}${now?' sg-now':''}${best?' sg-best':''}">${_inrShort(v)}${now?'<span class="sr-only"> (current settings)</span>':''}</td>`; }).join('')+'</tr>';
  });
  document.getElementById('wi-sgrid').innerHTML=h+'</tbody>';
  document.getElementById('wi-sgrid-sub').textContent=`Profit for the year in each cell; rows change the price, columns change ${g.title.toLowerCase()}. Everything else stays as you entered it.`;
  const cc=g.current_col, cr=G.current_row;
  let read='';
  if(cc!=null){ const bi=colBest[cc], bv=g.cells[bi][cc], nv=g.cells[cr][cc];
    read=(bi===cr?`At your current ${g.title.toLowerCase()}, your current price earns the most of the prices tested (${_inrShort(nv)} for the year).`
      :`At your current ${g.title.toLowerCase()}, a price change of ${G.rows[bi]} earns the most (${_inrShort(bv)} for the year, against ${_inrShort(nv)} now).`);
    const lossCols=g.cols.filter((_,j)=>g.cells[cr][j]<0); if(lossCols.length) read+=` At the current price, the café makes a loss when ${g.title.toLowerCase()} is ${lossCols.join(', ')}.`;
  }
  const same=(j,k)=>g.cells.every(row=>Math.abs(row[j]-row[k])<=Math.max(1,Math.abs(row[j]))*0.005);
  if(cc!=null&&cc<g.cols.length-1&&same(cc,cc+1)&&_sgKey!=='seats')
    read+=` More ${g.title.toLowerCase()} than now would not add profit, because the café is already full: the seats, not demand, limit what it can take in. Compare price with seats to see what more room would earn.`;
  if(colBest.some(b=>b===g.cells.length-1)) read+=' Where the best figure is in the last row, it is the highest price tested; a higher price could earn more or less.';
  read+=' The outlined cell is your current settings; the bold figure in each column is the best price for that column.'
    +(_sgKey==='seats'?' Fixed costs stay as you entered them, though more seats would usually cost more.':'');
  document.getElementById('wi-sgrid-read').textContent=read;
  document.getElementById('wi-sgrid-key').textContent='Green cells make a profit and red cells a loss; darker means larger. L = lakh (1,00,000), cr = crore (1,00,00,000).';
}

// ════════ 7. PREDICTIVE LEVERS ════════
const _LEV={time:{label:'time slot',chart:'ch-lev-time'},spend:{label:'spend per visit',chart:'ch-lev-spend'},seat:{label:'seat type',chart:'ch-lev-seat'}};
let _levEx={time:[],spend:[],seat:[]}, _levOpts=null, _levBase=null, _levBusy=false, _levTimer=null, _levPending=null;
const _TIME_COL={Morning:'#cfae2a',Afternoon:'#d2692c',Evening:'#a8508e','Late Night':'#2b6ea3'};
const _LPAL=['#2b6ea3','#d2692c','#2f8f6a','#cfae2a','#a8508e','#6a86d8','#c4534a','#7f9a24'];
function _spendKey(s){ const m=String(s).match(/\d+/); return m?(+m[0])+(/^>/.test(s)?1:0):1e9; }
function _levOrder(kind,arr){
  const a=arr.slice();
  if(kind==='time'){ const o=['Morning','Afternoon','Evening','Late Night']; a.sort((x,y)=>(o.indexOf(x)<0?99:o.indexOf(x))-(o.indexOf(y)<0?99:o.indexOf(y))); }
  if(kind==='spend') a.sort((x,y)=>_spendKey(x)-_spendKey(y));
  return a;
}
function _entCol(kind,label){
  if(kind==='time'&&_TIME_COL[label]) return _TIME_COL[label];
  if(kind==='seat'&&_ENT.seat[label]) return _ENT.seat[label];
  if(kind==='spend'){ const k=_normCat(label), i=_ORD.spend.indexOf(k); if(i>=0) return _RAMP4[i]; }
  const all=_levOpts&&_levOpts[kind]?_levOrder(kind,_levOpts[kind]):[];
  const i=all.indexOf(label); return _LPAL[(i<0?all.length:i)%_LPAL.length];
}
function _levActive(){ return Object.values(_levEx).some(a=>a.length); }
let _levFull=null;
function _levToggle(kind,label){
  if(!_levOpts) return;
  const opts=_levOpts[kind]||[], on=!_levEx[kind].includes(label);
  if(on&&opts.length-_levEx[kind].length<=1){ _levSay('Keep at least one option ticked in each chart.','err'); return; }
  _levEx[kind]=on?_levEx[kind].concat([label]):_levEx[kind].filter(x=>x!==label);
  _levMenuSync(kind); _levPaint(); _levApply();
}
function _levSay(t,cls){ const st=document.getElementById('lev-status'); if(st){ st.textContent=t; st.className='lev-status'+(cls?' '+cls:''); } }
function _levPaint(){
  if(!_levOpts||!_levFull) return;
  const F=_levFull, src={time:(F.peak_hours||{}).distribution||{},spend:(F.revenue||{}).distribution||{},seat:(F.seating||{}).overall_pct||{}};
  const dark=_theme()==='dark';
  Object.keys(_LEV).forEach(kind=>{
    const labels=_levOrder(kind,Object.keys(src[kind]));
    const off=l=>_levEx[kind].includes(l);
    const cols=labels.map(l=>{ const c=_mapColor(_entCol(kind,l),dark); return off(l)?_fade(c,0.16):c; });
    const id=_LEV[kind].chart, ch=Chart.getChart(id);
    if(ch&&ch.$levKind===kind&&ch.data.labels.length===labels.length){
      ch.data.datasets[0].backgroundColor=cols; ch.update(_RM()?'none':undefined); return; }
    const disp=kind==='spend'?(l=>'₹'+l):(l=>l);
    const c2=mkChart(id,{type:'doughnut',data:{labels:labels.map(disp),datasets:[{data:labels.map(l=>src[kind][l]||0),backgroundColor:cols,borderWidth:2,borderColor:_surf(),hoverOffset:8}]},
      options:{animation:{duration:_RM()?0:500},
        onClick:(e,els)=>{ if(els&&els.length) _levToggle(kind,labels[els[0].index]); },
        onHover:(e,els)=>{ const t=e.native&&e.native.target; if(t) t.style.cursor=(els&&els.length)?'pointer':'default'; },
        plugins:{legend:{onClick:(e,item)=>_levToggle(kind,labels[item.index]),
            labels:{padding:12,boxWidth:10,generateLabels:c=>{ const it=Chart.overrides.doughnut.plugins.legend.labels.generateLabels(c);
              it.forEach((x,k)=>{ x.hidden=false; if(off(labels[k])) x.fontColor=_fade(_mut(),0.45); }); return it; }}},
          tooltip:{callbacks:{label:c=>c.label+': '+(+c.parsed).toFixed(1)+'% of all responses'+(off(labels[c.dataIndex])?' (not in the drill-down)':'')}}}}});
    if(c2) c2.$levKind=kind;
  });
}
function _levRender(d){
  if(d.lever_options&&!_levOpts) _levOpts=d.lever_options;
  if(!d.levers) _levFull=d;
  _levPaint();
  Object.keys(_LEV).forEach(_levMenu);
  const lv=d.levers;
  if(_levActive()&&lv){
    const parts=Object.keys(_LEV).filter(k=>_levEx[k].length).map(k=>{
      const kept=_levOrder(k,_levOpts[k]||[]).filter(v=>!_levEx[k].includes(v));
      return kept.map(v=>k==='spend'?'₹'+v:v).join(', ')+' ('+_LEV[k].label+')'; });
    const share=lv.n_full?Math.round(lv.n/lv.n_full*100):100;
    const st=document.getElementById('lev-status');
    if(st){ st.innerHTML='<strong>Drilled down to:</strong> '+_escH(parts.join('; '))+'. The charts below describe these '+lv.n.toLocaleString()+' of '+lv.n_full.toLocaleString()
      +' students ('+share+'%). The predictive models are still trained on all responses, so nothing is removed from the analysis. Visitor and revenue projections show the part of a 1,000-student catchment these students make up.'
      +(lv.n<100?' This is a small group, so treat its percentages with care.':'');
      st.className='lev-status on'; }
  } else _levSay('Showing all responses. Untick an option under Select, or click a slice, to drill down to particular groups of students.');
  const rb=document.getElementById('lev-reset'); if(rb) rb.disabled=!_levActive();
}
function _levMenu(kind){
  const wrap=document.querySelector('.lev-menu-wrap[data-kind="'+kind+'"]'); if(!wrap||!_levOpts) return;
  const m=wrap.querySelector('.lev-menu'), opts=_levOrder(kind,_levOpts[kind]||[]);
  if(m.querySelectorAll('input').length===opts.length){ _levMenuSync(kind); return; }
  const kept=opts.filter(o=>!_levEx[kind].includes(o)).length;
  m.innerHTML='<div class="tm-head">Include these</div>'+opts.map(o=>{ const on=!_levEx[kind].includes(o);
    return '<label class="tm-row"><input type="checkbox" value="'+_escH(o)+'"'+(on?' checked':'')+(on&&kept===1?' disabled title="At least one must stay ticked"':'')+'>'
      +'<span class="tm-sw" style="background:'+_entCol(kind,o)+'"></span>'+_escH(kind==='spend'?'₹'+o:o)+'</label>'; }).join('');
  m.querySelectorAll('input').forEach(cb=>cb.addEventListener('change',()=>{
    const v=cb.value; const prev=_levEx[kind].slice();
    _levEx[kind]=cb.checked?_levEx[kind].filter(x=>x!==v):_levEx[kind].concat([v]);
    _levMenuSync(kind); _levPaint(); _levApply();
  }));
}
function _levMenuSync(kind){
  const ins=[...document.querySelectorAll('.lev-menu-wrap[data-kind="'+kind+'"] input')];
  const kept=ins.filter(i=>!_levEx[kind].includes(i.value)).length;
  ins.forEach(i=>{ const on=!_levEx[kind].includes(i.value); i.checked=on; i.disabled=on&&kept===1; i.title=i.disabled?'At least one must stay ticked':''; });
}
let _levSeq=0;
async function _levApply(){
  const seq=++_levSeq;
  const ex={}; Object.keys(_levEx).forEach(k=>{ if(_levEx[k].length) ex[k]=_levEx[k].slice(); });
  const slow=setTimeout(()=>{ if(seq===_levSeq) _levSay('Updating…'); },350);
  try{
    const res=await fetch(BASE+'/predictive_data'+(Object.keys(ex).length?'?exclude='+encodeURIComponent(JSON.stringify(ex)):''));
    const d=await res.json();
    if(seq!==_levSeq) return;                         // a newer click is already on its way
    if(d.error){ _levSay(d.error,'err'); return; }
    renderPredictive(d);
  }catch(e){ if(seq===_levSeq) _levSay('Could not update: '+e.message,'err'); }
  finally{ clearTimeout(slow); }
}
let _levPrepTimer=null;
async function _levPrepPoll(){
  clearTimeout(_levPrepTimer);
  const el=document.getElementById('lev-prep'); if(!el) return;
  try{
    const p=await (await fetch(BASE+'/lever_status')).json();
    if(p.state==='running'){ el.textContent='Preparing every drill-down combination in advance: '+(p.done||0).toLocaleString()+' of '+(p.total||0).toLocaleString()+' ready. The levers already respond straight away.'; _levPrepTimer=setTimeout(_levPrepPoll,2000); }
    else if(p.state==='done'){ el.textContent='All '+(p.total||0).toLocaleString()+' drill-down combinations are ready.'; }
    else { el.textContent=''; _levPrepTimer=setTimeout(_levPrepPoll,3000); }
  }catch(e){ _levPrepTimer=setTimeout(_levPrepPoll,4000); }
}
function levReset(){ _levEx={time:[],spend:[],seat:[]}; Object.keys(_LEV).forEach(_levMenuSync); _levPaint(); _levApply(); }
function levGoto(){ const s=document.getElementById('pred-levers'); if(s) s.scrollIntoView({behavior:_RM()?'auto':'smooth',block:'start'}); const h=s&&s.querySelector('.section-title'); if(h){ h.tabIndex=-1; setTimeout(()=>h.focus({preventScroll:true}),_RM()?0:450); } }
// notes where a chart cannot meaningfully change, or did not change
const _PRED_CHARTS=['ch-time-dist','ch-slot-rev','ch-spend-by-slot','ch-peak-fi','ch-slot-foot','ch-weekly-pattern','ch-foot-dist','ch-spend-dist','ch-rev-slot','ch-rev-projection','ch-spend-univ','ch-seat-overall','ch-seat-by-time','ch-mem-models','ch-mem-spend','ch-mem-dist','ch-mem-fi'];
const _SEC_ANCHOR={peak_hours:'ch-time-dist',footfall:'ch-slot-foot',revenue:'ch-spend-dist',seating:'ch-seat-overall',membership:'ch-mem-models',model_summary:'ch-mem-models'};
function _chartSig(id){ const ch=Chart.getChart(id); if(!ch) return null; try{ return JSON.stringify([ch.data.labels,ch.data.datasets.map(d=>d.data)]); }catch(e){ return null; } }
function _levNotes(d){
  document.querySelectorAll('#pred-body .lever-note').forEach(n=>n.remove());
  if(!_levActive()){ _levBase={}; _PRED_CHARTS.forEach(id=>_levBase[id]=_chartSig(id)); return; }
  const notes=d.lever_notes||{}, done=new Set();
  const put=(target,text,inside)=>{ if(!target) return; const p=document.createElement('p'); p.className='lever-note'; p.textContent=text;
    if(inside) target.appendChild(p); else target.insertAdjacentElement('afterend',p); };
  Object.keys(notes).forEach(k=>{
    if(k.startsWith('sec:')){ const a=document.getElementById(_SEC_ANCHOR[k.slice(4)]); const sec=a&&a.closest('.dash-section'); const h=sec&&(sec.querySelector('.sec-head')||sec.querySelector('.section-title')); put(h,notes[k],false); return; }
    const el=document.getElementById(k); if(!el) return; done.add(k);
    const card=el.closest('.chart-card'); if(card) put(card,notes[k],true); else put(el,notes[k],true);
  });
  if(_levBase) _PRED_CHARTS.forEach(id=>{
    if(done.has(id)) return; const s=_chartSig(id); if(s&&_levBase[id]&&s===_levBase[id]){
      const card=document.getElementById(id).closest('.chart-card');
      put(card,'No change: drilling down to these students does not change the numbers behind this chart, so it matches the original.',true); }
  });
}
function _levGotoButtons(){
  document.querySelectorAll('#pred-body > .dash-section:not(#pred-levers) > .section-title').forEach(t=>{
    if(!/peak hours|footfall|revenue|seating|membership/i.test(t.textContent)) return;
    const w=document.createElement('div'); w.className='sec-head'; t.parentNode.insertBefore(w,t); w.appendChild(t);
    const b=document.createElement('button'); b.type='button'; b.className='lev-goto'; b.textContent='Go to levers ↑'; b.addEventListener('click',levGoto); w.appendChild(b);
  });
  document.querySelectorAll('.lev-menu-wrap > button').forEach(b=>b.addEventListener('click',ev=>{ ev.stopPropagation(); _menuToggle(b.parentElement); }));
}

// ── reset everything interactive (used before exporting a report) ──
function _interactResetAll(){
  try{ if(_PS.period!=='annual'){ _PS.period='annual'; _PS.idx=0; _psRender(); } }catch(e){}
  try{ if(_armHmAll){ _armHmAll=false; if(Chart.getChart('ch-arm-heatmap')) buildArmChart('heatmap'); } }catch(e){}
  try{ _drillResetAll(); }catch(e){}
  try{ Object.keys(_SEL).forEach(_ruleClear); }catch(e){}
  try{ ['ch-scatter2d','ch-arm-bubble','ch-arm-scatter'].forEach(_zoomReset); }catch(e){}
  try{ if(_radarSel.size) _radarApply(new Set(),true); }catch(e){}
  try{ if(window._sc3d){ window._sc3d.reset(); } }catch(e){}
  _menuClose();
}

// ── hooks run after a chart is built ──
_INIT['ch-cluster-radar']=()=>_radarInit();
_INIT['ch-scatter2d']=()=>_zoomable('ch-scatter2d');
_INIT['ch-arm-bubble']=()=>{ _zoomable('ch-arm-bubble'); _attachClick('ch-arm-bubble'); };
_INIT['ch-arm-scatter']=()=>{ _zoomable('ch-arm-scatter'); _attachClick('ch-arm-scatter'); };
function _afterChart(id,cfg){
  if(cfg&&cfg.__drilled){ return; }
  _CFG[id]=cfg;
  if(_DRILL_OPTS[id]&&!_drillRestoring) _drillInit(id,cfg);
  else if(_DRILL_OPTS[id]) _attachClick(id);
  if(_INIT[id]) try{ _INIT[id](); }catch(e){ console.warn('init '+id,e); }
}
document.addEventListener('DOMContentLoaded',()=>{
  try{ Chart.register(_radarCallouts,_rulePlacer); }catch(e){}
  _levGotoButtons();
});

function _hmSplit(s){ s=String(s); const i=s.indexOf(': '); return i<0?['',s]:[s.slice(0,i),s.slice(i+2)]; }
function _hmCut(s,n){ return s.length>n?s.slice(0,Math.max(1,n-1)).trimEnd()+'…':s; }
function _hmRowLabel(s,n){ const [q,a]=_hmSplit(s); if(!q) return _hmCut(a,n); if(q.length+a.length+2<=n) return q+': '+a; return _hmCut(q,Math.max(8,n-a.length-2))+': '+a; }
function _hmColLabel(s){ const [q,a]=_hmSplit(s); return q?[_hmCut(q,20),a]:[_hmCut(a,20)]; }
function _qFmt(q){ return q==null?'—':(q<0.001?'<0.001':(+q).toFixed(3)); }

function renderPredictive(d){
  if(d.lever_options&&!_levOpts){ _levOpts=d.lever_options; setTimeout(_levPrepPoll,500); }
  _renderPredictiveCore(d);
  try{ _levRender(d); _levNotes(d); }catch(e){ console.warn(e); }
}

// ══ ARM ══
let armSorted=[];
let armAllSorted=[];
let armViewExpanded=false;
let armChartBuilt={bubble:false,scatter:false,bar:false,heatmap:false};
// ════════════════════════════════════════════════════════════
//  v53 — SURFACE / DEEP VIEW
//  Surface (default): what an owner needs to decide how to run the café.
//  Deep: every analysis technique, as before. Works like the Light/Dark switch.
// ════════════════════════════════════════════════════════════
function _view(){ return document.documentElement.getAttribute('data-view')==='deep'?'deep':'surface'; }
// charts that only appear in the deep view (methods rather than decisions)
const _DEEP_CHARTS=['ch-hours-freq','ch-univ-budget','ch-elbow','ch-scatter2d','ch-scatter3d','ch-cluster-age-mem','ch-cluster-age-dist',
  'ch-peak-fi','ch-weekly-pattern','ch-mem-models'];
// whole sections that only appear in the deep view (matched on their title)
const _DEEP_SECTIONS=[/^Correlation Heatmap/i,/^Regression Analysis/i,/^Model Performance Summary/i];
// friendlier section titles for the surface view
const _SURFACE_TITLES=[[/^Data Summary — Distributions/i,'What students want'],[/^EDA — Cross-Tabulations/i,'Patterns that shape the layout and hours'],
  [/^K-Means Clustering/i,'Customer groups'],[/^Association rules/i,'What tends to go together']];
function _markViews(){
  _DEEP_CHARTS.forEach(id=>{ const el=document.getElementById(id); const c=el&&el.closest('.chart-card'); if(c) c.classList.add('deep-only'); });
  ['pca-explain','arm-note'].forEach(id=>{ const el=document.getElementById(id); if(el) el.classList.add('deep-only'); });
  const cm=document.getElementById('mem-conf-matrix'); if(cm&&cm.closest('.chart-card')) cm.closest('.chart-card').classList.add('deep-only');
  document.querySelectorAll('#s-analytics .dash-section, #s-predictive .dash-section').forEach(sec=>{
    const t=sec.querySelector('.section-title'); if(!t) return;
    const txt=t.textContent.trim();
    if(_DEEP_SECTIONS.some(r=>r.test(txt))) sec.classList.add('deep-only');
    const st=_SURFACE_TITLES.find(p=>p[0].test(txt)); if(st&&!t.dataset.deepTitle){ t.dataset.deepTitle=txt; t.dataset.surfaceTitle=st[1]; }
  });
  // association rules: the surface view shows only the "worth acting on" card
  const armSec=document.getElementById('arm-table-view')&&document.getElementById('arm-table-view').closest('.dash-section');
  if(armSec){ armSec.querySelectorAll('.tab-strip,#arm-table-view,#arm-bubble-view,#arm-scatter-view,#arm-bar-view,#arm-heatmap-view').forEach(e=>e.classList.add('deep-only')); }
  // the data-cleaning summary is for the analyst, not the owner
  document.querySelectorAll('[data-screen="s-summary"], .nav-tile[onclick*="openSummary"]').forEach(e=>e.classList.add('deep-only'));
  // a note at the top of the analysis screens in the surface view
  ['s-analytics','s-predictive'].forEach(sid=>{
    const body=document.querySelector('#'+sid+' .dash-body, #'+sid+' .pred-body')||document.getElementById(sid); if(!body||body.querySelector('.view-banner')) return;
    const b=document.createElement('div'); b.className='view-banner surface-only';
    b.innerHTML='<p><strong>Surface view.</strong> Showing what matters for deciding how to run the café. The full set of analysis techniques is in the deep view.</p>'
      +'<button type="button" class="tool-btn" onclick="setView(\'deep\')">Show deep view</button>';
    body.prepend(b);
  });
}
function setView(v,persist){
  v=v==='deep'?'deep':'surface';
  document.documentElement.setAttribute('data-view',v);
  if(persist!==false){ try{ localStorage.setItem('sc-view',v); }catch(e){} }
  document.querySelectorAll('[data-view-btn]').forEach(b=>b.setAttribute('aria-pressed',b.dataset.viewBtn===v?'true':'false'));
  document.querySelectorAll('.section-title[data-deep-title]').forEach(t=>{ t.textContent=v==='deep'?t.dataset.deepTitle:t.dataset.surfaceTitle; });
  // charts that were hidden need their size again
  requestAnimationFrame(()=>{ Object.values(Chart.instances).forEach(c=>{ try{ c.resize(); }catch(e){} }); if(window._sc3d) try{ window._sc3d.draw(); }catch(e){} });
  // if the summary screen is open and is now hidden, go home
  if(v==='surface'){ const a=document.querySelector('.screen.active'); if(a&&a.id==='s-summary') goTo('s-nav'); }
}
document.addEventListener('DOMContentLoaded',()=>{ _markViews(); setView(_view(),false); });

// ── Rules worth acting on (both views; the only rules view in the surface view) ──
function _ruleAction(con){
  const i=String(con).indexOf(': '), col=(i<0?'':con.slice(0,i)).toLowerCase(), val=i<0?con:con.slice(i+2), vl=val.toLowerCase();
  const yes=/^(yes|1|true)$/i.test(val), no=/^(no|0|false)$/i.test(val);
  if(/member/.test(col)) return yes?'Offer membership to these students first, for example at the counter or when they sign up for WiFi.':(no?'Do not push membership to this group; pay-per-visit offers suit them better.':'Use this when setting membership prices for this group.');
  if(/late|night/.test(col)) return yes?'Plan late opening hours around this group.':'This group will not use late hours; do not open late for them.';
  if(/seat/.test(col)) return 'Provide enough '+vl+' seating for this group.';
  if(/bever|drink|coffee/.test(col)) return 'Stock and promote '+vl+' for this group.';
  if(/food|snack|meal/.test(col)) return 'Put '+vl+' on the menu for this group.';
  if(/noise|quiet/.test(col)) return 'Give this group a '+vl+' zone.';
  if(/time|slot/.test(col)) return 'Expect this group in the '+vl+'; staff and open for it.';
  if(/frequen|visit/.test(col)) return 'These are likely regulars; a loyalty card or punch card would suit them.';
  if(/spend|budget|price/.test(col)) return 'Price combos and offers for this group around '+val+'.';
  if(/wifi|power|outlet|tool/.test(col)) return 'Make sure this facility is reliable where this group sits.';
  if(/group/.test(col)) return 'Make group tables bookable for this group.';
  if(/distance|campus|univ/.test(col)) return 'Use this to decide where to advertise.';
  return 'Keep this pattern in mind when planning offers for this group.';
}
function _ownerRules(){
  // technical summary tiles belong to the deep view
  document.querySelectorAll('#kpi-strip .mini-kpi').forEach(k=>{ if(/association rules|customer groups \(k-means\)/i.test(k.textContent)) k.classList.add('deep-only'); });
  const box=document.getElementById('owner-rules'); if(!box) return;
  const R=(window._RULES||[]), M=window._RULES_META||{}, N=M.n_tx||window._nFinal||0;
  if(!R.length){ box.querySelector('.or-body').innerHTML='<p class="or-empty">No rules were strong enough to report for this dataset.</p>'; return; }
  // practical relevance: how many students follow the rule beyond what chance alone would give
  const beyond=r=>(r.n_both||Math.round(r.support*N))*(1-1/Math.max(r.lift,1e-9));
  let pool=R.filter(r=>r.significant!==false&&r.lift>=1.2);
  if(pool.length<3) pool=R.filter(r=>r.significant!==false);
  if(!pool.length) pool=R.slice();
  // at most two rules per outcome, so one strong pattern does not crowd out the rest
  const seen={}, top=[];
  pool.slice().sort((a,b)=>beyond(b)-beyond(a)).forEach(r=>{ if(top.length<6&&(seen[r.consequent]||0)<2){ top.push(r); seen[r.consequent]=(seen[r.consequent]||0)+1; } });
  const strength=l=>l>=1.6?'Strong':(l>=1.3?'Moderate':'Weak');
  const q=v=>v==null?'not tested':(v<0.001?'below 0.001':(+v).toFixed(3));
  box.querySelector('.or-body').innerHTML='<div class="tbl-wrap"><table class="clean-table or-table"><thead><tr><th>Pattern</th><th>Students</th><th>Strength</th><th>Significant?</th><th>What the café can do</th></tr></thead><tbody>'
    +top.map(r=>{ const n=r.n_both||Math.round(r.support*N), na=r.n_antecedent||Math.round(r.support/r.confidence*N);
      return '<tr><td><div class="or-rule"><span>If</span> '+_escH(r.antecedent)+'<br><span>then</span> '+_escH(r.consequent)+'</div>'
        +'<div class="or-plain">'+Math.round(r.confidence*100)+'% of the '+na.toLocaleString()+' students with the first answer give the second; that is '+(+r.lift).toFixed(2)+'× the rate among all students.</div></td>'
        +'<td>'+n.toLocaleString()+'<div class="or-plain">about '+Math.round(beyond(r)).toLocaleString()+' more than chance</div></td>'
        +'<td>'+strength(r.lift)+'<div class="or-plain">lift '+(+r.lift).toFixed(2)+'</div></td>'
        +'<td>'+(r.significant==null?'—':(r.significant?'<span class="sig sig-yes">Yes</span>':'<span class="sig sig-no">No</span>'))+'<div class="or-plain">adjusted p '+q(r.q_value)+'</div></td>'
        +'<td>'+_escH(_ruleAction(r.consequent))+'</td></tr>'; }).join('')+'</tbody></table></div>';
}
document.addEventListener('DOMContentLoaded',()=>{
  const note=document.getElementById('arm-note'); const sec=note&&note.closest('.dash-section'); if(!sec||document.getElementById('owner-rules')) return;
  const c=document.createElement('div'); c.className='chart-card owner-rules'; c.id='owner-rules';
  c.innerHTML='<div class="cc-title">Rules worth acting on</div>'
    +'<div class="cc-sub">The patterns that hold for the most students beyond chance, that are statistically significant and at least a weak link (lift 1.2 or more), with what each means for the café.</div>'
    +'<div class="card-explain" data-title="How these rules were chosen">Each rule reads "students who gave the first answer also tend to give the second". Students is how many gave both answers. Strength comes from lift: how many times more likely the second answer is among these students than among all students (1.2 to 1.3 weak, 1.3 to 1.6 moderate, above 1.6 strong). Significant means the link passed a chi-square test after adjusting for testing many rules at once. The rules are ranked by how many students follow them beyond what chance alone would give, so a rule that is strong but rare ranks below one that affects many students.</div>'
    +'<div class="or-body"><p class="or-empty">Loading…</p></div>';
  const title=sec.querySelector('.section-title'); title.insertAdjacentElement('afterend',c);
});
// ── Pricing tab: Annual / Quarterly / Monthly ──
const _PS={period:'annual', idx:0};
const _rs=v=>'₹'+Math.round(v).toLocaleString();
function _psSetup(){
  const pane=document.getElementById('wi-tab-pricing'); if(!pane||pane.querySelector('.ps-controls')) return;
  const box=document.createElement('div'); box.className='ps-controls';
  box.innerHTML='<div class="ps-seg" role="group" aria-label="Period shown in the pricing charts">'
    +'<button type="button" class="ps-seg-btn" data-p="annual">Annual</button><button type="button" class="ps-seg-btn" data-p="quarter">Quarterly</button><button type="button" class="ps-seg-btn" data-p="month">Monthly</button></div>'
    +'<div class="ps-chips" id="ps-chips"></div>';
  const note=document.createElement('p'); note.className='wi-note ps-best'; note.id='wi-price-best';
  const grid=pane.querySelector('.chart-grid-2'); pane.insertBefore(note,grid); pane.insertBefore(box,grid);
  box.querySelectorAll('.ps-seg-btn').forEach(b=>b.addEventListener('click',()=>{ _PS.period=b.dataset.p; _PS.idx=_PS.period==='month'?11:0; _psRender(); }));
}
function _psPeriodText(){
  if(_PS.period==='quarter') return {t:'in Q'+(_PS.idx+1)+' (months '+(_PS.idx*3+1)+'–'+(_PS.idx*3+3)+')', s:'Q'+(_PS.idx+1)};
  if(_PS.period==='month') return {t:'in month '+(_PS.idx+1), s:'Month '+(_PS.idx+1)};
  return {t:'over the year (12 months)', s:'Year'};
}
function _psRender(){
  const ps=window._wiPS; if(!ps||!ps.length) return;
  const P=_PS.period, k=_PS.idx;
  const rev=ps.map(p=>P==='annual'?p.annual_rev:(P==='quarter'?p.q_rev[k]:p.m_rev[k]));
  const prof=ps.map(p=>P==='annual'?p.annual_profit:(P==='quarter'?p.q_profit[k]:p.m_profit[k]));
  const labels=ps.map((p,i)=>p.label+(i===3?' (today)':''));
  const bR=rev.indexOf(Math.max(...rev)), bP=prof.indexOf(Math.max(...prof));
  const pt=_psPeriodText();
  // controls
  document.querySelectorAll('.ps-seg-btn').forEach(b=>{ const on=b.dataset.p===P; b.classList.toggle('on',on); b.setAttribute('aria-pressed',on?'true':'false'); });
  const chips=document.getElementById('ps-chips');
  if(chips){
    const n=P==='quarter'?4:(P==='month'?12:0);
    chips.innerHTML=n?'<span class="drill-label">'+(P==='quarter'?'Quarter':'Month')+':</span>'+Array.from({length:n},(_,i)=>'<button type="button" class="drill-chip'+(i===k?' on':'')+'" aria-pressed="'+(i===k)+'" data-i="'+i+'">'+(P==='quarter'?'Q'+(i+1):'M'+(i+1))+'</button>').join(''):'';
    chips.querySelectorAll('.drill-chip').forEach(c=>c.addEventListener('click',()=>{ _PS.idx=+c.dataset.i; _psRender(); }));
  }
  // charts (updated in place so the bars move smoothly between periods)
  const draw=(id,vals,best,isProfit)=>{
    const cols=vals.map((v,i)=>isProfit?(v>=0?(i===best?'#147a55':'#2f8f6acc'):'#7f9a24cc'):(i===best?'#147a55':'#2b6ea3cc'));
    const brd=vals.map((v,i)=>i===best?_ink():'rgba(0,0,0,0)');
    const lbl=(isProfit?'Profit ':'Revenue ')+pt.t;
    const ch=Chart.getChart(id);
    if(ch&&ch.$ps){ const ds=ch.data.datasets[0]; ds.data=vals; ds.backgroundColor=cols.map(c=>_mapColor(c,_theme()==='dark')); ds.borderColor=brd; ds.label=lbl; ch.data.labels=labels; ch.update(); }
    else { const c=mkChart(id,{type:'bar',data:{labels,datasets:[{label:lbl,data:vals,backgroundColor:cols,borderColor:brd,borderWidth:2,borderRadius:5}]},
        options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>(isProfit?'Profit ':'Revenue ')+_psPeriodText().t+': '+_rs(c.parsed.y)}}},
          scales:{x:{ticks:{color:'#55635D'},grid:{color:'rgba(23,35,31,.06)'}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:'rgba(23,35,31,.06)'}}}}}); if(c) c.$ps=true; }
    const sub=document.getElementById(id).closest('.chart-card').querySelector('.cc-sub');
    if(sub) sub.textContent=(isProfit?'Profit after all costs ':'Revenue ')+pt.t+'. Highlighted: the best price for this period ('+ps[best].label+').'+(isProfit?' Red bars are a loss.':'');
  };
  draw('wi-ch-price-sweep-rev',rev,bR,false);
  draw('wi-ch-price-sweep-profit',prof,bP,true);
  // the recommendation is always based on the whole year
  const note=document.getElementById('wi-price-best'), B=window._wiPSBest||{};
  if(note){
    const a=ps.find(p=>p.label===B.annual), f=ps.find(p=>p.label===B.full);
    let t='Best price for the year: '+B.annual+(a?', with '+_rs(a.annual_profit)+' profit over 12 months.':'.');
    if(f&&a&&B.full!==B.annual){
      const fm=a.full_from_month;
      t+=' Once the café is full'+(fm?' (from month '+fm+')':'')+', '+B.full+' earns more each month ('+_rs(f.m_profit[11])+' against '+_rs(a.m_profit[11])+' in month 12), because extra visitors drawn by a lower price can no longer get a seat. Consider moving to '+B.full+' once the café fills up.';
    }
    if(P!=='annual'&&ps[bP].label!==B.annual) t+=' (In the period shown, '+ps[bP].label+' earns the most; the recommendation above is for the whole year.)';
    note.textContent=t;
  }
}
let _armHmAll=false;
function _armHmButton(total){
  const card=document.getElementById('arm-heatmap-view'); if(!card) return;
  let b=card.querySelector('.hm-expand');
  if(total<=25){ if(b) b.remove(); return; }
  if(!b){ b=document.createElement('button'); b.type='button'; b.className='tool-btn hm-expand';
    b.addEventListener('click',()=>{ _armHmAll=!_armHmAll; buildArmChart('heatmap'); _syncSoon(); });
    _cardTools(card).appendChild(b); }
  b.textContent=_armHmAll?'Show top 25 only':'Expand to all '+total+' rules';
  b.setAttribute('aria-pressed',_armHmAll?'true':'false');
  _syncToolPad(card);
}
const gridCol='rgba(23,35,31,.05)';

function buildARM(RULES){
  const rules=RULES||[];
  const _m=window._RULES_META||{}; const _n=document.getElementById('arm-note');
  if(_n){
    let t=rules.length+' rules, each at least '+Math.round(((_m.min_lift||1.1)-1)*100)+'% more likely than chance (lift of '+(_m.min_lift||1.1).toFixed(2)+' or more).';
    if(_m.excluded_low_lift) t+=' Left out: '+_m.excluded_low_lift+' pairings at or near chance level';
    if(_m.excluded_structural) t+=(_m.excluded_low_lift?' and ':' Left out: ')+_m.excluded_structural+' that the cleaning rules create by themselves (time slot with late-night preference, spend with budget)';
    if(_m.excluded_low_lift||_m.excluded_structural) t+='.';
    _n.textContent=t;
  }
  const _vb=document.getElementById('arm-view-more-btn');
  if(_vb) _vb.textContent='▼ View all '+rules.length+' rules';
  armSorted=rules.slice();
  armAllSorted=rules.slice();
  renderArmTable('arm-tbody-main',armSorted.slice(0,10));
  renderArmTable('arm-tbody-all',armAllSorted);
  _armExplain();
  _ownerRules();
}

function renderArmTable(tbodyId, rules){
  const tb=document.getElementById(tbodyId);
  if(!tb) return;
  const maxLift=Math.max(...rules.map(r=>r.lift));
  tb.innerHTML=rules.map(r=>`<tr onmousemove="showTip(event,'<strong>${r.antecedent}</strong><br>⟹ ${r.consequent}<br>Support: ${(r.support*100).toFixed(1)}%<br>Confidence: ${(r.confidence*100).toFixed(1)}%<br>Lift: ${r.lift}<br>Conviction: ${r.conviction}')" onmouseleave="hideTip()">
    <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.antecedent}">${r.antecedent}</td>
    <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.consequent}">${r.consequent}</td>
    <td style="color:var(--info)">${(r.support*100).toFixed(1)}%</td>
    <td style="color:var(--accent2)">${(r.confidence*100).toFixed(1)}%</td>
    <td style="color:var(--warn)">${r.lift}</td>
    <td>${r.n_both!=null?r.n_both.toLocaleString():'—'}</td>
    <td>${_qFmt(r.q_value)}</td>
    <td>${r.significant==null?'—':(r.significant?'<span class="sig sig-yes">Yes</span>':'<span class="sig sig-no">No</span>')}</td>
    <td><div class="lift-bar" style="width:${(r.lift/maxLift*100).toFixed(0)}%"></div></td>
  </tr>`).join('');
}

function toggleViewMore(){
  armViewExpanded=!armViewExpanded;
  document.getElementById('arm-all-rules').style.display=armViewExpanded?'block':'none';
  document.getElementById('arm-view-more-btn').textContent=armViewExpanded?'▲ Show less':'▼ View all '+armAllSorted.length+' rules';
  if(armViewExpanded) renderArmTable('arm-tbody-all',armAllSorted);
}

function sortArm(field){armSorted.sort((a,b)=>typeof a[field]==='string'?a[field].localeCompare(b[field]):(field==='q_value'?a[field]-b[field]:b[field]-a[field]));renderArmTable('arm-tbody-main',armSorted.slice(0,10))}
function sortArmAll(field){armAllSorted.sort((a,b)=>typeof a[field]==='string'?a[field].localeCompare(b[field]):(field==='q_value'?a[field]-b[field]:b[field]-a[field]));renderArmTable('arm-tbody-all',armAllSorted)}

function armTab(type){
  document.querySelectorAll('#s-analytics .tab-btn').forEach((b,i)=>b.classList.toggle('active',['table','bubble','scatter','bar','heatmap'][i]===type));
  // Show container FIRST so Chart.js can measure dimensions, then build
  document.getElementById('arm-table-view').style.display=type==='table'?'block':'none';
  document.getElementById('arm-bubble-view').style.display=type==='bubble'?'block':'none';
  document.getElementById('arm-scatter-view').style.display=type==='scatter'?'block':'none';
  document.getElementById('arm-bar-view').style.display=type==='bar'?'block':'none';
  document.getElementById('arm-heatmap-view').style.display=type==='heatmap'?'block':'none';
  if(!armChartBuilt[type]){
    armChartBuilt[type]=true;
    // Small delay ensures the container is visible before Chart.js measures it
    setTimeout(()=>{buildArmChart(type);_syncSoon();}, 30);
  }
}

function buildArmChart(type){
  if(type==='bubble'){
    const topR=(window._RULES||[]).slice(0,60);
    const maxL=Math.max(...topR.map(r=>r.lift));
    mkChart('ch-arm-bubble',{type:'bubble',data:{datasets:[{label:'Rules',data:topR.map(r=>({x:r.support,y:r.confidence,r:Math.max(4,r.lift/maxL*22),antecedent:r.antecedent,consequent:r.consequent,lift:r.lift})),backgroundColor:topR.map(r=>`hsl(${(r.lift-1)*120},70%,60%)`),borderWidth:0}]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`[${c.raw.antecedent}] => [${c.raw.consequent}]
Sup:${(c.raw.x*100).toFixed(1)}% Conf:${(c.raw.y*100).toFixed(1)}% Lift:${c.raw.lift}`}}},scales:{x:{title:{display:true,text:'Support',color:'#55635D'},ticks:{color:'#55635D',callback:v=>(v*100).toFixed(0)+'%'},grid:{color:gridCol}},y:{title:{display:true,text:'Confidence',color:'#55635D'},ticks:{color:'#55635D',callback:v=>(v*100).toFixed(0)+'%'},grid:{color:gridCol}}}}});
  } else if(type==='scatter'){
    const _rScatter=window._RULES||[];
    mkChart('ch-arm-scatter',{type:'scatter',data:{datasets:[{label:'All Rules',data:_rScatter.map(r=>({x:r.confidence,y:r.lift,support:r.support,ant:r.antecedent,con:r.consequent})),backgroundColor:_rScatter.map(r=>`rgba(43,110,163,${0.3+r.lift/5})`),pointRadius:5,pointHoverRadius:8}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`[${c.raw.ant}]=>[${c.raw.con}] Lift:${c.raw.y.toFixed(3)}`}}},scales:{x:{title:{display:true,text:'Confidence',color:'#55635D'},ticks:{color:'#55635D',callback:v=>(v*100).toFixed(0)+'%'},grid:{color:gridCol}},y:{title:{display:true,text:'Lift',color:'#55635D'},ticks:{color:'#55635D'},grid:{color:gridCol}}}}});
  } else if(type==='bar'){
    const top20=(window._RULES||[]).slice(0,20);
    mkChart('ch-arm-bar',{type:'bar',data:{labels:top20.map(r=>r.antecedent+'  →  '+r.consequent),datasets:[{label:'Confidence',data:top20.map(r=>+(r.confidence*100).toFixed(1)),backgroundColor:top20.map(r=>`hsl(${r.lift*60},70%,55%)`),borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`Conf: ${c.parsed.x.toFixed(1)}%`}}},scales:{x:{ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gridCol}},y:{ticks:{color:'#55635D',font:{size:9}},grid:{color:gridCol}}}}});
  } else if(type==='heatmap'){
    // Build mini heatmap from unique antecedents × consequents (top 12×12)
    const _all=(window._RULES||[]).slice().sort((a,b)=>b.lift-a.lift);
    const top=_armHmAll?_all:_all.slice(0,25);
    const ants=[...new Set(top.map(r=>r.antecedent))];
    const cons=[...new Set(top.map(r=>r.consequent))];
    const _hs=document.getElementById('arm-heatmap-sub');
    if(_hs) _hs.textContent=(top.length<_all.length?'The '+top.length+' strongest rules (highest lift) of '+_all.length:'All '+top.length+' rules')
      +': one square per rule, '+ants.length+' "if" answers down the side and '+cons.length+' "then" answers along the bottom. Hover a square for details.';
    const hmap=[];
    ants.forEach((a,ai)=>cons.forEach((c,ci)=>{
      const match=top.find(r=>r.antecedent===a&&r.consequent===c);
      hmap.push({x:ci,y:ai,v:match?match.support:0,ant:a,con:c});
    }));
    const el2=document.getElementById('ch-arm-heatmap');
    if(!el2) return;
    el2.dataset.h=Math.max(380, ants.length*30+150);        // room for every "if" row
    { const _old=Chart.getChart(el2); if(_old) _old.destroy(); }
    if(el2.parentElement){ el2.parentElement.style.transition=_RM()?'':'height .35s ease'; el2.parentElement.style.height=el2.dataset.h+'px'; }
    _fitCanvas(el2,null);
    _armHmButton(_all.length);
    _armSetCuts(hmap.map(h=>h.v));
    _hmLegend('ch-arm-heatmap',_armLegendHtml());
    new Chart(el2,{type:'scatter',plugins:window.ChartDataLabels?[ChartDataLabels]:[],data:{datasets:[{data:hmap,
      pointStyle:'rect',
      pointRadius:ctx=>{const c=ctx.chart;return Math.min(c.width/(cons.length*2.2),c.height/(ants.length*2.2))},
      backgroundColor:ctx=>_armColor(ctx.raw.v),
      borderWidth:0
    }]},options:{
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw.v>0?`If ${ctx.raw.ant} → ${ctx.raw.con}: support ${(ctx.raw.v*100).toFixed(1)}%`:`No rule: ${ctx.raw.ant} → ${ctx.raw.con}`}},
        datalabels:{display:ctx=>ctx.dataset.data[ctx.dataIndex].v>0,formatter:v=>(v.v*100).toFixed(0)+'%',font:{size:10,weight:'600'},
          color:ctx=>{const c=_armColor(ctx.dataset.data[ctx.dataIndex].v);return (c===_ARM_STEPS[2]||c===_ARM_STEPS[3])?'#FFFFFF':'#17231F';}}},
      scales:{
        x:{type:'linear',min:-0.5,max:cons.length-0.5,afterBuildTicks:a=>{a.ticks=cons.map((_,i)=>({value:i}));},ticks:{stepSize:1,autoSkip:false,callback:v=>Number.isInteger(v)&&cons[v]?_hmColLabel(cons[v]):'',color:'#55635D',maxRotation:0,font:{size:11}},grid:{display:false}},
        y:{type:'linear',min:-0.5,max:ants.length-0.5,reverse:true,afterBuildTicks:a=>{a.ticks=ants.map((_,i)=>({value:i}));},ticks:{stepSize:1,autoSkip:false,callback:v=>Number.isInteger(v)&&ants[v]?_hmRowLabel(ants[v],40):'',color:'#55635D'},grid:{display:false}}
      }
    }});
  }
}
// ══ Predictive Analytics ══
let _predictiveLoaded=false;

async function openPredictive(){
  goTo('s-predictive');
  if(_predictiveLoaded) return;
  try{
    const res=await fetch(BASE+'/predictive_data');
    const d=await res.json();
    if(d.error){alert('Predictive error: '+d.error);goTo('s-nav');return;}
    window._PREDDATA=d;
    renderPredictive(d);
    _predictiveLoaded=true;
  }catch(e){alert('Failed to load predictive data: '+e.message);goTo('s-nav');}
}

function insightBox(elId, title, text, accuracy, baseline){
  const el=document.getElementById(elId);
  if(!el)return;
  const acc=accuracy!==undefined?`<span class="pill pill-blue" style="margin-left:8px">Model accuracy: ${(accuracy*100).toFixed(1)}%${(baseline!==undefined&&baseline!==null)?' · naive baseline '+(baseline*100).toFixed(1)+'%':''}</span>`:'';
  el.innerHTML=`<h4>${title}${acc}</h4><p>${text}</p>`;
}

function _renderPredictiveCore(d){
  // v41: show the section BEFORE drawing, so every chart can measure its real size
  document.getElementById('pred-loading').style.display='none';
  document.getElementById('pred-body').style.display='block';
  const gC='rgba(23,35,31,.05)';
  const PAL=['#2b6ea3','#d2692c','#2f8f6a','#cfae2a','#a8508e','#6a86d8','#c4534a','#7f9a24'];
  const TIME_ORDER=['Morning','Afternoon','Evening','Late Night'];

  // ── A. PEAK HOURS ──────────────────────────────────────────────────────────
  const ph=d.peak_hours;
  // time slots come from the data (and drop out when a lever leaves them out)
  const timeLabels=TIME_ORDER.filter(t=>t in (ph.distribution||{})).concat(Object.keys(ph.distribution||{}).filter(t=>!TIME_ORDER.includes(t)));
  const TIME_COL=timeLabels.map(t=>_entCol('time',t));
  const timePcts=timeLabels.map(t=>ph.distribution[t]||0);
  mkChart('ch-time-dist',{type:'doughnut',data:{labels:timeLabels,datasets:[{data:timePcts,backgroundColor:TIME_COL,borderWidth:2,borderColor:'#FFFFFF',hoverOffset:8}]},
    options:{plugins:{legend:{labels:{color:'#55635D',padding:14}},tooltip:{callbacks:{label:c=>`${c.label}: ${c.parsed.toFixed(1)}% of customers`}}}}});

  mkChart('ch-slot-rev',{type:'bar',data:{labels:timeLabels,
    datasets:[{label:'Revenue (₹)',data:timeLabels.map(t=>ph.revenue_per_100[t]||0),backgroundColor:TIME_COL,borderRadius:6,borderSkipped:false}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()} per 100 customers`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  mkChart('ch-spend-by-slot',{type:'bar',data:{labels:timeLabels,
    datasets:[{label:'Avg Spend (₹)',data:timeLabels.map(t=>ph.spend_by_slot[t]||0),backgroundColor:'#2b6ea3cc',borderRadius:5}]},
    options:{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.x}`}}},
      scales:{x:{ticks:{color:'#55635D',callback:v=>'₹'+v},grid:{color:gC}},y:{ticks:{color:'#55635D'},grid:{color:gC}}}}});
  // Update subtitle dynamically with the real top-spend slot
  const _spendSlots=ph.spend_by_slot;
  const _topSpendSlot=Object.keys(_spendSlots).reduce((a,b)=>(_spendSlots[a]||0)>=(_spendSlots[b]||0)?a:b,'');
  const _spendSubEl=document.getElementById('spend-slot-sub');
  if(_spendSubEl && _topSpendSlot) _spendSubEl.textContent=`${_topSpendSlot} visitors tend to spend the most (₹${_spendSlots[_topSpendSlot]}). Use this for time-based pricing/offers.`;

  const pkFi=Object.entries(ph.feature_importances).sort((a,b)=>b[1]-a[1]);
  mkChart('ch-peak-fi',{type:'bar',data:{labels:pkFi.map(f=>f[0]),
    datasets:[{label:'Importance',data:pkFi.map(f=>+(f[1]*100).toFixed(1)),backgroundColor:'#2b6ea3',borderRadius:5}]},
    options:{indexAxis:'y',plugins:{legend:{display:false}},
      scales:{x:{ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gC}},y:{ticks:{color:'#55635D',font:{size:9}},grid:{color:gC}}}}});

  insightBox('peak-insight','Peak Hours Insight',ph.insight,ph.rf_accuracy,ph.baseline);

  // ── B. FOOTFALL ────────────────────────────────────────────────────────────
  const ft=d.footfall;
  document.getElementById('foot-kpis').innerHTML=[
    {val:ft.avg_visits_per_week,lab:'Avg Visits/Week',col:'var(--accent)'},
    {val:ft.daily_proj_1000.toLocaleString(),lab:'Projected Daily (1k catchment)',col:'var(--accent2)'},
    {val:(ft.daily_proj_1000*7).toLocaleString(),lab:'Projected Weekly',col:'var(--accent3)'},
    {val:(ft.daily_proj_1000*30).toLocaleString(),lab:'Projected Monthly',col:'var(--warn)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  mkChart('ch-slot-foot',{type:'bar',data:{labels:timeLabels,
    datasets:[{label:'Visitors',data:timeLabels.map(t=>ft.slot_footfall[t]||0),backgroundColor:TIME_COL,borderRadius:6}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`~${c.parsed.y} visitors`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D'},grid:{color:gC}}}}});

  const wkDays=Object.keys(ft.weekly_pattern),wkVals=Object.values(ft.weekly_pattern);
  mkChart('ch-weekly-pattern',{type:'line',data:{labels:wkDays,datasets:[{label:'Visits',data:wkVals,
    borderColor:'#2b6ea3',backgroundColor:'rgba(43,110,163,.15)',pointBackgroundColor:'#2b6ea3',
    pointRadius:6,borderWidth:2.5,fill:true,tension:.4}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`~${c.parsed.y} visits`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D'},grid:{color:gC}}}}});

  const distLabels=Object.keys(ft.foot_by_dist),distVals=Object.values(ft.foot_by_dist);
  mkChart('ch-foot-dist',{type:'bar',data:{labels:distLabels,
    datasets:[{label:'Avg Visits/Week',data:distVals,backgroundColor:PAL.slice(0,distLabels.length).map(c=>c+'cc'),borderRadius:5}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y} visits/wk avg`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D'},grid:{color:gC}}}}});

  insightBox('foot-insight','Footfall Insight',ft.insight);

  // ── C. REVENUE ─────────────────────────────────────────────────────────────
  const rv=d.revenue;
  document.getElementById('rev-kpis').innerHTML=[
    {val:'₹'+rv.avg_spend,lab:'Avg Spend/Visit',col:'var(--accent)'},
    {val:'₹'+rv.daily_revenue.toLocaleString(),lab:'Daily Revenue (1k)',col:'var(--accent2)'},
    {val:'₹'+rv.weekly_revenue.toLocaleString(),lab:'Weekly Revenue',col:'var(--accent3)'},
    {val:'₹'+rv.monthly_revenue.toLocaleString(),lab:'Monthly Revenue',col:'var(--warn)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  const spLabels=_levOrder('spend',Object.keys(rv.distribution)),spVals=spLabels.map(l=>rv.distribution[l]);
  mkChart('ch-spend-dist',{type:'pie',data:{labels:spLabels.map(l=>'₹'+l),datasets:[{data:spVals,backgroundColor:spLabels.map(l=>_entCol('spend',l)),borderWidth:2,borderColor:'#FFFFFF',hoverOffset:8}]},
    options:{plugins:{legend:{labels:{color:'#55635D'}},tooltip:{callbacks:{label:c=>`${c.label}: ${c.parsed.toFixed(1)}%`}}}}});

  mkChart('ch-rev-slot',{type:'bar',data:{labels:timeLabels,
    datasets:[{label:'Revenue (₹)',data:timeLabels.map(t=>rv.rev_by_slot[t]||0),backgroundColor:TIME_COL,borderRadius:6}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  const months=['M1','M2','M3','M4','M5','M6','M7','M8','M9','M10','M11','M12'];
  mkChart('ch-rev-projection',{type:'line',data:{labels:months,datasets:[{label:'Monthly Revenue (₹)',
    data:rv.monthly_projection,borderColor:'#2f8f6a',backgroundColor:'rgba(47,143,106,.1)',
    pointBackgroundColor:'#2f8f6a',pointRadius:5,borderWidth:2.5,fill:true,tension:.4}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${Math.round(c.parsed.y).toLocaleString()}`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  const univLabels=Object.keys(rv.spend_by_univ),univSpend=univLabels.map(u=>rv.spend_by_univ[u].mean);
  mkChart('ch-spend-univ',{type:'bar',data:{labels:univLabels,
    datasets:[{label:'Avg Spend (₹)',data:univSpend,backgroundColor:PAL.slice(0,univLabels.length).map(c=>c+'cc'),borderRadius:5}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y} avg spend`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v},grid:{color:gC}}}}});

  insightBox('rev-insight','Revenue Insight',rv.insight);

  // ── D. SEATING ─────────────────────────────────────────────────────────────
  const se=d.seating;
  const seatLabels=Object.keys(se.overall_pct),seatVals=Object.values(se.overall_pct);
  mkChart('ch-seat-overall',{type:'doughnut',data:{labels:seatLabels,datasets:[{data:seatVals,backgroundColor:seatLabels.map(l=>_entCol('seat',l)),borderWidth:2,borderColor:'#FFFFFF',hoverOffset:8}]},
    options:{plugins:{legend:{labels:{color:'#55635D',padding:12}},tooltip:{callbacks:{label:c=>`${c.label}: ${c.parsed.toFixed(1)}%`}}}}});

  // Seating stacked bar by time slot
  const seatTypes=seatLabels;
  const seatByTimeDS=seatTypes.map((st,si)=>({
    label:st,
    data:timeLabels.map(t=>se.by_time_slot[t]?+(se.by_time_slot[t][st]||0).toFixed(1):0),
    backgroundColor:_entCol('seat',st)+'cc',borderRadius:3,
  }));
  mkChart('ch-seat-by-time',{type:'bar',data:{labels:timeLabels,datasets:seatByTimeDS},
    options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}}},
      scales:{x:{stacked:true,ticks:{color:'#55635D'},grid:{color:gC}},
              y:{stacked:true,ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gC}}}}});

  insightBox('seat-insight','Seating Demand Insight',se.insight,se.rf_accuracy,se.baseline);

  // ── E. MEMBERSHIP ──────────────────────────────────────────────────────────
  const mem=d.membership;
  document.getElementById('mem-kpis').innerHTML=[
    {val:mem.overall_rate+'%',lab:'Overall Membership Rate',col:'var(--accent)'},
    {val:(mem.accuracy*100).toFixed(1)+'%',lab:'Test Accuracy — '+(mem.best_model||'best model'),col:'var(--accent2)'},
    {val:(mem.f1_1*100).toFixed(1)+'%',lab:'F1 Score (Members)',col:'var(--accent3)'},
    {val:(mem.precision_1*100).toFixed(1)+'%',lab:'Precision (Members)',col:'var(--warn)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  const mNames=Object.keys(mem.model_scores),mAccs=Object.values(mem.model_scores).map(v=>+(v.mean*100).toFixed(1));
  mkChart('ch-mem-models',{type:'bar',data:{labels:mNames,datasets:[{label:'CV Accuracy',data:mAccs,
    backgroundColor:['#2b6ea3cc','#2f8f6acc','#d2692ccc'],borderRadius:6}]},
    options:{plugins:{legend:{display:false}},
      scales:{x:{ticks:{color:'#55635D',font:{size:9},maxRotation:15},grid:{color:gC}},
              y:{min:0,max:100,ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gC}}}}});

  const cmData=[
    {v:mem.confusion_matrix[0][0],l:'True No',bg:'rgba(47,143,106,.25)',vc:'var(--accent2)'},
    {v:mem.confusion_matrix[0][1],l:'False Yes',bg:'rgba(127,154,36,.15)',vc:'var(--danger)'},
    {v:mem.confusion_matrix[1][0],l:'False No',bg:'rgba(127,154,36,.15)',vc:'var(--danger)'},
    {v:mem.confusion_matrix[1][1],l:'True Yes',bg:'rgba(47,143,106,.25)',vc:'var(--accent2)'},
  ];
  try{ _addChartGuide('mem-conf-matrix'); }catch(e){}
  { const ct=document.getElementById('mem-conf-title'); if(ct) ct.textContent='Confusion matrix — '+(mem.best_model||'best model'); }
  document.getElementById('mem-conf-matrix').innerHTML=cmData.map(c=>
    `<div class="conf-cell" style="background:${c.bg}"><div class="cv" style="color:${c.vc}">${c.v}</div><div class="cl">${c.l}</div></div>`).join('');

  const msLabels=Object.keys(mem.by_spend_tier),msVals=Object.values(mem.by_spend_tier);
  mkChart('ch-mem-spend',{type:'bar',data:{labels:msLabels.map(l=>'₹'+l),datasets:[{label:'Membership Rate %',data:msVals,
    backgroundColor:PAL.slice(0,msLabels.length).map(c=>c+'cc'),borderRadius:6}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y}% willing`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{min:0,max:100,ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gC}}}}});

  const mdLabels=Object.keys(mem.by_distance),mdVals=Object.values(mem.by_distance);
  mkChart('ch-mem-dist',{type:'bar',data:{labels:mdLabels,datasets:[{label:'Membership Rate %',data:mdVals,
    backgroundColor:PAL.slice(0,mdLabels.length).map(c=>c+'cc'),borderRadius:6}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y}% willing`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{min:0,max:100,ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gC}}}}});

  const mFi=Object.entries(mem.feature_importances).sort((a,b)=>b[1]-a[1]).slice(0,10);
  mkChart('ch-mem-fi',{type:'bar',data:{labels:mFi.map(f=>f[0]),
    datasets:[{label:'Importance',data:mFi.map(f=>+(f[1]*100).toFixed(1)),backgroundColor:'#2b6ea3',borderRadius:5}]},
    options:{indexAxis:'y',plugins:{legend:{display:false}},
      scales:{x:{ticks:{color:'#55635D',callback:v=>v+'%'},grid:{color:gC}},
              y:{ticks:{color:'#55635D',font:{size:9}},grid:{color:gC}}}}});

  insightBox('mem-insight','Membership Prediction Insight',mem.insight,mem.best_cv_accuracy||mem.accuracy,mem.baseline);

  // ── F. Model Performance Summary ─────────────────────────────────────────
  renderModelSummary(d.model_summary);

  document.getElementById('pred-loading').style.display='none';
  document.getElementById('pred-body').style.display='block';
}

function renderModelSummary(ms){
  if(!ms) return;

  // ── Full scorecard table ─────────────────────────────────────────────
  const allModels = [...(ms.classification||[]), ...(ms.regression||[])];
  // Sort by score desc within each task group
  const taskOrder = ['Membership Prediction','Seating Type Prediction','Visit Frequency','Avg Spend per Visit'];
  const grouped = taskOrder.map(task=>({
    task, rows: allModels.filter(m=>m.task===task).sort((a,b)=>b.score-a.score)
  })).filter(g=>g.rows.length>0);

  const rankBadge = (i)=>{
    const cls = i===0?'gold':i===1?'silver':i===2?'bronze':'';
    return `<span class="rank-badge ${cls}">${i+1}</span>`;
  };

  let tableHtml = `<table class="scorecard-table">
    <thead><tr>
      <th>#</th><th>Model</th><th>Task</th><th>Metric</th>
      <th>Score</th><th style="min-width:140px">Performance Bar</th>
    </tr></thead><tbody>`;

  grouped.forEach(({task, rows})=>{
    const isCls = rows[0].metric==='Accuracy';
    const maxAbs = isCls ? 100 : Math.max(...rows.map(r=>Math.abs(r.score)), 0.01);
    rows.forEach((m,i)=>{
      const scoreDisp = isCls ? m.score.toFixed(1)+'%' : m.score.toFixed(3);
      const isGood = isCls ? m.score>=60 : m.score>=0.3;
      const scoreCol = isGood ? 'var(--accent2)' : m.score<(isCls?40:0) ? 'var(--danger)' : 'var(--warn)';
      const barW = Math.max(0, (isCls ? m.score : Math.max(0,m.score)/maxAbs*100)).toFixed(0);
      const barCol = isGood ? '#2f8f6a' : m.score<(isCls?40:0) ? '#7f9a24' : '#d2692c';
      tableHtml += `<tr>
        <td>${rankBadge(i)}</td>
        <td style="color:var(--text);font-weight:600">${m.model}</td>
        <td style="color:var(--muted2);font-size:.78rem">${task}</td>
        <td style="color:var(--muted2)">${m.metric}</td>
        <td class="score-cell" style="color:${scoreCol}">${scoreDisp}<br><span style="color:var(--muted2);font-size:.72rem;font-weight:400">± ${isCls?m.std.toFixed(1)+'%':m.std.toFixed(3)}</span></td>
        <td><div class="score-bar-wrap"><div class="score-bar-bg"><div class="score-bar-fill" style="width:${barW}%;background:${barCol}"></div></div></div></td>
      </tr>`;
    });
    // Spacer row between task groups
    tableHtml += `<tr><td colspan="6" style="height:6px;background:rgba(43,110,163,.05)"></td></tr>`;
  });

  tableHtml += '</tbody></table>';
  document.getElementById('model-scorecard-wrap').innerHTML = tableHtml;
}

// Free the port when the tab is closed
// ══ What-If Scenario Modelling ══
function openWhatIf(){
  goTo('s-whatif');
}

// ── Scenario storage for Compare tab ───────────────────────────
let _wiScenarios = [];

function wiTab(name){
  document.querySelectorAll('.wi-tab-pane').forEach(p=>p.style.display='none');
  document.getElementById('wi-tab-'+name).style.display='block';
  document.querySelectorAll('#s-whatif .tab-btn').forEach((b,i)=>{
    b.classList.toggle('active', b.getAttribute('onclick').includes("'"+name+"'"));
  });
}

function saveScenario(){
  if(!window._lastWiData) return;
  const s = window._lastWiData.scenario_summary;
  if(_wiScenarios.length >= 5){
    alert('Maximum 5 scenarios saved. Remove one first.');
    return;
  }
  _wiScenarios.push({n: _wiScenarios.length+1, ...s});
  _renderCompareTab();
  wiTab('compare');
}

function removeScenario(n){
  _wiScenarios = _wiScenarios.filter(s=>s.n!==n);
  _wiScenarios.forEach((s,i)=>s.n=i+1);
  _renderCompareTab();
}

function _renderCompareTab(){
  const tbody = document.getElementById('wi-compare-tbody');
  const gC='rgba(23,35,31,.05)';
  if(_wiScenarios.length===0){
    tbody.innerHTML='<tr><td colspan="13" style="text-align:center;color:var(--muted2);padding:20px">No scenarios saved yet.</td></tr>';
    return;
  }
  tbody.innerHTML = _wiScenarios.map(s=>`
    <tr>
      <td style="color:var(--accent);font-weight:700">${s.n}</td>
      <td>${s.catchment.toLocaleString()}</td>
      <td>${s.growth_pct}%</td>
      <td>${s.price_adj_pct >= 0?'+':''}${s.price_adj_pct}%</td>
      <td>${s.comp_impact_pct}%</td>
      <td>₹${s.fixed_costs.toLocaleString()}</td>
      <td style="font-weight:700;color:var(--accent2)">₹${s.monthly_revenue.toLocaleString()}</td>
      <td style="font-weight:700;color:var(--accent2)">₹${s.annual_revenue.toLocaleString()}</td>
      <td style="color:${s.profit_monthly_m1>=0?'var(--accent2)':'var(--danger)'}">₹${s.profit_monthly_m1.toLocaleString()}</td>
      <td>${s.breakeven_month ? 'Month '+s.breakeven_month : '> 12 mo'}</td>
      <td>${s.members.toLocaleString()}</td>
      <td>₹${s.tier_combined_monthly.toLocaleString()}</td>
      <td><button onclick="removeScenario(${s.n})"
          style="background:none;border:1px solid var(--border2);border-radius:4px;color:var(--danger);cursor:pointer;padding:2px 8px;font-size:.75rem">✕</button></td>
    </tr>`).join('');

  // Comparison chart
  mkChart('wi-ch-compare',{type:'bar',
    data:{
      labels: _wiScenarios.map(s=>'Scenario '+s.n),
      datasets:[
        {label:'Annual Revenue (₹)', data:_wiScenarios.map(s=>s.annual_revenue),
         backgroundColor:'#2f8f6acc', borderRadius:6},
        {label:'Annual Tier Rev (₹)', data:_wiScenarios.map(s=>s.tier_combined_monthly*12),
         backgroundColor:'#2b6ea3cc', borderRadius:6},
      ]
    },
    options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}},
      tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
      scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},
              y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}
  });
}

async function runWhatIf(){
  const btn=document.getElementById('wi-run-btn');
  btn.disabled=true; btn.textContent='Computing…';
  const body={
    catchment:        +document.getElementById('wi-catchment').value,
    monthly_growth:   +document.getElementById('wi-growth').value,
    price_adjustment: +document.getElementById('wi-price').value,
    operating_hours:  +document.getElementById('wi-hours').value,
    membership_target:+document.getElementById('wi-memtarget').value,
    occupancy_rate:   +document.getElementById('wi-occupancy').value,
    total_seats:      +document.getElementById('wi-seats').value,
    fixed_costs:      +document.getElementById('wi-fixedcosts').value,
    competitor_impact:+document.getElementById('wi-competitor').value,
    staff_cost_per_shift:+document.getElementById('wi-staffcost').value,
    tier1_price:      +document.getElementById('wi-tier1price').value,
    tier1_uptake:     +document.getElementById('wi-tier1uptake').value,
    tier2_price:      +document.getElementById('wi-tier2price').value,
  };
  try{
    const res=await fetch(BASE+'/whatif',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    const d=await res.json();
    if(d.error){alert('Scenario error: '+d.error);btn.disabled=false;btn.textContent='Run scenario';return;}
    window._lastWiData = d;
    renderWhatIf(d);
    btn.disabled=false; btn.textContent='Run scenario';
    document.getElementById('wi-save-btn').style.display='inline-flex';
  }catch(e){
    alert('Failed: '+e.message);
    btn.disabled=false; btn.textContent='Run scenario';
  }
}

function renderWhatIf(d){
  const gC='rgba(23,35,31,.05)';
  const PAL=['#2b6ea3','#d2692c','#2f8f6a','#cfae2a','#a8508e','#6a86d8','#c4534a','#7f9a24'];
  const res=d.revenue; const ft=d.footfall; const mem=d.membership; const seat=d.seating;
  const be=d.breakeven; const ps=d.price_sweep; const comp=d.competitor;
  const sf=d.staffing; const tr=d.tiers;

  // ── KPI strip ────────────────────────────────────────────────
  document.getElementById('wi-kpi-strip').innerHTML=[
    {val:ft.daily.toLocaleString(),           lab:'Daily Visitors Served',     col:'var(--accent)'},
    {val:(seat.capacity_utilisation_pct!==undefined?seat.capacity_utilisation_pct+'%':'—'), lab:(seat.capacity_limited?'Seat Capacity Used (FULL — demand '+seat.daily_demand+'/day)':'Seat Capacity Used'), col:seat.capacity_limited?'var(--danger)':'var(--accent2)'},
    {val:'₹'+res.daily.toLocaleString(),      lab:'Daily Revenue (₹)',         col:'var(--accent2)'},
    {val:'₹'+res.monthly.toLocaleString(),    lab:'Monthly Revenue (₹)',       col:'var(--accent3)'},
    {val:'₹'+res.annual.toLocaleString(),     lab:'Annual Revenue (₹)',        col:'var(--warn)'},
    {val:mem.members_estimated.toLocaleString(),lab:'Est. Members',            col:'var(--info)'},
    {val:be.breakeven_month ? 'Month '+be.breakeven_month : '>12 mo',
                                               lab:'Break-Even Month',         col:'var(--purple)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  // Show results area, activate Revenue tab
  document.getElementById('wi-results').style.display='block';
  wiTab('revenue');
  document.getElementById('wi-results').scrollIntoView({behavior:'smooth'});

  // ══ TAB 1: REVENUE ════════════════════════════════════════════
  const months=['M1','M2','M3','M4','M5','M6','M7','M8','M9','M10','M11','M12'];
  const capM=seat.capacity_reached_month, unc=res.monthly_projection_uncapped||[];
  const projSub=document.getElementById('wi-proj-sub');
  if(projSub){
    const capTxt=(seat.effective_capacity_per_day||0).toLocaleString()+' visits a day';
    projSub.textContent = capM===1
      ? 'Revenue is flat because the café is full from month 1: the seats can host '+capTxt+' and demand is already '+(seat.daily_demand||0).toLocaleString()+'. Growth adds visitors who cannot be seated, not revenue. Add seats, hours or occupancy to lift it. Dashed line: revenue if every visitor could be seated.'
      : capM
      ? 'Growth lifts revenue until month '+capM+', when the seats fill ('+capTxt+'); after that it stays flat. Dashed line: revenue if every visitor could be seated.'
      : 'Monthly revenue with the growth rate applied. Demand stays within what the seats can host ('+capTxt+') all year.';
  }
  const projSets=[
    {label:'Revenue the seats can earn (₹)',data:res.monthly_projection,borderColor:'#2f8f6a',
     backgroundColor:'rgba(47,143,106,.12)',pointBackgroundColor:'#2f8f6a',
     pointRadius:5,borderWidth:2.5,fill:true,tension:.4}];
  if(capM&&unc.length) projSets.push({label:'If every visitor could be seated (₹)',data:unc,borderColor:'#d2692c',backgroundColor:'rgba(0,0,0,0)',
     pointBackgroundColor:'#d2692c',pointRadius:3,borderWidth:2,borderDash:[6,5],fill:false,tension:.4});
  mkChart('wi-ch-projection',{type:'line',data:{labels:months,datasets:projSets},options:{plugins:{legend:{display:projSets.length>1,labels:{color:'#55635D',boxWidth:10}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ₹${Math.round(c.parsed.y).toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  const slotLabels=Object.keys(d.slot_breakdown);
  mkChart('wi-ch-slots',{type:'bar',data:{labels:slotLabels,datasets:[
    {label:'Revenue (₹)',data:slotLabels.map(t=>d.slot_breakdown[t].revenue),
     backgroundColor:PAL.slice(0,slotLabels.length).map(c=>c+'cc'),borderRadius:6}
  ]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  const sens=d.sensitivity; const sensKeys=Object.keys(sens);
  document.getElementById('wi-sens-tbody').innerHTML=sensKeys.map((k,i)=>`
    <tr class="${k==='Base'?'base-row':''}">
      <td>${k}</td><td>${sens[k].daily_footfall.toLocaleString()}</td>
      <td>₹${sens[k].daily_revenue.toLocaleString()}</td>
      <td>₹${sens[k].monthly_revenue.toLocaleString()}</td>
    </tr>`).join('');
  mkChart('wi-ch-sensitivity',{type:'bar',data:{labels:sensKeys,datasets:[
    {label:'Monthly Revenue (₹)',data:sensKeys.map(k=>sens[k].monthly_revenue),
     backgroundColor:sensKeys.map(k=>k==='Base'?'#2f8f6acc':'#2b6ea3cc'),borderRadius:5}
  ]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  mkChart('wi-ch-membership',{type:'bar',data:{
    labels:['Monthly Walk-in','Monthly Membership','Annual Walk-in','Annual Membership'],
    datasets:[{label:'Revenue (₹)',
      data:[res.monthly, mem.monthly_revenue, res.annual, mem.annual_revenue],
      backgroundColor:['#2f8f6acc','#2b6ea3cc','#2f8f6a','#2b6ea3'],borderRadius:6}]
  },options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D',font:{size:9}},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  mkChart('wi-ch-seating',{type:'bar',data:{
    labels:['Daily Demand','Daily Seat Capacity','Visitors Served'],
    datasets:[{label:'Visits per day',data:[seat.daily_demand,seat.effective_capacity_per_day,seat.visitors_served],
     backgroundColor:['#2b6ea3cc','#2f8f6acc','#d2692ccc'],borderRadius:6}]
  },options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y} visits/day`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D'},grid:{color:gC}}}}});

  document.getElementById('wi-arima-text').textContent=(d.model_notes?('How this scenario is calculated: '+d.model_notes.join(' ')+'  ·  '):'')+d.arima_note;

  // ══ TAB 2: BREAK-EVEN ════════════════════════════════════════
  document.getElementById('wi-be-kpis').innerHTML=[
    {val:'₹'+be.fixed_costs_monthly.toLocaleString(),      lab:'Fixed Costs/Month',      col:'var(--danger)'},
    {val:be.variable_cost_rate_pct+'%',                    lab:'Variable Cost Rate',      col:'var(--warn)'},
    {val:'₹'+be.total_cost_monthly.toLocaleString(),       lab:'Total Monthly Cost',      col:'var(--accent3)'},
    {val:(be.breakeven_daily_visitors===null?'n/a':be.breakeven_daily_visitors.toLocaleString()), lab:'Break-Even Daily Visitors'+(be.breakeven_within_capacity?'':' (above seat capacity)'),col:be.breakeven_within_capacity?'var(--accent)':'var(--danger)'},
    {val:be.breakeven_month ? 'Month '+be.breakeven_month : '>12 mo',
                                                            lab:'Break-Even Month',       col:'var(--purple)'},
    {val:be.is_profitable_at_base ? '✅ Yes' : '❌ No',   lab:'Profitable at Base',      col:be.is_profitable_at_base?'var(--accent2)':'var(--danger)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  mkChart('wi-ch-profit',{type:'bar',data:{labels:months,datasets:[
    {label:'Monthly Profit (₹)',data:be.profit_monthly_projection,
     backgroundColor:be.profit_monthly_projection.map(v=>v>=0?'#2f8f6acc':'#7f9a24cc'),borderRadius:5}
  ]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${Math.round(c.parsed.y).toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  mkChart('wi-ch-breakeven-bar',{type:'bar',__semantic:true,data:{
    labels:['Fixed Costs','Labour','Variable Costs','Walk-in Revenue','Membership Fees'],
    datasets:[{label:'Amount (₹)',
      data:[be.fixed_costs_monthly, be.labour_cost_monthly, be.variable_cost_monthly, res.monthly, be.membership_income_monthly],
      backgroundColor:['#c4534acc','#c4534a99','#c4534a66','#2f8f6acc','#2f8f6a99'],borderRadius:6}]
  },options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  const beSummaryEl=document.getElementById('wi-be-summary');
  beSummaryEl.innerHTML=`
    <strong style="color:var(--text)">Break-Even Analysis</strong><br>
    Fixed monthly costs: <strong style="color:var(--danger)">₹${be.fixed_costs_monthly.toLocaleString()}</strong><br>
    Labour (staff shifts): <strong style="color:var(--danger)">₹${be.labour_cost_monthly.toLocaleString()}</strong><br>
    Variable costs (${be.variable_cost_rate_pct}% of walk-in revenue): <strong style="color:var(--warn)">₹${be.variable_cost_monthly.toLocaleString()}</strong><br>
    Total monthly cost: <strong style="color:var(--accent3)">₹${be.total_cost_monthly.toLocaleString()}</strong><br>
    Walk-in revenue: <strong style="color:var(--accent2)">₹${res.monthly.toLocaleString()}</strong> + membership fees: <strong style="color:var(--accent2)">₹${be.membership_income_monthly.toLocaleString()}</strong><br>
    Month-1 profit: <strong style="color:${be.profit_monthly_projection[0]>=0?'var(--accent2)':'var(--danger)'}">₹${Math.round(be.profit_monthly_projection[0]).toLocaleString()}</strong><br>
    Break-even daily visitors needed: <strong style="color:var(--accent)">${be.breakeven_daily_visitors===null?'n/a':'≥ '+be.breakeven_daily_visitors.toLocaleString()}</strong> (seats can host ${seat.effective_capacity_per_day.toLocaleString()}/day)<br>
    ${be.breakeven_month
      ? `<strong style="color:var(--accent2)">Break-even reached at Month ${be.breakeven_month}</strong> based on projected growth.`
      : `<strong style="color:var(--danger)">Break-even not reached within 12 months</strong> at current parameters. Consider increasing catchment, price, or reducing costs.`}
  `;

  // ══ TAB 3: PRICING SENSITIVITY ═══════════════════════════════
  window._wiPS=ps; window._wiPSBest={annual:d.best_price_label, full:d.best_price_full_label};
  _psSetup(); _psRender();
  try{ _sgRender(d); }catch(e){ console.warn('sens grid',e); }
  const _bestA=d.best_price_label;
  document.getElementById('wi-price-sweep-tbody').innerHTML=ps.map((p,i)=>`
    <tr class="${p.label===_bestA?'best-row':(i===3?'base-row':'')}">
      <td>${p.label}${i===3?' (today)':''}${p.label===_bestA?' <span class="sig sig-yes">Best for the year</span>':''}</td>
      <td>₹${p.adj_spend.toLocaleString()}${p.demand_retained_pct!==undefined&&p.demand_retained_pct<100?' · '+p.demand_retained_pct+'% of visits kept':''}</td>
      <td>₹${p.annual_rev.toLocaleString()}</td>
      <td style="color:${p.annual_profit>=0?'var(--accent2)':'var(--danger)'}">₹${p.annual_profit.toLocaleString()}</td>
      <td style="color:${p.m_profit[11]>=0?'var(--accent2)':'var(--danger)'}">₹${p.m_profit[11].toLocaleString()}</td>
      <td>${p.annual_profit>=0?'<span style="color:var(--accent2)">Profitable over the year</span>':'<span style="color:var(--danger)">Loss over the year</span>'}</td>
    </tr>`).join('');

  // ══ TAB 4: COMPETITOR ════════════════════════════════════════
  const baseAnnual = comp[0].annual_rev;
  const _cLast=comp[comp.length-1];
  document.getElementById('wi-comp-kpis').innerHTML=[
    {val:'-'+d.inputs.competitor_impact_pct+'%',          lab:'Current Competitor Impact',            col:'var(--danger)'},
    {val:'₹'+_cLast.annual_rev.toLocaleString(),           lab:'Annual Revenue at '+_cLast.label,      col:'var(--warn)'},
    {val:'₹'+_cLast.revenue_loss.toLocaleString(),         lab:'Revenue Lost ('+_cLast.label+')',      col:'var(--danger)'},
    {val:seat.capacity_limited?'Yes':'No',                 lab:'Café full even after the loss?',       col:seat.capacity_limited?'var(--accent2)':'var(--muted2)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  mkChart('wi-ch-comp-rev',{type:'bar',data:{labels:comp.map(c=>c.label),datasets:[
    {label:'Annual Revenue (₹)',data:comp.map(c=>c.annual_rev),
     backgroundColor:comp.map((c,i)=>c.is_current?'#d2692ccc':(i===0?'#2f8f6acc':'#2b6ea3cc')),borderRadius:6}
  ]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D',font:{size:9}},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  mkChart('wi-ch-comp-loss',{type:'bar',data:{labels:comp.map(c=>c.label),datasets:[
    {label:'Revenue Loss (₹)',data:comp.map(c=>c.revenue_loss),
     backgroundColor:comp.map((_,i)=>i===0?'rgba(47,143,106,.3)':'#7f9a24cc'),borderRadius:6}
  ]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`-₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D',font:{size:9}},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  document.getElementById('wi-comp-tbody').innerHTML=comp.map((c,i)=>`
    <tr class="${i===0?'base-row':''}">
      <td>${c.label}${i===0?' (Base)':''}${c.is_current&&i!==0?' (Current)':''}</td>
      <td>${c.catchment_eff.toLocaleString()}</td>
      <td style="color:var(--accent2)">₹${c.annual_rev.toLocaleString()}</td>
      <td style="color:${i===0?'var(--accent2)':'var(--danger)'}">
        ${c.revenue_loss===0?'—':'-₹'+c.revenue_loss.toLocaleString()}</td>
    </tr>`).join('');

  // ══ TAB 5: STAFFING ══════════════════════════════════════════
  document.getElementById('wi-staff-kpis').innerHTML=[
    {val:sf.slots.reduce((a,s)=>a+s.staff_needed,0),   lab:'Total Staff Needed (All Slots)', col:'var(--accent)'},
    {val:'₹'+sf.total_daily_labour.toLocaleString(),   lab:'Total Daily Labour Cost',        col:'var(--danger)'},
    {val:'₹'+sf.net_daily_revenue.toLocaleString(),    lab:'Net Daily Revenue (After Labour)',col:sf.net_daily_revenue>=0?'var(--accent2)':'var(--danger)'},
    {val:'₹'+sf.staff_cost_per_shift.toLocaleString(), lab:'Staff Cost per Shift',           col:'var(--warn)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  mkChart('wi-ch-staff-count',{type:'bar',data:{labels:sf.slots.map(s=>s.slot),datasets:[
    {label:'Staff Needed',data:sf.slots.map(s=>s.staff_needed),
     backgroundColor:PAL.slice(0,4).map(c=>c+'cc'),borderRadius:6}
  ]},options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.y} staff`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D'},grid:{color:gC}}}}});

  mkChart('wi-ch-staff-profit',{type:'bar',data:{labels:sf.slots.map(s=>s.slot),datasets:[
    {label:'Slot Revenue',data:sf.slots.map(s=>s.slot_revenue),backgroundColor:'#2f8f6acc',borderRadius:4},
    {label:'Labour Cost', data:sf.slots.map(s=>s.labour_cost),  backgroundColor:'#7f9a24cc',borderRadius:4},
    {label:'Profit After Labour',data:sf.slots.map(s=>s.profit_after_labour),
     backgroundColor:sf.slots.map(s=>s.profit_after_labour>=0?'#2b6ea3cc':'#d2692ccc'),borderRadius:4},
  ]},options:{plugins:{legend:{labels:{color:'#55635D',boxWidth:10}},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  document.getElementById('wi-staff-tbody').innerHTML=sf.slots.map(s=>`
    <tr>
      <td style="font-weight:600;color:var(--text)">${s.slot}</td>
      <td>${s.footfall.toLocaleString()}</td>
      <td style="color:var(--accent)">${s.staff_needed}</td>
      <td style="color:var(--danger)">₹${s.labour_cost.toLocaleString()}</td>
      <td style="color:var(--accent2)">₹${s.slot_revenue.toLocaleString()}</td>
      <td style="color:${s.profit_after_labour>=0?'var(--accent2)':'var(--danger)'}">₹${s.profit_after_labour.toLocaleString()}</td>
    </tr>`).join('');

  // ══ TAB 6: TIERS ═════════════════════════════════════════════
  document.getElementById('wi-tier-kpis').innerHTML=[
    {val:tr.tier1_count.toLocaleString(),              lab:`Basic Members (${tr.tier1_name})`,    col:'var(--accent)'},
    {val:'₹'+tr.tier1_monthly_rev.toLocaleString(),   lab:'Basic Monthly Revenue',               col:'var(--accent2)'},
    {val:tr.tier2_count.toLocaleString(),              lab:`Premium Members (${tr.tier2_name})`,  col:'var(--purple)'},
    {val:'₹'+tr.tier2_monthly_rev.toLocaleString(),   lab:'Premium Monthly Revenue',             col:'var(--warn)'},
    {val:'₹'+tr.combined_monthly_rev.toLocaleString(),lab:'Combined Monthly Revenue',            col:'var(--accent2)'},
    {val:(tr.uplift_vs_single<0?'-₹':'₹')+Math.abs(tr.uplift_vs_single).toLocaleString(),    lab:'Uplift vs Single Tier',               col:tr.uplift_vs_single>=0?'var(--accent2)':'var(--danger)'},
  ].map(k=>`<div class="mini-kpi"><div class="mk-val" style="color:${k.col}">${k.val}</div><div class="mk-lab">${k.lab}</div></div>`).join('');

  mkChart('wi-ch-tier-split',{type:'doughnut',data:{
    labels:['Basic Tier','Premium Tier'],
    datasets:[{data:[tr.tier1_monthly_rev,tr.tier2_monthly_rev],
     backgroundColor:['#2f8f6acc','#2b6ea3cc'],borderWidth:2,borderColor:'#FFFFFF',hoverOffset:8}]
  },options:{plugins:{legend:{labels:{color:'#55635D'}},
    tooltip:{callbacks:{label:c=>`₹${c.parsed.toLocaleString()} (${c.label})`}}}}});

  mkChart('wi-ch-tier-compare',{type:'bar',data:{
    labels:['Single Tier (Avg Price)','Two-Tier Combined'],
    datasets:[{label:'Monthly Revenue (₹)',
      data:[tr.single_tier_monthly, tr.combined_monthly_rev],
      backgroundColor:['#2b6ea3cc','#2f8f6acc'],borderRadius:6}]
  },options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`₹${c.parsed.y.toLocaleString()}`}}},
    scales:{x:{ticks:{color:'#55635D'},grid:{color:gC}},y:{ticks:{color:'#55635D',callback:v=>'₹'+v.toLocaleString()},grid:{color:gC}}}}});

  document.getElementById('wi-tier-summary').innerHTML=`
    <strong style="color:var(--text)">Membership Tier Analysis</strong><br>
    Total members: <strong style="color:var(--accent)">${tr.total_members.toLocaleString()}</strong><br>
    Basic (₹${tr.tier1_price}/mo): <strong style="color:var(--accent2)">${tr.tier1_count.toLocaleString()} members → ₹${tr.tier1_monthly_rev.toLocaleString()}/mo</strong><br>
    Premium (₹${tr.tier2_price}/mo): <strong style="color:var(--purple)">${tr.tier2_count.toLocaleString()} members → ₹${tr.tier2_monthly_rev.toLocaleString()}/mo</strong><br>
    Combined two-tier monthly: <strong style="color:var(--accent2)">₹${tr.combined_monthly_rev.toLocaleString()}</strong>
    &nbsp;vs single-tier: <strong style="color:var(--muted2)">₹${tr.single_tier_monthly.toLocaleString()}</strong><br>
    ${tr.uplift_vs_single>=0
      ? `<strong style="color:var(--accent2)">Two-tier pricing adds ₹${tr.uplift_vs_single.toLocaleString()}/month</strong> vs a single average price point.`
      : `<strong style="color:var(--danger)">Two-tier pricing yields ₹${Math.abs(tr.uplift_vs_single).toLocaleString()}/month less</strong> — consider adjusting tier prices or uptake splits.`}
    <br>Annual combined: <strong style="color:var(--accent2)">₹${tr.combined_annual_rev.toLocaleString()}</strong>
  `;
}

// v41: the page no longer tells the server to shut down when the tab closes or refreshes.

// ════════════════════════════════════════════════════════════
//  UTILITY TILE — Button 1: Analyse New Dataset
// ════════════════════════════════════════════════════════════
async function analyseNewDataset(){
  if(!confirm('This will clear all current data and restart the app from the beginning. Are you sure?')) return;

  // Clear the backend _store so the reloaded page starts with no data
  try{ await fetch(BASE+'/reset', {method:'POST'}); }catch(e){}

  // Mark that this reload was triggered by Analyse New Dataset.
  // The reloaded page will detect this flag and wait for the Colab proxy
  // to re-establish the connection before allowing file upload.
  sessionStorage.setItem('_reloading', '1');
  sessionStorage.setItem('sc-new-dataset', '1');   // v54: no bundled icon after this

  window.location.reload();
}

// ════════════════════════════════════════════════════════════
//  v45 — REPORTS
//  The report is built from a list of content blocks read from each section
//  (headings, key figures, charts, tables, notes), then written as a PDF or
//  an interactive HTML page in the dashboard's light design.
//  • Main report  = Data summary + Exploratory analysis + Predictive analytics
//  • What-if report = the scenario currently on screen, downloaded separately
// ════════════════════════════════════════════════════════════
const _REPORTS = {
  main: {
    kind:'Analysis report', file:'study_cafe_report',
    sections:[
      {label:'Data summary',         screen:'s-summary',    deepOnly:true, load:()=>openSummary(),    ready:()=>document.getElementById('sum-body').style.display==='block'},
      {label:'Exploratory analysis', screen:'s-analytics',  load:()=>openAnalytics(),  ready:()=>!!(window._ANALYTICS&&Chart.getChart('ch-cluster-bar'))},
      {label:'Predictive analytics', screen:'s-predictive', load:()=>openPredictive(), ready:()=>document.getElementById('pred-body').style.display==='block'&&!!Chart.getChart('ch-mem-fi')},
    ]},
  whatif: {
    kind:'What-if scenario report', file:'study_cafe_whatif_report',
    sections:[
      {label:'What-if scenario', screen:'s-whatif', load:()=>openWhatIf(), ready:()=>!!window._lastWiData},
    ]},
};
let _exportScope='main';
const _sleep=ms=>new Promise(r=>setTimeout(r,ms));
const _frames=async n=>{ for(let i=0;i<n;i++) await new Promise(r=>requestAnimationFrame(r)); };

// ── Progress dialog helpers ────────────────────────────────
function _progOpen(title){
  document.getElementById('export-prog-title').textContent = title;
  document.getElementById('export-prog-fill').style.width  = '0%';
  document.getElementById('export-prog-pct').textContent   = '0%';
  document.getElementById('export-prog-step').textContent  = 'Starting...';
  document.getElementById('export-prog-log').innerHTML     = '';
  document.getElementById('export-prog-backdrop').classList.add('open');
}
function _progClose(){
  document.getElementById('export-prog-backdrop').classList.remove('open');
  const cb = document.querySelector('#export-prog-box .prog-close-btn');
  if(cb) cb.remove();
}
function _progSet(pct, step){
  document.getElementById('export-prog-fill').style.width = pct+'%';
  document.getElementById('export-prog-pct').textContent  = Math.round(pct)+'%';
  document.getElementById('export-prog-step').textContent = step;
}
function _progLog(msg){
  const log = document.getElementById('export-prog-log');
  log.innerHTML += msg + '<br>';
  log.scrollTop  = log.scrollHeight;
}

// ── Entry points ───────────────────────────────────────────
function _openExportModal(scope){
  _exportScope=scope;
  const m=document.getElementById('export-modal');
  const h=m.querySelector('h3'), p=m.querySelector('p');
  if(scope==='whatif'){
    if(h) h.textContent='Download the what-if scenario report';
    if(p) p.textContent='Covers the scenario currently on screen: its inputs, key figures and every tab. Run the scenario again first if you have changed any input.';
  } else {
    if(h) h.textContent='Download the analysis report';
    if(p) p.textContent='Choose how much detail to include. Both reports are written in plain words and open with the key findings. What-if scenarios have their own report, downloaded from the What-if section.';
  }
  const lv=document.getElementById('exp-level');
  if(lv){ lv.style.display=scope==='whatif'?'none':'';
    const cur=_view(); lv.querySelectorAll('input').forEach(i=>{ i.checked=(i.value===cur); }); }
  document.getElementById('export-modal-backdrop').classList.add('open');
}
function downloadAnalyses(){ _openExportModal('main'); }
function downloadWhatIfReport(){
  const note=document.getElementById('wi-report-note');
  if(!window._lastWiData){ if(note) note.textContent='Run a scenario first, then download its report.'; return; }
  if(note) note.textContent='';
  _openExportModal('whatif');
}
function _closeExportModal(){
  document.getElementById('export-modal-backdrop').classList.remove('open');
}
document.addEventListener('DOMContentLoaded', function(){
  document.querySelectorAll('.nav-tile').forEach(t=>{
    t.tabIndex=0; t.setAttribute('role','link');
    t.addEventListener('keydown',e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); t.click(); } });
  });
  const _bd=document.getElementById('export-modal-backdrop');
  if(_bd) _bd.addEventListener('click', function(e){ if(e.target === this) _closeExportModal(); });
  // What-if gets its own report button, next to "Save and compare"
  const sv=document.getElementById('wi-save-btn');
  if(sv && !document.getElementById('wi-report-btn')){
    const b=document.createElement('button');
    b.type='button'; b.id='wi-report-btn'; b.className='wi-run-btn'; b.textContent='Download scenario report';
    b.addEventListener('click', downloadWhatIfReport);
    sv.insertAdjacentElement('afterend', b);
    const n=document.createElement('div'); n.id='wi-report-note';
    n.style.cssText='font-size:13px;color:var(--danger);margin-top:8px;min-height:1em';
    (sv.parentElement.parentElement||sv.parentElement).appendChild(n);
  }
});

async function _runExport(format){
  const _lvIn=document.querySelector('#exp-level input:checked');
  _exportLevel=(_exportScope==='main')?(_lvIn?_lvIn.value:_view()):null;
  _closeExportModal();
  const spec=_REPORTS[_exportScope]||_REPORTS.main;
  const prevView=_view();
  _progOpen(format==='pdf' ? 'Building the PDF report...' : 'Building the HTML report...');
  const btns=[document.getElementById('dl-analyses-btn'),document.getElementById('wi-report-btn')].filter(Boolean);
  btns.forEach(b=>b.disabled=true);
  const prevScreen=(document.querySelector('.screen.active')||{}).id||'s-nav';
  const prevTheme=_theme();
  let ok=false;
  try{
    // Reports are always produced in the light design, whatever is on screen.
    if(prevTheme==='dark'){ _applyTheme('light', false); await _frames(2); }
    if(_exportLevel && _exportLevel!==prevView){ setView(_exportLevel, false); await _frames(3); }
    const model=await _collectReport(spec);
    if(!model.sections.some(s=>s.blocks.length)) throw new Error('Nothing to export yet.');
    if(format==='pdf') await _writePdf(model); else await _writeHtml(model);
    ok=true;
  } catch(e){
    _progLog('Failed: ' + (e&&e.message?e.message:e));
    _progSet(0, 'Export failed — see the log');
    const box = document.getElementById('export-prog-box');
    if(!box.querySelector('.prog-close-btn')){
      box.insertAdjacentHTML('beforeend',
        '<button class="prog-close-btn" onclick="_progClose()" '+
        'style="margin-top:14px;width:100%;padding:9px;border-radius:6px;'+
        'border:1px solid var(--border2);background:none;color:var(--muted2);'+
        'cursor:pointer;font-size:13px">Close</button>');
    }
  } finally {
    if(prevTheme==='dark') _applyTheme('dark', false);
    if(_exportLevel && _view()!==prevView){ try{ setView(prevView, false); }catch(e){} }
    _exportLevel=null;
    try{ goTo(prevScreen); }catch(e){}
    btns.forEach(b=>b.disabled=false);
    if(ok) setTimeout(_progClose, 1800);
  }
}

// ── Reading a section into content blocks ──────────────────
const _txt=el=>el?(el.innerText||el.textContent||'').replace(/\u00a0/g,' ').replace(/[ \t]+/g,' ').replace(/\s*\n\s*/g,'\n').trim():'';
const _one=el=>_txt(el).replace(/\n+/g,' ');
function _ownText(el){ // text of an element without its child chips/pills
  if(!el) return '';
  let s=''; el.childNodes.forEach(n=>{ if(n.nodeType===3) s+=n.textContent; });
  return s.replace(/\s+/g,' ').trim() || _one(el);
}
function _canvasToJpeg(cv, maxW){
  const w=cv.width, h=cv.height; if(!w||!h) return null;
  const k=Math.min(1,(maxW||1500)/w), oc=document.createElement('canvas');
  oc.width=Math.round(w*k); oc.height=Math.round(h*k);
  const c=oc.getContext('2d'); c.fillStyle='#FFFFFF'; c.fillRect(0,0,oc.width,oc.height); c.drawImage(cv,0,0,oc.width,oc.height);
  return {url:oc.toDataURL('image/jpeg',0.9), w:oc.width, h:oc.height};
}
function _chartConfig(inst){
  // A plain-data copy of the chart, so the HTML report can redraw it live.
  let cfg;
  try{ cfg=JSON.parse(JSON.stringify({type:inst.config.type, data:inst.config.data, options:inst.config.options},
        (k,v)=>typeof v==='function'?undefined:v)); }catch(e){ return null; }
  const id=inst.canvas.id, ds0=inst.data.datasets[0];
  if((id==='ch-corr'||id==='ch-arm-heatmap') && ds0 && cfg.data.datasets[0]){
    // these two are drawn with per-cell functions; bake the results in
    const data=ds0.data||[], xM=inst.options.scales.x.max+0.5, yM=inst.options.scales.y.max+0.5;
    cfg.data.datasets[0].pointStyle='rect';
    cfg.data.datasets[0].__cells=[xM,yM];
    cfg.data.datasets[0].backgroundColor=data.map(d=>{
      const v=d.v;
      return id==='ch-corr'?_corrColor(v):_armColor(v);
    });
    cfg.__axisLabels={x:[],y:[]};
    for(let i=0;i<xM;i++){ try{ cfg.__axisLabels.x.push((v=>Array.isArray(v)?v:String(v))(inst.options.scales.x.ticks.callback(i))); }catch(e){ cfg.__axisLabels.x.push(''); } }
    for(let i=0;i<yM;i++){ try{ cfg.__axisLabels.y.push(String(inst.options.scales.y.ticks.callback(i))); }catch(e){ cfg.__axisLabels.y.push(''); } }
  }
  try{ Object.keys(inst.options.scales||{}).forEach(k=>{
      const sc=inst.scales[k], cb=inst.options.scales[k].ticks&&inst.options.scales[k].ticks.callback;
      if(!sc||sc.type==='category'||typeof cb!=='function'||!cfg.options.scales||!cfg.options.scales[k]||cfg.__axisLabels) return;
      let out; try{ out=String(cb.call(sc,1234,0,[{value:1234}])); }catch(e){ return; }
      const m=out.match(/^([^\d-]*)([\d.,]+)(.*)$/); if(!m) return;
      const num=parseFloat(m[2].replace(/,/g,'')); let mul=null; if(Math.abs(num-1234)<1) mul=1; else if(Math.abs(num-123400)<1) mul=100; if(mul===null) return;
      cfg.options.scales[k].__fmt={pre:m[1],suf:m[3],mul:mul};
    }); }catch(e){}
  cfg.options=cfg.options||{}; cfg.options.responsive=true; cfg.options.maintainAspectRatio=false; cfg.options.animation=false;
  return cfg;
}
function _readGuide(card,id){
  const g=card.querySelector('.chart-guide[data-for="'+id+'"]'); if(!g) return null;
  return {p:_one(g.querySelector('.cg-p')), a:[...g.querySelectorAll('.cg-axes > div')].map(d=>[_one(d.querySelector('dt')),_one(d.querySelector('dd'))])};
}
const _WI_TABS={'wi-tab-revenue':'Revenue','wi-tab-breakeven':'Break-even','wi-tab-pricing':'Pricing','wi-tab-competitor':'Competitor','wi-tab-staffing':'Staffing','wi-tab-tiers':'Membership tiers','wi-tab-compare':'Saved scenarios compared'};
const _SKIP='.view-banner,.ps-controls,#pred-levers,.drill-bar,.card-tools,.lev-goto,.rule-callout,.zoom-hint,.hm-legend,.tool-menu,.chart-guide,.wi-help,.wi-assumptions,.tab-strip,button,.cluster-legend,#arm-all-rules,.topbar,script,style,#sum-loading,#pred-loading,#analytics-loading,.pred-loading,.view-more-btn,#wi-report-note,.spinner';

function _tableBlock(tbl, title){
  const head=[...tbl.querySelectorAll('thead th')].map(th=>_one(th).replace(/\s*[↕⇅]\s*$/,''));
  const rows=[...tbl.querySelectorAll('tbody tr')].map(tr=>[...tr.children].map(td=>_one(td)));
  // drop columns that are pure decoration (empty in every row, e.g. bar cells)
  const n=Math.max(head.length, ...rows.map(r=>r.length), 0), keep=[];
  for(let c=0;c<n;c++){ if(rows.some(r=>(r[c]||'')!=='')) keep.push(c); }
  const H=keep.map(c=>head[c]||''), R=rows.map(r=>keep.map(c=>r[c]||'')).filter(r=>r.some(x=>x!==''));
  if(!R.length) return null;
  if(R.length===1 && R[0].length===1) return null;      // "nothing saved yet" placeholders
  return {type:'table', title:title||'', head:H, rows:R};
}
function _walk(el, out){
  if(!el||el.nodeType!==1) return;
  if(el.matches(_SKIP)) return;
  const cs=getComputedStyle(el); if(cs.display==='none'||cs.visibility==='hidden') return;
  if(el.matches('.wi-tab-pane')){
    const before=out.length; out.push({type:'h3', text:_WI_TABS[el.id]||'Scenario'});
    [...el.children].forEach(c=>_walk(c,out));
    if(out.length===before+1) out.pop();                // tab with nothing in it
    return;
  }
  if(el.matches('.section-title,.pred-section-label') || (el.tagName==='H3' && el.parentElement && el.parentElement.matches('.sum-section,.wi-controls'))){
    const t=_one(el); if(t) out.push({type:'h3', text:t}); return;
  }
  if(el.matches('.wi-controls')){
    const rows=[...el.querySelectorAll('.wi-field')].map(f=>{ const l=f.querySelector('label'), i=f.querySelector('input,select'); return [_one(l), i?String(i.value):'', _one(f.querySelector('.wi-help'))]; }).filter(r=>r[0]);
    if(rows.length){ out.push({type:'h3', text:'Scenario inputs'}); out.push({type:'table', title:'', head:['Input','Value','What it means'], rows}); }
    const asm=_one(el.querySelector('.wi-assumptions')); if(asm) out.push({type:'text', title:'Fixed assumptions', text:asm});
    return;
  }
  if(el.matches('.mini-kpi-row,.kpi-grid')){
    const items=[...el.querySelectorAll('.mini-kpi,.kpi')].filter(k=>getComputedStyle(k).display!=='none').map(k=>({
      value:_one(k.querySelector('.mk-val,.k-val')), label:_one(k.querySelector('.mk-lab,.k-label')), sub:_one(k.querySelector('.k-sub'))})).filter(k=>k.value||k.label);
    if(items.length) out.push({type:'kpis', items}); return;
  }
  if(el.matches('.col-grid')||el.id==='col-grid'){
    const rows=[...el.querySelectorAll('.col-card')].map(c=>{
      const kv={}; const rest=[];
      c.querySelectorAll('.cc-row').forEach(r=>{ const k=_one(r.querySelector('.cc-k')), v=_one(r.querySelector('.cc-v'));
        if(/^(type|unique|missing)$/i.test(k)) kv[k.toLowerCase()]=v; else rest.push(k+': '+v); });
      return [_one(c.querySelector('.cc-name')), kv.type||'', kv.unique||'', kv.missing||'', rest.join('; ')];
    });
    if(rows.length) out.push({type:'table', title:'', head:['Column','Type','Unique','Missing','Details'], rows}); return;
  }
  if(el.matches('.rel-card')){
    out.push({type:'card', title:_ownText(el.querySelector('.rc-title')), text:_one(el.querySelector('.rc-desc')), value:_one(el.querySelector('.rc-val')),
              rows:[...el.querySelectorAll('.rc-row')].map(r=>[_one(r.querySelector('.rk')),_one(r.querySelector('.rv'))]).filter(r=>r[0]||r[1])}); return;
  }
  if(el.matches('.reg-insight')){
    const h=el.querySelector('h4'); let title=_ownText(h), body=[...el.querySelectorAll('p')].map(_one).join('\n');
    const pill=h?_one(h.querySelector('.pill')):''; if(pill) body=pill+'. '+body;
    const extra=[...el.querySelectorAll('.coef-row')].map(r=>_one(r.querySelector('.cr-name'))+': '+_one(r.querySelector('.cr-val')));
    if(extra.length) body+=(body?'\n':'')+extra.join('   ');
    if(!body) body=_txt(el).replace(title,'').trim();
    if(title||body) out.push({type:'text', title, text:body}); return;
  }
  if(el.matches('.wi-note')||el.id==='arm-note'||el.id==='relevant-info'||el.matches('.pred-intro')){
    const t=_one(el); if(t) out.push({type:'text', title:'', text:t}); return;
  }
  if(el.matches('.lever-note')){ const t=_one(el); if(t) out.push({type:'text', title:'Levers', text:t}); return; }
  if(el.matches('.chart-card')){
    el.querySelectorAll('.card-explain').forEach(x=>{ const t=_txt(x); if(t) out.push({type:'text', title:x.dataset.title||'', text:t}); });
    const title=_one(el.querySelector('.cc-title')), sub=_one(el.querySelector('.cc-sub'));
    const cvs=[...el.querySelectorAll('canvas')], tbls=[...el.querySelectorAll('table')]; let any=false;
    const wide=el.getBoundingClientRect().width > (el.parentElement?el.parentElement.getBoundingClientRect().width*0.7:9999);
    cvs.forEach(cv=>{
      if(cv.id==='ch-scatter3d' && window._sc3d){
        const prevFill=window._sc3d.o.fill, prevHint=window._sc3d.o.hint; window._sc3d.o.fill=true; window._sc3d.o.hint=false; window._sc3d.hover=null; window._sc3d.draw();
        const im=_canvasToJpeg(cv); window._sc3d.o.fill=prevFill; window._sc3d.o.hint=prevHint; window._sc3d.draw();
        if(im){ out.push({type:'chart3d', title, sub:'PC1, PC2 and PC3, explained in the table that follows. In the HTML report, drag to rotate.', img:im, wide:true,
                 data:{points:window._sc3d.o.points, names:window._sc3d.o.names, axes:window._sc3d.o.axes}, guide:_readGuide(el,cv.id)}); any=true; }
        return;
      }
      const inst=Chart.getChart(cv); if(!inst||cv.width<40||cv.height<40) return;
      if(!inst.data.datasets.length || !inst.data.datasets.some(d=>d.data&&d.data.length)) return;
      const im=_canvasToJpeg(cv); if(!im) return;
      out.push({type:'chart', title, sub, img:im, wide, cfg:_chartConfig(inst), kind:inst.config.type, guide:_readGuide(el,cv.id)}); any=true;
    });
    if(!cvs.length && tbls.length && sub && el.id==='pca-explain'){ out.push({type:'text', title, text:sub}); any=true; }
    tbls.forEach(t=>{ const b=_tableBlock(t, any?'':title); if(b){ out.push(b); any=true; } });
    const cg=el.querySelector('.conf-grid');
    if(cg){ const v=[...cg.querySelectorAll('.conf-cell .cv')].map(_one);
      if(v.length===4){ out.push({type:'table', title:title+(sub?' ('+sub+')':''), head:['','Predicted: No','Predicted: Yes'], rows:[['Actual: No',v[0],v[1]],['Actual: Yes',v[2],v[3]]]});
        const gd=_readGuide(el,'mem-conf-matrix'); if(gd) out.push({type:'text', title:'About this table', text:gd.p+'\n'+gd.a.map(r=>r[0]+': '+r[1]).join('\n')}); return; } }
    el.querySelectorAll('.lever-note').forEach(n=>{ const t=_one(n); if(t){ out.push({type:'text', title:'Levers', text:t}); any=true; } });
    if(!cvs.length && !tbls.length){
      let body=_txt(el); if(title && body.indexOf(title)===0) body=body.slice(title.length).trim(); if(sub && body.indexOf(sub)===0) body=body.slice(sub.length).trim();
      if(body) out.push({type:'text', title, text:body});
    }
    return;
  }
  if(el.tagName==='TABLE'){
    const holder=el.closest('.sum-section'); const b=_tableBlock(el,''); if(b) out.push(b); return;
  }
  [...el.children].forEach(c=>_walk(c,out));
}

async function _collectReport(spec){
  const isMain=spec===_REPORTS.main, level=isMain?(_exportLevel||_view()):null;
  const model={kind:spec.kind+(isMain?(level==='deep'?' (deep view)':' (surface view)'):''),
               file:spec.file+(isMain?'_'+level:''), date:new Date(), sections:[], facts:[], level};
  const secs=spec.sections.filter(s=>!(s.deepOnly&&level==='surface'));
  const n=secs.length;
  for(let si=0; si<n; si++){
    const sec=secs[si], base=5+si*(60/n);
    _progSet(base, 'Preparing '+sec.label+'...'); _progLog('Section: '+sec.label);
    goTo(sec.screen);
    if(!sec.ready()){ try{ await sec.load(); }catch(e){ _progLog('  could not load: '+e.message); } }
    const deadline=Date.now()+300000;
    while(!sec.ready() && Date.now()<deadline) await _sleep(250);
    if(!sec.ready()){ _progLog('  not ready, skipped'); continue; }
    goTo(sec.screen);
    const screenEl=document.getElementById(sec.screen);
    try{ _interactResetAll(); }catch(e){}
    // Build charts that are only created when their tab is opened
    if(sec.screen==='s-analytics'){
      for(const t of ['bubble','scatter','bar','heatmap']){ armTab(t); await _sleep(220); }
      armTab('table');
    }
    // Show every tab's content at once so each chart can size itself, then read it
    const revealed=[];
    screenEl.querySelectorAll('#arm-bubble-view,#arm-scatter-view,#arm-bar-view,#arm-heatmap-view,#arm-table-view,.wi-tab-pane').forEach(v=>{
      revealed.push([v, v.style.display]); v.style.display='block'; });
    await _frames(3);
    Object.values(Chart.instances).forEach(ch=>{ if(screenEl.contains(ch.canvas)){ try{ ch.resize(); ch.update('none'); }catch(e){} } });
    if(window._sc3d && screenEl.contains(window._sc3d.canvas)) window._sc3d.draw();
    await _frames(3); await _sleep(350);
    _progSet(base+(40/n), 'Reading '+sec.label+'...');
    let blocks=[]; _walk(screenEl, blocks);
    blocks=blocks.filter((b,i)=>{ const nx=blocks[i+1];
      if(b.type==='table'&&nx&&nx.type==='table'&&nx.rows.length>=b.rows.length&&JSON.stringify(nx.head)===JSON.stringify(b.head)
         &&JSON.stringify(nx.rows.slice(0,b.rows.length))===JSON.stringify(b.rows)){ if(!nx.title) nx.title=(b.title||'').replace(/^Top /,'All ').replace(/^All a/,'All a'); return false; }
      return true; });
    if(sec.screen==='s-predictive' && typeof _levActive==='function' && _levActive()){
      const st=document.getElementById('lev-status'); if(st&&_one(st)) blocks.unshift({type:'text', title:'Drill-down applied', text:_one(st)}); }
    revealed.forEach(([v,d])=>{ v.style.display=d; });
    if(sec.screen==='s-analytics') armTab('table');
    model.sections.push({label:sec.label, blocks});
    _progLog('  '+blocks.filter(b=>b.type==='chart'||b.type==='chart3d').length+' charts, '+blocks.filter(b=>b.type==='table').length+' tables');
  }
  let S=null;
  try{ const r=await fetch(BASE+'/summary_data'); const d=await r.json();
    if(!d.error){ S=d; model.facts.push(d.n_final.toLocaleString()+' of '+d.n_start.toLocaleString()+' survey responses kept after cleaning'); } }catch(e){}
  if(isMain){ try{ _plainLayer(model, S); }catch(e){ console.warn('plain layer', e); _progLog('  plain-language notes skipped: '+e.message); } }
  return model;
}

// ════════ v57 PLAIN-LANGUAGE REPORTS ════════
let _exportLevel=null;
const _n0=v=>Math.round(+v).toLocaleString('en-IN');
const _p0=v=>Math.round(+v)+'%';
function _band(k){ const t=String(k), n=t.match(/\d+/g)||[];
  if(/^>/.test(t)) return 'over ₹'+n[0]; if(/^<|^≤/.test(t)) return 'up to ₹'+n[0]; return n.length>=2?'₹'+n[0]+'–'+n[1]:t; }
function _sharesOf(obj){ const t=Object.values(obj||{}).reduce((a,b)=>a+(+b||0),0)||1;
  return Object.entries(obj||{}).map(([k,v])=>[k,(+v||0)/t*100]).sort((a,b)=>b[1]-a[1]); }
const _lc=s=>String(s).split(' ').map(w=>/[A-Z].*[A-Z]/.test(w)?w:w.toLowerCase()).join(' ');
const _cap=s=>{ s=String(s); return s.charAt(0).toUpperCase()+s.slice(1); };
function _inrWords(v){ v=+v; const a=Math.abs(v); if(a>=1e7) return '₹'+(v/1e7).toFixed(2)+' crore'; if(a>=1e5) return '₹'+(v/1e5).toFixed(1)+' lakh'; return '₹'+Math.round(v).toLocaleString('en-IN'); }
function _freqWords(l){ const s=String(l);
  if(/^daily$/i.test(s)) return 'daily visitors'; if(/^rarely$/i.test(s)) return 'students who rarely visit';
  if(/^1$/.test(s)) return 'once-a-week visitors'; if(/^\d/.test(s)) return 'students who visit '+s.replace('-','–')+' times a week'; return '"'+s+'"'; }

function _plainFindings(level, S){
  const F=[], P=window._PREDDATA||{}, VC=window._VC||{}, PR=window._PRICING||{}, CL=window._CLUSTERS||{};
  const add=(title,text)=>{ if(text) F.push({type:'text', title, text}); };
  const tryAdd=fn=>{ try{ fn(); }catch(e){ console.warn('finding',e); } };
  tryAdd(()=>{ if(!S||!S.n_final) return;
    const why={duplicates:'duplicates',outliers:'impossible values',range_invalid:'answers outside the allowed options',straight_line:'the same score for every rating',
               late_night:'contradictory answers',spend_budget:'contradictory answers',contradiction:'contradictory answers',excessive_missing:'mostly blank',non_target:'outside the target group'};
    const reasons=[]; Object.entries(S.clean_log||{}).forEach(([k,v])=>{ if(!/removed/.test(k)||!v||!(v.count>0)) return;
      const key=Object.keys(why).find(w=>k.indexOf(w)>=0); if(!key) return; const r=why[key]; if(!reasons.includes(r)) reasons.push(r); });
    add('Who these results describe', `${_n0(S.n_final)} of the ${_n0(S.n_start)} survey responses were usable after cleaning. `
      +`${_n0(S.n_start-S.n_final)} were removed`+(reasons.length?` (${reasons.join(', ')})`:'')+`. Every figure here describes these ${_n0(S.n_final)} students. `
      +`They said what they would do, so read the figures as intentions, not sales.`); });
  tryAdd(()=>{ const d=(P.peak_hours||{}).distribution; if(!d) return; const e=Object.entries(d).sort((a,b)=>b[1]-a[1]); const hi=e[0], lo=e[e.length-1];
    add('When students would come', `${hi[0]} is the most popular time: ${_p0(hi[1])} of students prefer it, against ${_p0(lo[1])} for ${lo[0].toLowerCase()}. `
      +`Put the most staff and seats into ${hi[0].toLowerCase()} and keep ${lo[0].toLowerCase()} lean.`); });
  tryAdd(()=>{ const d=(P.seating||{}).overall_pct; if(!d) return; const e=Object.entries(d).sort((a,b)=>b[1]-a[1]);
    add('Where they would sit', `${_cap(_lc(e[0][0]))} (${_p0(e[0][1])}) and ${_lc(e[1][0])} (${_p0(e[1][1])}) are the most wanted seats, together ${_p0(e[0][1]+e[1][1])} of students. `
      +`${_cap(_lc(e[e.length-1][0]))} is wanted by only ${_p0(e[e.length-1][1])}. Give most of the floor to the first two.`); });
  tryAdd(()=>{ const b=_sharesOf(VC.Beverage_Preference), f=_sharesOf(VC.Food_Preference); if(!b.length||!f.length) return;
    add('What they want to eat and drink', `${_cap(_lc(b[0][0]))} is the most popular drink (${_p0(b[0][1])} of students), then ${_lc(b[1][0])} (${_p0(b[1][1])}). `
      +`${_cap(_lc(f[0][0]))} are the most wanted food (${_p0(f[0][1])}). Build the menu around these first.`); });
  tryAdd(()=>{ const m=P.membership||{}; if(m.overall_rate==null) return; const t=Object.entries(m.by_spend_tier||{}).sort((a,b)=>a[1]-b[1]);
    let s=`${_p0(m.overall_rate)} of students said they would take a monthly membership.`;
    if(t.length>=2) s+=` Interest rises with spending: ${_p0(t[t.length-1][1])} of students who spend ${_band(t[t.length-1][0])} a visit would join, against ${_p0(t[0][1])} of those who spend ${_band(t[0][0])}. Offer memberships to your bigger spenders first.`;
    add('Interest in a membership', s); });
  tryAdd(()=>{ const c=PR.membership; if(!c) return; const i=c.points.findIndex(p=>p.price===c.best_price); const b=c.points[i], nx=c.points[i+1];
    add('What to charge for a membership', `A fee of about ${_INR(b.price)} a month would bring in the most from interested students: ${_p0(b.share*100)} of them said they would still join at that price.`
      +(nx?` At ${_INR(nx.price)}, only ${_p0(nx.share*100)} would.`:'')+' People usually say they would pay more than they do, so try the fee on a small group before committing.'); });
  tryAdd(()=>{ const c=PR.visit; if(!c) return; const k=c.elasticity.findIndex(e=>e.e<-1.2); if(k<0) return; const a=c.points[k], b=c.points[k+1];
    add('What a visit can cost', `Most students can pay up to about ${_INR(a.price)} a visit (${_p0(a.share*100)} of them). Above that, many drop out: only ${_p0(b.share*100)} can pay ${_INR(b.price)}.`); });
  tryAdd(()=>{ const h=PR.headroom; if(!h) return; const s=h.shares;
    add('Room to raise prices', `${_p0((s[1]+s[2])*100)} of students could spend more than they do now, but ${_p0(s[0]*100)} are already at their limit. A general price rise would lose those students first, so raise prices on extras rather than the basics.`); });
  tryAdd(()=>{ const S2=PR.segments; if(!S2||!S2.groups.length) return; const g=S2.groups.find(x=>/Visit_Frequency/.test(x.key))||S2.groups[0];
    const r=g.rows.slice().sort((a,b)=>b.share-a.share), hi=r[0], lo=r[r.length-1];
    add('Who would pay more', `${/Visit_Frequency/.test(g.key)?_freqWords(hi.label).replace(/^./,c=>c.toUpperCase()):'"'+hi.label+'"'} are the best prospects for a premium membership: ${_p0(hi.share*100)} would pay ${_INR(S2.threshold)} or more a month, against ${_p0(lo.share*100)} of ${/Visit_Frequency/.test(g.key)?_freqWords(lo.label):'"'+lo.label+'"'}.`); });
  tryAdd(()=>{ const cm=CL.cluster_means||[]; if(!cm.length) return; const tot=cm.reduce((a,c)=>a+c.size,0)||1;
    const big=cm.slice().sort((a,b)=>b.size-a.size)[0], keen=cm.slice().sort((a,b)=>b.membership-a.membership)[0];
    add('The kinds of students', `The students fall into ${cm.length} groups with similar habits. The largest is "${big.name}" (${_p0(big.size/tot*100)} of students)`
      +(keen===big?`, and it is also the keenest on membership: ${_p0(keen.membership*100)} of its students would join. `
                  :`. The group keenest on membership is "${keen.name}", where ${_p0(keen.membership*100)} would join. `)
      +'Design the space and offers around these groups.'); });
  tryAdd(()=>{ const R=(window._RULES||[]).filter(r=>r.significant!==false&&r.lift>=1.2); if(!R.length) return;
    const rank=R.slice().sort((a,b)=>(b.n_both||0)*(1-1/b.lift)-(a.n_both||0)*(1-1/a.lift));
    const top=rank.find(r=>!/:\s*No$/i.test(String(r.consequent)))||rank[0];
    add('A pattern worth acting on', `Students who answered "${top.antecedent}" are ${(+top.lift).toFixed(1)} times as likely as average to also answer "${top.consequent}". `
      +'The "Rules worth acting on" table turns patterns like this into actions.'); });
  tryAdd(()=>{ const f=P.footfall||{}, r=P.revenue||{}; if(!f.daily_proj_1000) return;
    add('Visits and takings to expect', `With 1,000 students nearby, the survey points to about ${_n0(f.daily_proj_1000)} visits a day`
      +(r.daily_revenue?` and about ${_inrWords(r.daily_revenue)} in takings a day (${_inrWords(r.monthly_revenue)} a month)`:'')
      +', if the café has room for them all. The What-if section works out what your own seats, prices and costs would give.'); });
  if(level==='deep'){
    tryAdd(()=>{ const m=P.membership||{}, ms=P.model_summary||{}; if(m.accuracy==null) return;
      let s=`The membership model is right about ${Math.round(m.accuracy*10)} times in 10, against ${Math.round((m.baseline||0.5)*10)} in 10 by always guessing the most common answer, so it is useful for spotting likely members.`;
      const rg=(ms.regression||[])[0]; if(rg&&rg.metric&&/R/.test(rg.metric)) s+=` The ${String(rg.task).toLowerCase()} forecast explains about ${_p0(rg.score*100)} of the differences between students, so treat it as rough.`;
      const ph=P.peak_hours||{}; if(ph.rf_accuracy!=null&&ph.baseline!=null) s+=` Predicting one student's favourite time is hard (right ${_p0(ph.rf_accuracy*100)} of the time against ${_p0(ph.baseline*100)} by guessing), so plan from the overall shares instead.`;
      add('How far to trust the forecasts', s); });
    tryAdd(()=>{ const rp=(S&&(S.reliability_post||S.reliability_pre))||{}, a=(rp.cronbach_alpha||{}).value, c=(rp.completeness||{}).completeness_pct; if(a==null) return;
      add('How trustworthy the answers are', 'After cleaning: '+(a<0.7?`the three facility ratings (power outlets, Wi-Fi, productivity tools) do not measure one shared idea (Cronbach's alpha ${(+a).toFixed(2)}, below the usual 0.70), so each is read on its own.`
        :`the three facility ratings hang together well (Cronbach's alpha ${(+a).toFixed(2)}), so they can be read as one "facilities" score.`)
        +(c!=null?` ${(+c).toFixed(1)}% of all answer boxes are filled in. (The Data summary shows the same checks on the raw survey, before cleaning.)`:'')); });
  }
  return F;
}

// one plain sentence under each heading
const _LEADS=[
  [/^Data summary$/,'How the raw survey was checked and cleaned before any analysis, and whether the answers can be trusted.'],
  [/^Exploratory analysis$/,'What the survey says about the students: what they want, how their answers relate, which groups they form and what they would pay.'],
  [/^Predictive analytics$/,'Forecasts built from the survey: when the café will be busy, how many visits and how much money to expect, which seats are needed and who is likely to buy a membership.'],
  [/^Overview$/,'The size of the survey before and after cleaning.'],
  [/^Relevant data/i,'Which survey questions the analysis uses.'],
  [/^Cleaning techniques/i,'Each fix made to the raw answers, how many rows it affected and why.'],
  [/^Data reliability/i,'Checks on whether the answers are complete, consistent and given with care. These checks were run on the raw survey, before cleaning.'],
  [/^Column profiles/i,'A short profile of every survey question: what kind of answer it takes, how many different answers it got and how many are blank.'],
  [/^What students want|^Data summary — distributions/i,'What students want from a study café: when they would come, where they would sit, and what they would eat and drink.'],
  [/^Patterns that shape|cross-tab/i,'How answers relate to each other, for example which seats group studiers prefer. These shape the layout and the opening hours.'],
  [/^Correlation/i,'Which answers tend to rise and fall together. A strong link does not prove that one causes the other.'],
  [/^Regression/i,'Which answers best predict whether a student would take a membership, each one judged with the others held steady.'],
  [/^Customer groups|k-means/i,'The students split into a few groups with similar habits. Each group needs slightly different things from the café.'],
  [/^Pricing and willingness/i,'What students said they would pay, and how many would still come at higher prices.'],
  [/^Association rules|^What tends to go together/i,'Answers that often appear together, such as students who visit daily also wanting a membership.'],
  [/^Peak hours/i,'When in the day the café will be busiest, and what each time of day brings in.'],
  [/^Footfall/i,'How many visits a day to expect, and how they spread over the week.'],
  [/^Revenue forecast/i,'How much money the expected visits would bring in, now and over the year.'],
  [/^Seating demand/i,'How many of each kind of seat the café will need, and how that changes through the day.'],
  [/^Membership willingness/i,'Who is likely to buy a membership, and which answers drive that.'],
  [/^Model performance/i,'How reliable each forecast is, compared with simple guessing.'],
];
function _leadFor(t){ const m=_LEADS.find(x=>x[0].test(String(t).trim())); return m?m[1]:null; }

// plain titles and explanations for the data-reliability cards
function _plainCard(b){
  const t=String(b.title||'').replace(/^[^A-Za-z]+/,''), v=parseFloat(String(b.value||'').replace(/[^\d.\-]/g,''));
  const row=k=>{ const r=(b.rows||[]).find(x=>new RegExp(k,'i').test(x[0])); return r?r[1]:null; };
  if(/^Cronbach/i.test(t)){ b.title="Do the facility ratings agree? (Cronbach's alpha)";
    b.text='Checks whether the three facility ratings (power outlets, Wi-Fi, productivity tools) measure one shared idea. 0.70 or more would mean they do. '
      +(isNaN(v)?'':(v<0.7?`At ${v.toFixed(2)} they don't, so each rating is read on its own.`:`At ${v.toFixed(2)} they do, so they can be combined.`)); return; }
  if(/^Inter-?item/i.test(t)){ b.title='How closely the facility ratings move together';
    b.text='How closely the answers to each pair of facility ratings move together, from 0 (not at all) to 1 (exactly). Values around 0.15 to 0.50 mean the questions are related but ask about different things.'; return; }
  if(/^Shapiro/i.test(t)){ b.title='Do ages follow a bell curve? (Shapiro-Wilk test)';
    const nm=row('normal'); b.text='Tests whether ages follow a bell-shaped curve. '+(nm&&/no/i.test(nm)?'They do not, which is normal for a student survey where ages bunch at the young end. It does not affect the other results.':'They do.'); return; }
  if(/^Completeness/i.test(t)){ b.title='How complete the answers are';
    b.text='The share of all answer boxes that are filled in. Above 85% is usually considered fine.'+(isNaN(v)?'':` Here it is ${v}%.`); return; }
  if(/^Response bias/i.test(t)){ b.title='Careless answering';
    b.text='Looks for students who answered without thinking: giving every rating the same score (straight-lining), always agreeing (4 or 5) or always disagreeing (1 or 2). A small share in a large survey is normal.'; return; }
  if(/^Spend\s*>\s*Budget/i.test(t)){ b.title='Spending more than their own budget';
    const rm=row('removed'); b.text='Students who said they spend more per visit than their own budget allows. Their answers contradict each other'+(rm&&/yes/i.test(rm)?', so they were removed.':'.'); return; }
  if(/^High-?missing/i.test(t)){ b.title='Questions with many blank answers';
    b.text='Questions where more than 30% of answers are blank would need care.'+(/none/i.test(String(b.text||'')+String(b.value||''))?' There are none.':''); return; }
}

// plain table headings and values
const _HEAD_PLAIN={'Antecedent':'If a student answered…','Consequent':'…they also tend to answer','Support':'Share of all students','Confidence':'How often it holds',
  'Lift':'Times more likely than average (lift)','Adjusted p-value':'Chance it is a fluke (adjusted p-value)','Significant?':'Reliable?','Unique':'Different answers',
  'Type':'Kind of answer','Missing':'Blank answers','Rows Affected':'Rows affected','Method / Rationale':'What was done and why'};
const _TYPE_PLAIN={object:'text or category',int64:'whole number',float64:'number',bool:'yes or no',category:'category'};
function _plainTable(b){
  const typeCol=b.head.indexOf('Type'), isProfile=typeCol>=0&&b.head.includes('Unique');
  if(isProfile) b.rows.forEach(r=>{ r[typeCol]=_TYPE_PLAIN[r[typeCol]]||r[typeCol]; });
  const mi=b.head.indexOf('Metric'), si=b.head.indexOf('Score');
  if(mi>=0&&si>=0){ b.head=b.head.concat(['In plain words']);
    b.rows=b.rows.map(r=>{ const m=String(r[mi]), s=parseFloat(r[si]); let w='';
      if(/accuracy/i.test(m)&&!isNaN(s)) w=`right about ${Math.round((s>1?s/100:s)*10)} times in 10`;
      else if(/R²|R2/.test(m)&&!isNaN(s)) w=`explains about ${Math.round(s*100)}% of the differences between students`;
      else if(/F1/i.test(m)&&!isNaN(s)) w='balance of correct and missed predictions; 1 is perfect';
      else if(/MAE|error/i.test(m)) w='typical size of a miss';
      return r.concat([w]); }); }
  b.head=b.head.map(h=>_HEAD_PLAIN[h]||h);
}

// glossary: only the words that appear in this report
const _GLOSS=[
  ['95% range',/95% range/,'The range the true figure very likely falls in, given how many students answered. A narrow range means a precise figure.'],
  ['Statistically significant',/significan|Reliable\?/i,'Unlikely to be a coincidence: there is less than a 5% chance the pattern appeared by luck alone.'],
  ['p-value',/p-value|\bp below|\bp = /,'The chance of seeing a pattern this strong by pure luck. The smaller, the more convincing; below 0.05 is the usual bar.'],
  ['Elasticity',/elasticit/i,'How strongly demand reacts to price. −0.3 means a 10% price rise loses about 3% of buyers; −2 means it loses about 20%.'],
  ['Correlation',/correlation/i,'How closely two answers move together, from −1 (opposite) through 0 (unrelated) to 1 (in step). It does not show that one causes the other.'],
  ['Regression',/regression/i,'A method that estimates how much each answer, on its own, raises or lowers the chance of something, such as taking a membership.'],
  ['Coefficient',/coefficient/i,'The size and direction of one answer\'s effect in a regression. Positive raises the chance, negative lowers it.'],
  ['K-means clustering',/k-means|cluster/i,'A method that sorts students into groups with similar answers, so each group can be described and served on its own.'],
  ['PCA (PC1, PC2, PC3)',/\bPC1\b|\bPCA\b/,'A way to squeeze many answers into two or three summary scores, so the groups can be drawn on one chart.'],
  ['Association rule',/association rule|If a student answered/i,'A pattern of answers that often appear together, such as "students who visit daily also want a membership".'],
  ['Support',/Share of all students|\bsupport\b/i,'How common a pattern is: the share of all students who match it.'],
  ['Confidence',/How often it holds|\bconfidence\b/i,'How often a rule holds: of the students who match the first part, the share who also match the second.'],
  ['Lift',/\blift\b/i,'How many times more likely the second answer is among students who match the first, compared with everyone. 1 means no link; 1.5 means 50% more likely.'],
  ['Accuracy',/accuracy/i,'The share of students a model classifies correctly.'],
  ['Baseline',/baseline|by guessing|by always guessing/i,'What you would get by always guessing the most common answer. A model is only useful if it beats this.'],
  ['R²',/R²|\bR2\b/,'How much of the differences between students a model explains, from 0 (none) to 1 (all).'],
  ['Feature importance',/importance/i,'How much a model relies on each answer. Higher means the answer matters more for the prediction.'],
  ["Cronbach's alpha",/cronbach/i,'A check on whether several questions measure one shared idea. 0.70 or more means they do.'],
  ['Shapiro-Wilk test',/shapiro/i,'A check on whether numbers, here ages, follow a bell-shaped curve.'],
  ['Likert scale',/likert/i,'A 1-to-5 rating question, from "not important" to "very important".'],
  ['Precision and recall',/precision|recall/i,'Precision: of the students a model says will join, the share who would. Recall: of the students who would join, the share the model finds.'],
  ['Confusion matrix',/confusion|Predicted: /i,'A table of a model\'s right and wrong calls: correct yeses, correct noes, false alarms and misses.'],
  ['Footfall',/footfall/i,'The number of visits to the café.'],
  ['Catchment',/catchment/i,'The students who live or study close enough to be likely customers.'],
  ['Straight-lining',/straight-?lin/i,'Giving the same score to every rating question, a sign of answering without reading.'],
];
function _reportText(model){
  const parts=[];
  model.sections.forEach(s=>s.blocks.forEach(b=>{
    ['title','text','sub','value'].forEach(k=>{ if(b[k]) parts.push(String(b[k])); });
    if(b.head) parts.push(b.head.join(' ')); if(b.rows) parts.push(b.rows.map(r=>r.join(' ')).join(' '));
    if(b.items) parts.push(b.items.map(i=>i.label+' '+i.value).join(' '));
    if(b.guide){ parts.push(b.guide.p||''); (b.guide.a||[]).forEach(a=>parts.push(a.join(' '))); }
  }));
  return parts.join(' \n ');
}

function _plainLayer(model, S){
  // 1. leads, plain cards and plain tables inside every section
  model.sections.forEach(sec=>{
    const out=[]; const l0=_leadFor(sec.label); if(l0) out.push({type:'lead', text:l0});
    sec.blocks.forEach(b=>{
      if(b.type==='card') _plainCard(b);
      if(b.type==='table') _plainTable(b);
      out.push(b);
      if(b.type==='h3'){ const l=_leadFor(b.text); if(l) out.push({type:'lead', text:l}); }
    });
    sec.blocks=out;
  });
  // 2. key findings first
  const F=_plainFindings(model.level, S);
  if(F.length) model.sections.unshift({label:'Key findings in plain words', coverNote:F.length+' findings',
    blocks:[{type:'lead', text:(model.level==='deep'
      ?'The main results of the full analysis, in plain words. The sections that follow show the evidence, and every technical term is explained in the list at the end.'
      :'What the survey means for running the café, in plain words. The sections that follow show the charts behind each point.')}].concat(F)});
  // 3. glossary of the words this report uses
  const txt=_reportText(model), G=_GLOSS.filter(g=>g[1].test(txt));
  if(G.length) model.sections.push({label:'Words used in this report', coverNote:G.length+' terms',
    blocks:[{type:'lead', text:'Plain meanings for the technical words that appear in this report.'},
            {type:'table', title:'', head:['Word','What it means'], rows:G.map(g=>[g[0],g[2]])}]});
  model.facts.unshift(model.level==='deep'?'Deep report: the full analysis, with every technique explained in plain words.'
                                         :'Surface report: what matters for running the café, in plain words.');
}

// ── PDF writer ─────────────────────────────────────────────
function _pdfSafe(s){
  return String(s==null?'':s).replace(/₹\s?/g,'Rs ').replace(/[→⟹⇒]/g,'->').replace(/↔/g,'<->').replace(/≥/g,'>=').replace(/≤/g,'<=').replace(/R²/g,'R-squared').replace(/²/g,'2').replace(/[−–](?=\s?\d|\s?Rs|\s?₹)/g,'-').replace(/−/g,'-')
    .replace(/[–—]/g,'-').replace(/[“”]/g,'"').replace(/[‘’]/g,"'").replace(/✅\s?/g,'').replace(/❌\s?/g,'').replace(/⚠\s?/g,'').replace(/…/g,'...')
    .replace(/[^\x09\x0A\x20-\x7E\u00A0-\u00FF]/g,'').replace(/[ \t]+/g,' ').trim();
}
async function _writePdf(model){
  _progSet(68, 'Loading the PDF library...');
  if(typeof window.jspdf === 'undefined'){
    await new Promise((res,rej)=>{
      const s=document.createElement('script');
      s.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
      s.onload=res; s.onerror=()=>rej(new Error('The PDF library could not be loaded (check the internet connection).')); document.head.appendChild(s);
    });
  }
  const {jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'portrait', unit:'pt', format:'a4', compress:true});
  const PW=595.28, PH=841.89, M=44, CW=PW-M*2, BOT=PH-52;
  const INK=[23,35,31], SLATE=[85,99,93], RULE=[220,227,221], GREEN=[31,92,74], PAPER=[245,247,244];
  const T=(size,style,col)=>{ doc.setFont('helvetica',style||'normal'); doc.setFontSize(size); doc.setTextColor(...(col||INK)); };
  let y=M;
  const newPage=()=>{ doc.addPage(); y=M; };
  const need=h=>{ if(y+h>BOT) newPage(); };
  const lines=(s,w,size,style)=>{ doc.setFont('helvetica',style||'normal'); doc.setFontSize(size||9); return doc.splitTextToSize(_pdfSafe(s), w); };

  // cover
  doc.setFillColor(...GREEN); doc.rect(0,0,PW,250,'F');
  T(30,'bold',[255,255,255]); doc.text('Study café feasibility', M, 150);
  T(14,'normal',[255,255,255]); doc.text(_pdfSafe(model.kind), M, 178);
  T(10,'normal',[214,232,224]); doc.text(model.date.toLocaleDateString('en-GB',{year:'numeric',month:'long',day:'numeric'}), M, 204);
  y=300; T(11,'bold'); doc.text('In this report', M, y); y+=10;
  doc.setDrawColor(...RULE); doc.setLineWidth(0.8); doc.line(M,y,PW-M,y); y+=20;
  model.sections.forEach((s,i)=>{ T(11,'normal'); doc.text((i+1)+'.  '+_pdfSafe(s.label), M, y);
    const nc=s.blocks.filter(b=>b.type==='chart'||b.type==='chart3d').length, nt=s.blocks.filter(b=>b.type==='table').length;
    T(9,'normal',SLATE); doc.text(s.coverNote||(nc+' charts, '+nt+' tables'), PW-M, y, {align:'right'}); y+=20; });
  y+=14; T(9.5,'normal',SLATE);
  model.facts.concat(['Survey analysis produced by the Study Café Survey Data Analyzer.',
    model.kind.indexOf('What-if')===0 ? 'Figures reflect the scenario inputs listed in this report; they are projections, not results.'
                                      : 'What-if scenarios are not part of this report; they are downloaded separately from the What-if section.'])
    .forEach(f=>{ lines(f,CW,9.5).forEach(l=>{ T(9.5,'normal',SLATE); doc.text(l,M,y); y+=14; }); });

  const flushRow=st=>{ if(st.left){ y=Math.max(y, st.left.bottom); st.left=null; } };
  const drawChart=(b,x,w,st)=>{
    const titleL=lines(b.title||'',w-20,9.5,'bold').slice(0,2), subL=b.sub?lines(b.sub,w-20,8).slice(0,6):[];
    const maxH=b.wide?300:185; let iw=w-20, ih=iw*(b.img.h/b.img.w); if(ih>maxH){ ih=maxH; iw=ih*(b.img.w/b.img.h); }
    const gl=[]; if(b.guide){ lines(b.guide.p,w-20,7.5).forEach(t=>gl.push([t,'normal'])); b.guide.a.forEach(r=>{ lines(r[0]+': '+r[1],w-28,7.5).forEach((t,i)=>gl.push([t,i===0?'lead':'normal',r[0]])); }); }
    const gH=gl.length?(16+10+gl.length*9.2):0;
    const hh=12+titleL.length*12+subL.length*10+(subL.length?4:2)+ih+12+gH;
    return {h:hh, draw:(yy)=>{
      doc.setDrawColor(...RULE); doc.setLineWidth(0.8); doc.setFillColor(255,255,255); doc.roundedRect(x,yy,w,hh,4,4,'FD');
      let cy=yy+18; T(9.5,'bold'); titleL.forEach(l=>{ doc.text(l,x+10,cy); cy+=12; });
      T(8,'normal',SLATE); subL.forEach(l=>{ doc.text(l,x+10,cy); cy+=10; }); cy+=(subL.length?0:-2);
      try{ doc.addImage(b.img.url,'JPEG',x+10+(w-20-iw)/2,cy,iw,ih,undefined,'FAST'); }catch(e){}
      if(gl.length){
        let gy=cy+ih+14; doc.setDrawColor(...RULE); doc.setLineWidth(0.5); doc.line(x+10,gy-6,x+w-10,gy-6);
        T(8,'bold'); doc.text('About this chart',x+10,gy+4); gy+=14;
        gl.forEach(g=>{
          if(g[1]==='lead'){ const lab=_pdfSafe(g[2])+':'; T(7.5,'bold'); doc.text(lab,x+10,gy); const lw=doc.getTextWidth(lab+' ');
            T(7.5,'normal',SLATE); const rest=_pdfSafe(g[0]).slice(lab.length).trim(); doc.text(rest,x+10+lw,gy); }
          else { T(7.5,'normal',SLATE); doc.text(g[0],x+10,gy); }
          gy+=9.2; });
      }
    }};
  };
  const drawTable=b=>{
    const nC=Math.max(b.head.length, ...b.rows.map(r=>r.length));
    if(nC>9){ need(30); T(9,'italic',SLATE); lines((b.title?b.title+': ':'')+'this table has '+nC+' columns and is too wide for the page. It is included in full in the HTML report.',CW,9,'italic').forEach(l=>{ doc.text(l,M,y+10); y+=12; }); y+=10; return; }
    if(b.title){ need(40); T(9.5,'bold'); doc.text(_pdfSafe(b.title),M,y+10); y+=18; }
    // column widths from measured text: short columns get what they need, long ones share the rest
    const tw=(t,st)=>{ doc.setFont('helvetica',st); doc.setFontSize(8); return doc.getTextWidth(_pdfSafe(t)); };
    const nat=Array.from({length:nC},(_,c)=>{
      const hw=Math.max(...String(b.head[c]||'').split(/\s+/).map(wd=>tw(wd,'bold')), 0);       // longest header word
      const bw=Math.max(...b.rows.slice(0,60).map(r=>tw(r[c]||'','normal')), tw(b.head[c]||'','bold'), 0);
      return {min:Math.max(hw+12,30), want:Math.max(bw+12,hw+12,30)}; });
    let wC=nat.map(n=>n.want);
    if(wC.reduce((a,v)=>a+v,0)>CW){
      const cap=CW/nC*0.9; let fixed=0, flex=0; nat.forEach(n=>{ if(n.want<=cap) fixed+=n.want; else flex+=n.want; });
      wC=nat.map(n=>n.want<=cap?n.want:Math.max(n.min,(CW-fixed)*n.want/flex));
    }
    { const k=CW/wC.reduce((a,v)=>a+v,0); wC=wC.map(v=>v*k); }
    const xC=[]; let xx=M; for(let c=0;c<nC;c++){ xC.push(xx); xx+=wC[c]; }
    const head=()=>{ T(8,'bold',SLATE); const hl=b.head.map((h,c)=>lines(h,wC[c]-10,8,'bold')); T(8,'bold',SLATE); const hh=Math.max(1,...hl.map(l=>l.length))*10+10;
      need(hh+22); hl.forEach((l,c)=>l.forEach((t,i)=>doc.text(t,xC[c]+5,y+12+i*10))); y+=hh;
      doc.setDrawColor(197,208,200); doc.setLineWidth(0.9); doc.line(M,y,PW-M,y); };
    head();
    b.rows.forEach(r=>{
      const cl=r.map((v,c)=>lines(v,wC[c]-10,8)); T(8,'normal'); const rh=Math.max(1,...cl.map(l=>l.length))*10+9;
      if(y+rh>BOT){ newPage(); head(); T(8,'normal'); }
      cl.forEach((l,c)=>l.forEach((t,i)=>doc.text(t,xC[c]+5,y+12+i*10))); y+=rh;
      doc.setDrawColor(...RULE); doc.setLineWidth(0.5); doc.line(M,y,PW-M,y);
    });
    y+=16;
  };

  let done=0; const total=model.sections.reduce((a,s)=>a+s.blocks.length,0)||1;
  model.sections.forEach((sec,si)=>{
    newPage();
    T(8.5,'bold',GREEN); doc.text('Section '+(si+1), M, y+8); y+=18;
    T(22,'bold'); doc.text(_pdfSafe(sec.label), M, y+16); y+=28;
    doc.setFillColor(...GREEN); doc.rect(M,y,42,2.2,'F'); y+=22;
    const st={left:null};
    sec.blocks.forEach((b,bi)=>{
      done++; if(done%6===0) _progSet(70+Math.round(done/total*26), 'Writing page '+doc.getNumberOfPages()+'...');
      if(b.type==='chart' && !b.wide){
        const w=(CW-14)/2;
        if(st.left){ const c=drawChart(b, M+w+14, w, st); c.draw(st.left.top); y=Math.max(st.left.bottom, st.left.top+c.h)+14; st.left=null; return; }
        const c=drawChart(b, M, w, st); need(c.h+6); c.draw(y); st.left={top:y, bottom:y+c.h}; return;
      }
      if(st.left){ y=st.left.bottom+14; st.left=null; }
      if(b.type==='chart'||b.type==='chart3d'){ const c=drawChart(b, M, CW, st); need(c.h+6); c.draw(y); y+=c.h+14; return; }
      if(b.type==='h3'){ const nx=sec.blocks[bi+1]; let after=70;
        if(nx&&(nx.type==='chart'||nx.type==='chart3d')) after=drawChart(nx,M,(nx.wide||nx.type==='chart3d')?CW:(CW-14)/2,st).h+40; else if(nx&&nx.type==='kpis') after=110; else if(nx&&(nx.type==='card'||nx.type==='text')) after=150;
        need(after); y+=6; lines(b.text,CW,13,'bold').forEach(l=>{ T(13,'bold'); doc.text(l,M,y+12); y+=16; }); y+=6; return; }
      if(b.type==='kpis'){
        const per=Math.min(4,b.items.length)||1, w=CW/per;
        for(let i=0;i<b.items.length;i+=per){
          const row=b.items.slice(i,i+per), cells=row.map(it=>({v:lines(it.value,w-16,13.5,'bold'), l:lines(it.label+(it.sub?' ('+it.sub+')':''),w-16,8)}));
          const hh=14+Math.max(...cells.map(c=>c.v.length))*16+Math.max(...cells.map(c=>c.l.length))*10+8;
          need(hh+8); doc.setDrawColor(...RULE); doc.setLineWidth(0.8); doc.setFillColor(255,255,255); doc.rect(M,y,w*row.length,hh,'FD');
          cells.forEach((c,j)=>{ const x=M+j*w; if(j) doc.line(x,y,x,y+hh); let cy=y+20; T(13.5,'bold'); c.v.forEach(t=>{ doc.text(t,x+8,cy); cy+=16; });
            T(8,'normal',SLATE); cy-=3; c.l.forEach(t=>{ doc.text(t,x+8,cy); cy+=10; }); });
          y+=hh;
        }
        y+=16; return;
      }
      if(b.type==='lead'){ const ls=lines(b.text,CW,10); need(ls.length*14+10); T(10,'normal',SLATE); ls.forEach(l=>{ doc.text(l,M,y+10); y+=14; }); y+=10; return; }
      if(b.type==='table'){ drawTable(b); return; }
      if(b.type==='text'||b.type==='card'){
        const tl=b.title?lines(b.title,CW-24,9.5,'bold'):[], bl=b.text?lines(b.text,CW-24,8.5):[], rows=b.rows||[];
        const hh=14+tl.length*12+(b.value?20:0)+bl.length*11+rows.length*12+10;
        if(hh>BOT-M){ // very long note: flow it across pages without a box
          tl.forEach(l=>{ need(14); T(9.5,'bold'); doc.text(l,M,y+10); y+=12; }); bl.forEach(l=>{ need(13); T(8.5,'normal',SLATE); doc.text(l,M,y+10); y+=11; }); y+=12; return; }
        need(hh+6); doc.setDrawColor(...RULE); doc.setLineWidth(0.8); doc.setFillColor(255,255,255); doc.roundedRect(M,y,CW,hh,4,4,'FD');
        doc.setFillColor(143,98,18); doc.rect(M,y,2.4,hh,'F');
        let cy=y+18; T(9.5,'bold'); tl.forEach(l=>{ doc.text(l,M+12,cy); cy+=12; });
        if(b.value){ T(15,'bold'); doc.text(_pdfSafe(b.value),M+12,cy+6); cy+=20; }
        T(8.5,'normal',SLATE); bl.forEach(l=>{ doc.text(l,M+12,cy); cy+=11; });
        rows.forEach(r=>{ T(8.5,'normal',SLATE); doc.text(_pdfSafe(r[0]),M+12,cy); T(8.5,'bold'); doc.text(_pdfSafe(r[1]),PW-M-12,cy,{align:'right'}); cy+=12; });
        y+=hh+12; return;
      }
    });
    if(st.left){ y=st.left.bottom+14; st.left=null; }
  });
  // running footer
  const N=doc.getNumberOfPages();
  for(let p=2;p<=N;p++){ doc.setPage(p); doc.setDrawColor(...RULE); doc.setLineWidth(0.6); doc.line(M,PH-36,PW-M,PH-36);
    T(7.5,'normal',SLATE); doc.text('Study café feasibility  |  '+_pdfSafe(model.kind), M, PH-24); doc.text('Page '+p+' of '+N, PW-M, PH-24, {align:'right'}); }
  _progSet(98,'Saving the PDF...');
  doc.save(model.file+'.pdf');
  _progSet(100,'PDF saved'); _progLog('Saved '+model.file+'.pdf ('+N+' pages)');
}

// ── HTML writer ────────────────────────────────────────────
async function _writeHtml(model){
  _progSet(70,'Assembling the HTML report...');
  const esc=s=>String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const charts=[]; let body='', nav='';
  model.sections.forEach((sec,si)=>{
    nav+='<a href="#sec'+si+'">'+esc(sec.label)+'</a>';
    body+='<section id="sec'+si+'"><h2>'+esc(sec.label)+'</h2>';
    let grid=false; const open=()=>{ if(!grid){ body+='<div class="grid">'; grid=true; } }, close=()=>{ if(grid){ body+='</div>'; grid=false; } };
    sec.blocks.forEach(b=>{
      if(b.type==='chart'||b.type==='chart3d'){
        open(); const id='c'+charts.length;
        const tall=(b.type==='chart3d')?420:((b.kind==='radar'||b.kind==='polarArea')?340:(b.wide?340:290));
        body+='<figure class="card'+(b.wide?' wide':'')+'"><figcaption><b>'+esc(b.title)+'</b>'+(b.sub?'<span>'+esc(b.sub)+'</span>':'')+'</figcaption>'
             +'<div class="cv" style="height:'+tall+'px"><canvas id="'+id+'"'+(b.type==='chart3d'?' tabindex="0" aria-label="3D cluster plot. Drag or use the arrow keys to rotate."':'')+'></canvas>'
             +'<img class="fb" alt="'+esc(b.title)+'" src="'+b.img.url+'"/></div>'
             +(b.guide?'<div class="guide"><b>About this chart</b><p>'+esc(b.guide.p)+'</p><dl>'+b.guide.a.map(r=>'<div><dt>'+esc(r[0])+'</dt><dd>'+esc(r[1])+'</dd></div>').join('')+'</dl></div>':'')
             +'</figure>';
        charts.push(b.type==='chart3d'?{id, d3:b.data}:{id, cfg:b.cfg});
        return;
      }
      close();
      if(b.type==='h3'){ body+='<h3>'+esc(b.text)+'</h3>'; return; }
      if(b.type==='lead'){ body+='<p class="lead">'+esc(b.text)+'</p>'; return; }
      if(b.type==='kpis'){ body+='<div class="kpis">'+b.items.map(k=>'<div><b>'+esc(k.value)+'</b><span>'+esc(k.label)+(k.sub?' ('+esc(k.sub)+')':'')+'</span></div>').join('')+'</div>'; return; }
      if(b.type==='table'){ body+=(b.title?'<h4>'+esc(b.title)+'</h4>':'')+'<div class="tw"><table><thead><tr>'+b.head.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'
            +b.rows.map(r=>'<tr>'+r.map(v=>'<td>'+esc(v)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>'; return; }
      if(b.type==='text'||b.type==='card'){ body+='<div class="note">'+(b.title?'<b>'+esc(b.title)+'</b>':'')+(b.value?'<div class="val">'+esc(b.value)+'</div>':'')
            +(b.text?'<p>'+esc(b.text).replace(/\n/g,'<br>')+'</p>':'')+((b.rows||[]).length?'<dl>'+b.rows.map(r=>'<div><dt>'+esc(r[0])+'</dt><dd>'+esc(r[1])+'</dd></div>').join('')+'</dl>':'')+'</div>'; return; }
    });
    close(); body+='</section>';
  });
  const css='*{box-sizing:border-box;margin:0;padding:0}'
    +':root{--bg:#F5F7F4;--panel:#FFFFFF;--text:#17231F;--muted2:#55635D;--border:#DCE3DD;--border2:#C5D0C8;--accent:#1F5C4A;--accent3:#8F6212;'
    +"--fh:'Bricolage Grotesque','Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif;--fb:'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif}"
    +'html{scroll-behavior:smooth}body{background:var(--bg);color:var(--text);font-family:var(--fb);font-size:14px;line-height:1.5;font-variant-numeric:tabular-nums}'
    +'header{background:var(--accent);color:#fff;padding:40px 40px 32px}header h1{font-family:var(--fh);font-weight:600;font-size:34px;line-height:1.15}'
    +'header p{margin-top:6px;color:rgba(255,255,255,.82);font-size:14px}'
    +'nav{position:sticky;top:0;z-index:5;background:var(--bg);border-bottom:1px solid var(--border);padding:10px 40px;display:flex;gap:6px;flex-wrap:wrap}'
    +'nav a{color:var(--accent);text-decoration:none;font-weight:600;padding:6px 10px;border-radius:6px}nav a:hover{background:#E9EFEA}nav a:focus-visible,canvas:focus-visible{outline:2px solid var(--accent);outline-offset:2px}'
    +'main{max-width:1240px;padding:8px 40px 64px}section{padding-top:36px}'
    +'h2{font-family:var(--fh);font-weight:600;font-size:26px;margin-bottom:6px}h2::after{content:"";display:block;width:42px;height:3px;background:var(--accent);margin-top:10px}'
    +'h3{font-family:var(--fh);font-weight:600;font-size:20px;margin:28px 0 12px}h4{font-size:14px;font-weight:600;margin:16px 0 8px}'
    +'.kpis{display:flex;flex-wrap:wrap;background:var(--panel);border:1px solid var(--border);border-radius:6px;overflow:hidden;margin:12px 0 18px}'
    +'.kpis div{flex:1 1 150px;padding:16px 20px;border-right:1px solid var(--border);border-bottom:1px solid var(--border);margin:0 -1px -1px 0}'
    +'.kpis b{display:block;font-family:var(--fh);font-weight:600;font-size:26px;line-height:1.15}.kpis span{display:block;margin-top:4px;font-size:13px;color:var(--muted2)}'
    +'.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px;margin:12px 0 18px}'
    +'.card{background:var(--panel);border:1px solid var(--border);border-radius:6px;padding:20px;min-width:0}.card.wide{grid-column:1/-1}'
    +'figcaption b{display:block;font-size:14px;font-weight:600}figcaption span{display:block;font-size:13px;color:var(--muted2);margin-bottom:12px}'
    +'.cv{position:relative;width:100%}.cv canvas{display:block;width:100%;height:100%}.cv .fb{display:none;width:100%;height:100%;object-fit:contain}'
    +'.nolive .cv canvas{display:none}.nolive .cv .fb{display:block}'
    +'.tw{overflow-x:auto;background:var(--panel);border:1px solid var(--border);border-radius:6px;margin:8px 0 18px}table{border-collapse:collapse;width:100%;font-size:14px}'
    +'th{text-align:left;font-size:13px;font-weight:600;color:var(--muted2);padding:10px 14px;border-bottom:1px solid var(--border2);white-space:nowrap}'
    +'td{padding:10px 14px;border-bottom:1px solid var(--border);vertical-align:top}tr:last-child td{border-bottom:0}'
    +'.lead{color:var(--muted2);font-size:16px;line-height:1.6;max-width:72ch;margin:-4px 0 18px}'
    +'.note{background:var(--panel);border:1px solid var(--border);border-left:3px solid var(--accent3);border-radius:6px;padding:16px 20px;margin:8px 0 14px;max-width:920px}'
    +'.note b{display:block;font-size:14px;font-weight:600;margin-bottom:4px}.note p{color:var(--muted2)}.note .val{font-family:var(--fh);font-weight:600;font-size:26px;margin:2px 0 6px}'
    +'.note dl{margin-top:8px}.note dl div{display:flex;justify-content:space-between;gap:16px;padding:5px 0;border-top:1px solid var(--border);font-size:13px}.note dt{color:var(--muted2)}.note dd{font-weight:600}'
    +'.guide{margin-top:14px;padding-top:12px;border-top:1px solid var(--border);font-size:13px;color:var(--muted2);line-height:1.5}.guide b{display:block;color:var(--text);font-weight:600;margin-bottom:4px}'
    +'.guide p{margin-bottom:8px}.guide dl div{display:grid;grid-template-columns:128px 1fr;gap:10px;margin-top:4px}.guide dt{font-weight:600;color:var(--text)}.guide dd{margin:0}'
    +'@media (max-width:640px){.guide dl div{grid-template-columns:1fr;gap:0}}'
    +'footer{padding:24px 40px 40px;color:var(--muted2);font-size:13px;border-top:1px solid var(--border)}'
    +'@media (max-width:860px){header,nav,main,footer{padding-left:16px;padding-right:16px}.grid{grid-template-columns:1fr}.card.wide{grid-column:auto}}'
    +'@media print{nav{display:none}.card,.note,.tw,.kpis{break-inside:avoid}body{background:#fff}}';
  const boot='(function(){var D=JSON.parse(document.getElementById("report-data").textContent);'
    +'if(typeof Chart==="undefined"){document.body.className+=" nolive";}else{'
    +'Chart.defaults.font.family="\'Public Sans\',system-ui,-apple-system,\'Segoe UI\',sans-serif";Chart.defaults.font.size=12;Chart.defaults.color="#55635D";Chart.defaults.borderColor="rgba(23,35,31,.08)";'
    +'Chart.defaults.plugins.tooltip.backgroundColor="#17231F";Chart.defaults.plugins.tooltip.padding=10;Chart.defaults.plugins.tooltip.cornerRadius=6;Chart.defaults.plugins.legend.labels.boxWidth=10;Chart.defaults.plugins.legend.labels.boxHeight=10;}'
    +'D.forEach(function(c){var cv=document.getElementById(c.id);if(!cv)return;'
    +'if(c.d3){try{new Scatter3D(cv,{points:c.d3.points,names:c.d3.names,axes:c.d3.axes});cv.parentNode.querySelector(".fb").style.display="none";cv.style.display="block";}catch(e){}return;}'
    +'if(typeof Chart==="undefined"||!c.cfg)return;try{var g=c.cfg,ds=g.data.datasets[0];'
    +'if(ds&&ds.__cells){var n=ds.__cells;ds.pointRadius=function(x){var k=x.chart;return Math.min(k.width/(n[0]*2.2),k.height/(n[1]*2.2));};ds.pointHoverRadius=ds.pointRadius;'
    +'var L=g.__axisLabels||{x:[],y:[]};g.options.scales.x.ticks=g.options.scales.x.ticks||{};g.options.scales.y.ticks=g.options.scales.y.ticks||{};'
    +'g.options.scales.x.ticks.callback=function(v){return L.x[Math.round(v)]||"";};g.options.scales.y.ticks.callback=function(v){return L.y[Math.round(v)]||"";};'
    +'g.options.plugins=g.options.plugins||{};g.options.plugins.tooltip={callbacks:{label:function(t){var r=t.raw||{};return (L.y[r.y]||"")+" / "+(L.x[r.x]||"")+": "+(r.v==null?"":r.v);}}};}'
    +'Object.keys(g.options.scales||{}).forEach(function(k){var F=g.options.scales[k].__fmt;if(!F)return;g.options.scales[k].ticks=g.options.scales[k].ticks||{};g.options.scales[k].ticks.callback=function(v){return F.pre+(Math.round(v*F.mul*100)/100).toLocaleString()+F.suf;};});'
    +'new Chart(cv,g);}catch(e){cv.style.display="none";cv.parentNode.querySelector(".fb").style.display="block";}});})();';
  const html='<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>'
    +'<title>Study café feasibility — '+esc(model.kind)+'</title>'
    +'<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600&family=Public+Sans:wght@400;600&display=swap" rel="stylesheet">'
    +'<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"><\/script>'
    +'<style>'+css+'</style></head><body>'
    +'<header><h1>Study café feasibility</h1><p>'+esc(model.kind)+' · '+esc(model.date.toLocaleDateString('en-GB',{year:'numeric',month:'long',day:'numeric'}))
    +(model.facts.length?' · '+esc(model.facts[0]):'')+'</p></header>'
    +'<nav aria-label="Sections">'+nav+'</nav><main>'+body+'</main>'
    +'<footer>'+(model.kind.indexOf('What-if')===0?'Figures reflect the scenario inputs listed above; they are projections, not results.'
                 :'What-if scenarios are not part of this report; they are downloaded separately from the What-if section of the dashboard.')
    +' Charts are interactive when this file is opened with an internet connection; otherwise still images are shown.</footer>'
    +'<script type="application/json" id="report-data">'+JSON.stringify(charts).replace(/</g,'\\u003c')+'<\/script>'
    +'<script>'+Scatter3D.toString()+'<\/script>'
    +'<script>'+boot+'<\/script></body></html>';
  _progSet(94,'Saving the HTML report...');
  const blob=new Blob([html],{type:'text/html;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=model.file+'.html';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(a.href),10000);
  _progSet(100,'HTML report saved'); _progLog('Saved '+model.file+'.html ('+charts.length+' charts)');
}



// ════════════════════════════════════════════════════════════
//  MEMBERSHIP PREDICTOR  — SCREEN (s-predictor)
// ════════════════════════════════════════════════════════════
let _predHistory = [];

// ── Helpers ──────────────────────────────────────────────────
function _setFieldError(fieldId, errId, msg){
  const field = document.getElementById(fieldId);
  const span  = document.getElementById(errId);
  if(field){ field.classList.add('has-error'); }
  if(span) { span.textContent = msg; }
}
function clearFieldError(fieldId){
  const field = document.getElementById(fieldId);
  if(field){ field.classList.remove('has-error'); }
  // also re-check logical mismatch live whenever time or late-night changes
  _clearBanner();
}
function _clearAllErrors(){
  ['field-vf','field-spend','field-sh','field-gs',
   'field-age','field-dist','field-time','field-ln'].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.classList.remove('has-error');
  });
  ['err-vf','err-spend','err-sh','err-gs',
   'err-age','err-dist','err-time','err-ln'].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.textContent = '';
  });
  _clearBanner();
}
function _clearBanner(){
  const b = document.getElementById('pred-error-banner');
  if(b) b.style.display = 'none';
}
function _showBanner(messages){
  const b = document.getElementById('pred-error-banner');
  const l = document.getElementById('pred-error-list');
  if(!b||!l) return;
  l.innerHTML = messages.map(m=>`<li>${m}</li>`).join('');
  b.style.display = 'block';
  b.scrollIntoView({behavior:'smooth', block:'nearest'});
}

// ── Live mismatch check triggered whenever time or late-night changes ──
function onTimeOrLateNightChange(){
  const time = document.getElementById('inp-time').value;
  const ln   = document.getElementById('inp-ln').value;
  // Clear old errors on these two fields first
  document.getElementById('field-time').classList.remove('has-error');
  document.getElementById('field-ln').classList.remove('has-error');
  document.getElementById('err-time').textContent = '';
  document.getElementById('err-ln').textContent   = '';
  _clearBanner();
  // Check mismatch only when both are filled
  if(time && ln){
    if(time === 'Late Night' && ln === 'No'){
      _setFieldError('field-time','err-time','Preferred time is Late Night but Late Night Preference is No');
      _setFieldError('field-ln',  'err-ln',  'Set to No but Preferred Study Time is Late Night');
      _showBanner(['Logical mismatch: Preferred Study Time is "Late Night" but Late Night Study Preference is "No". These contradict each other — please correct one.']);
    }
    // v41: liking late-night study while preferring another slot is a valid answer (many respondents gave it)
  }
}

// ── v41: the form's choices are the categories present in the uploaded data ──
async function openPredictor(){
  goTo('s-predictor');
  try{
    const r=await fetch(BASE+'/predictor_options'); const o=await r.json();
    if(o.error) return;
    window._PRED_OPTS=o;
    const M={'inp-vf':'visit_freq','inp-spend':'spend','inp-sh':'study_hours','inp-gs':'group_study','inp-dist':'distance','inp-time':'pref_time','inp-ln':'late_night'};
    Object.entries(M).forEach(([id,k])=>{
      const sel=document.getElementById(id); const opts=o.options[k];
      if(!sel||!opts) return;
      const cur=sel.value; const txt={};
      Array.from(sel.options).forEach(op=>{txt[op.value]=op.textContent;});
      sel.innerHTML='';
      const ph=document.createElement('option'); ph.value=''; ph.textContent='— Select —'; ph.disabled=true; ph.selected=true; sel.appendChild(ph);
      opts.forEach(v=>{const op=document.createElement('option'); op.value=v; op.textContent=txt[v]||v; sel.appendChild(op);});
      if(opts.includes(cur)) sel.value=cur;
    });
    const a=document.getElementById('inp-age');
    if(a){a.min=o.age_min; a.max=o.age_max; a.placeholder='e.g. '+Math.round((o.age_min+o.age_max)/2)+' ('+o.age_min+'–'+o.age_max+')';}
  }catch(e){}
}

// ── Full validation run on submit ────────────────────────────
function _validateForm(){
  _clearAllErrors();
  const errors = [];

  // Required dropdown fields
  const required = [
    {id:'inp-vf',    field:'field-vf',    err:'err-vf',    label:'Visit Frequency / Week'},
    {id:'inp-spend', field:'field-spend', err:'err-spend', label:'Avg Spend per Visit'},
    {id:'inp-sh',    field:'field-sh',    err:'err-sh',    label:'Study Hours per Session'},
    {id:'inp-gs',    field:'field-gs',    err:'err-gs',    label:'Group Study Frequency'},
    {id:'inp-dist',  field:'field-dist',  err:'err-dist',  label:'Distance from Campus'},
    {id:'inp-time',  field:'field-time',  err:'err-time',  label:'Preferred Study Time'},
    {id:'inp-ln',    field:'field-ln',    err:'err-ln',    label:'Late Night Study Preference'},
  ];
  required.forEach(({id,field,err,label})=>{
    const val = document.getElementById(id).value;
    if(!val){
      _setFieldError(field, err, 'This field is required');
      errors.push(label + ' is required');
    }
  });

  // Age validation
  const ageVal = document.getElementById('inp-age').value;
  if(!ageVal || ageVal === ''){
    _setFieldError('field-age','err-age','Age is required');
    errors.push('Age is required');
  } else {
    const ageNum = parseInt(ageVal);
    const _aMin=(window._PRED_OPTS&&window._PRED_OPTS.age_min)||17, _aMax=(window._PRED_OPTS&&window._PRED_OPTS.age_max)||35;
    if(isNaN(ageNum) || ageNum < _aMin || ageNum > _aMax){
      _setFieldError('field-age','err-age','Age must be between '+_aMin+' and '+_aMax+' (the ages in your data)');
      errors.push('Age must be between '+_aMin+' and '+_aMax+' (the ages in your data)');
    }
  }

  // Logical mismatch check (only if both filled)
  const time = document.getElementById('inp-time').value;
  const ln   = document.getElementById('inp-ln').value;
  if(time && ln){
    if(time === 'Late Night' && ln === 'No'){
      _setFieldError('field-time','err-time','Late Night selected but Late Night Preference is No');
      _setFieldError('field-ln',  'err-ln',  'No selected but Preferred Time is Late Night');
      errors.push('Logical mismatch: Preferred Study Time is "Late Night" but Late Night Study Preference is "No"');
    }
  }

  if(errors.length > 0){
    _showBanner(errors);
    return false;
  }
  return true;
}

// ── Reset ─────────────────────────────────────────────────────
function resetPredictor(){
  // Reset all dropdowns to blank
  ['inp-vf','inp-spend','inp-sh','inp-gs','inp-dist','inp-time','inp-ln'].forEach(id=>{
    document.getElementById(id).value = '';
  });
  document.getElementById('inp-age').value = '';
  // Reset sliders
  document.getElementById('inp-power').value = 3; document.getElementById('pow-val').textContent  = 3;
  document.getElementById('inp-wifi').value  = 3; document.getElementById('wifi-val').textContent = 3;
  document.getElementById('inp-prod').value  = 3; document.getElementById('prod-val').textContent = 3;
  // Clear all errors
  _clearAllErrors();
  // Hide result, show submit, hide retry
  document.getElementById('pred-result').style.display = 'none';
  document.getElementById('pred-submit-btn').style.display = 'block';
  document.getElementById('pred-retry-btn').style.display  = 'none';
  // Scroll form back into view
  document.querySelector('#s-predictor .predictor-body').scrollTop = 0;
  document.getElementById('s-predictor').scrollTop = 0;
}

// ── Main predict function ─────────────────────────────────────
function runPrediction(){
  // Run full validation first — stop if anything fails
  if(!_validateForm()) return;

  const submitBtn = document.getElementById('pred-submit-btn');
  const retryBtn  = document.getElementById('pred-retry-btn');
  const resCard   = document.getElementById('pred-result');

  submitBtn.disabled    = true;
  submitBtn.textContent = 'Running model…';

  const body = {
    age:                document.getElementById('inp-age').value,
    visit_freq:         document.getElementById('inp-vf').value,
    spend:              document.getElementById('inp-spend').value,
    study_hours:        document.getElementById('inp-sh').value,
    distance:           document.getElementById('inp-dist').value,
    group_study:        document.getElementById('inp-gs').value,
    power_outlet:       document.getElementById('inp-power').value,
    wifi:               document.getElementById('inp-wifi').value,
    productivity_tools: document.getElementById('inp-prod').value,
    late_night:         document.getElementById('inp-ln').value,
    pref_time:          document.getElementById('inp-time').value,
  };

  fetch(BASE + '/predict_membership', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)
  })
  .then(r => r.json())
  .then(d => {
    submitBtn.disabled    = false;
    submitBtn.textContent = 'Predict membership';
    if(d.error){ alert('Prediction error: ' + d.error); return; }

    const isYes    = d.verdict === 'Yes';
    const mainCol  = isYes ? 'var(--accent2)' : 'var(--danger)';
    const confCol  = d.confidence === 'High' ? 'var(--accent2)'
                   : d.confidence === 'Moderate' ? 'var(--warn)' : 'var(--muted2)';
    const pct      = d.probability_pct;

    // ── Build result card ─────────────────────────────────────────
    resCard.innerHTML = `
      <div class="pred-verdict">
        <div class="pred-verdict-text" style="color:${mainCol}">
          ${isYes ? '✅ Likely Member' : '❌ Unlikely Member'}
        </div>
        <div class="pred-verdict-sub">
          Confidence: <span style="color:${confCol};font-weight:700">${d.confidence}</span>
          &nbsp;·&nbsp; Based on Random Forest trained on your cleaned dataset
        </div>
      </div>

      <div class="pred-gauge-wrap">
        <canvas id="gauge-canvas" width="200" height="200"></canvas>
        <div class="pred-gauge-label">
          <div class="pred-pct" style="color:${mainCol}">${pct}%</div>
          <div class="pred-pct-sub">membership probability</div>
        </div>
      </div>
      <p class="chart-guide" style="border-top:0;margin-top:8px;padding-top:0">The ring fills in proportion to the estimated chance that a student with these answers would take a membership: a full ring is 100%, half a ring is 50%.</p>

      <div class="pred-bar-wrap">
        <div class="pred-bar-labels"><span>0% — Will Not Join</span><span>Will Join — 100%</span></div>
        <div class="pred-bar-track">
          <div class="pred-bar-fill" style="width:${pct}%;background:#1F5C4A"></div>
          <div class="pred-bar-mid"></div>
        </div>
        <div class="pred-bar-caption">Decision threshold at 50%</div>
      </div>

      <div class="pred-factors-title">Top factors driving this prediction</div>
      ${d.top_factors.map((f,i)=>`
        <div class="pred-factor-row">
          <span class="pred-factor-rank">#${i+1}</span>
          <span class="pred-factor-name">${f.name}</span>
          <div class="pred-factor-bar-track">
            <div class="pred-factor-bar-fill" style="width:${Math.min(f.importance*4,100)}%;background:${f.effect_pts<0?'var(--danger)':'var(--accent2)'}"></div>
          </div>
          <span class="pred-factor-pct" style="color:${f.effect_pts<0?'var(--danger)':'var(--accent2)'}">${f.effect_pts>0?'+':''}${f.effect_pts} pts</span>
        </div>`).join('')}
      <div style="font-size:.72rem;color:var(--muted2);margin-top:8px;line-height:1.5">${d.factor_note||''}${(d.warnings||[]).map(w=>'<br>⚠ '+w).join('')}</div>
    `;
    resCard.style.display = 'block';

    // ── Draw gauge arc ───────────────────────────────────────────
    requestAnimationFrame(()=>{
      const cv = document.getElementById('gauge-canvas');
      if(!cv) return;
      const ctx = cv.getContext('2d');
      const cx=100, cy=100, r=82, lw=15;
      ctx.clearRect(0,0,200,200);
      ctx.beginPath(); ctx.arc(cx,cy,r, Math.PI*0.75, Math.PI*2.25);
      ctx.strokeStyle='rgba(23,35,31,.07)'; ctx.lineWidth=lw; ctx.lineCap='round'; ctx.stroke();
      const end = Math.PI*0.75 + (pct/100)*Math.PI*1.5;
      const grad = ctx.createLinearGradient(18,100,182,100);
      grad.addColorStop(0,'#B3261E'); grad.addColorStop(1, isYes?'#2f8f6a':'#B3261E');
      ctx.beginPath(); ctx.arc(cx,cy,r, Math.PI*0.75, end);
      ctx.strokeStyle=grad; ctx.lineWidth=lw; ctx.lineCap='round'; ctx.stroke();
    });

    // ── Switch buttons ───────────────────────────────────────────
    submitBtn.style.display = 'none';
    retryBtn.style.display  = 'block';

    // ── Add to history ────────────────────────────────────────────
    _predHistory.push({
      n:       _predHistory.length + 1,
      vf:      body.visit_freq,
      spend:   body.spend,
      sh:      body.study_hours,
      dist:    body.distance,
      pct:     pct,
      verdict: d.verdict,
      conf:    d.confidence,
    });
    _renderPredHistory();

    resCard.scrollIntoView({behavior:'smooth', block:'start'});
  })
  .catch(e => {
    submitBtn.disabled    = false;
    submitBtn.textContent = 'Predict membership';
    alert('Failed: ' + e.message);
  });
}

function _renderPredHistory(){
  if(_predHistory.length === 0) return;
  const histWrap = document.getElementById('pred-history');
  const tbody    = document.getElementById('pred-history-tbody');
  histWrap.style.display = 'block';
  tbody.innerHTML = _predHistory.slice().reverse().map(r=>`
    <tr>
      <td style="color:var(--muted)">${r.n}</td>
      <td>${r.vf}</td>
      <td>₹${r.spend}</td>
      <td>${r.sh} hrs</td>
      <td>${r.dist}</td>
      <td style="font-weight:700;color:${r.pct>=50?'var(--accent2)':'var(--danger)'}">${r.pct}%</td>
      <td class="${r.verdict==='Yes'?'verdict-yes':'verdict-no'}">${r.verdict==='Yes'?'✅ Yes':'❌ No'}</td>
      <td style="color:${r.conf==='High'?'var(--accent2)':r.conf==='Moderate'?'var(--warn)':'var(--muted2)'}">${r.conf}</td>
    </tr>`).join('');
}
</script>

<!-- Export format modal -->
<div id="export-modal-backdrop">
  <div id="export-modal">
    <h3>Download the analysis report</h3>
    <p>Covers Data summary, Exploratory analysis and Predictive analytics.</p>
    <fieldset class="exp-level" id="exp-level">
      <legend>Level of detail</legend>
      <label class="exp-choice"><input type="radio" name="exp-level" value="surface">
        <span><b>Surface report</b><span class="exp-desc">What matters for running the café, in plain words.</span></span></label>
      <label class="exp-choice"><input type="radio" name="exp-level" value="deep">
        <span><b>Deep report</b><span class="exp-desc">The full analysis, with every technique explained in plain words.</span></span></label>
    </fieldset>

    <div class="export-option" onclick="_runExport('pdf')" role="button" tabindex="0" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();_runExport('pdf')}">
      <div class="eo-icon"></div>
      <div class="eo-body">
        <div class="eo-title">PDF report <span class="eo-tag tag-static">PDF</span></div>
        <div class="eo-desc">A paginated A4 document with every chart, key figure and table. Best for printing or submitting.</div>
      </div>
    </div>

    <div class="export-option" onclick="_runExport('html')" role="button" tabindex="0" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();_runExport('html')}">
      <div class="eo-icon"></div>
      <div class="eo-body">
        <div class="eo-title">Interactive HTML report <span class="eo-tag tag-live">Live</span></div>
        <div class="eo-desc">One <code>.html</code> file that opens in any browser. Hover a chart for exact values and drag the 3D cluster plot to rotate it.</div>
      </div>
    </div>

    <button id="export-modal-cancel" onclick="_closeExportModal()">Cancel</button>
  </div>
</div>

<!-- Export progress dialog -->
<div id="export-prog-backdrop">
  <div id="export-prog-box">
    <h3 id="export-prog-title">Preparing Export…</h3>
    <div id="export-prog-track"><div id="export-prog-fill"></div></div>
    <div id="export-prog-pct">0%</div>
    <div id="export-prog-step">Initialising…</div>
    <div id="export-prog-log"></div>
  </div>
</div>
</body>
</html>
"""

# ── Inject predictor CSS into the existing <style> block ─────────
_PREDICTOR_CSS = """
/* ── Membership Predictor Screen ── */
#s-predictor{min-height:100vh;padding:0}
.predictor-body{width:100%;max-width:960px;padding:32px 20px;margin:0 auto}
.pred-intro{background:var(--card);border:1px solid var(--border2);border-radius:var(--r);padding:20px 22px;margin-bottom:28px;font-size:.86rem;color:var(--muted2);line-height:1.65}
.pred-section-label{font-size:.7rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin:0 0 12px;display:flex;align-items:center;gap:8px}
.pred-section-label::after{content:'';flex:1;height:1px;background:var(--border)}
.pred-form-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:16px;margin-bottom:24px}
.pred-field{display:flex;flex-direction:column;gap:6px}
.pred-field label{font-size:.75rem;font-weight:700;color:var(--muted2);text-transform:uppercase;letter-spacing:.07em;display:flex;align-items:center;gap:6px}
.pred-field label .impact-tag{font-size:.62rem;font-weight:700;padding:2px 6px;border-radius:4px;letter-spacing:.04em;text-transform:uppercase}
.impact-high{background:rgba(127,154,36,.18);color:var(--danger)}
.impact-med{background:rgba(210,105,44,.18);color:#d2692c}
.pred-field select,.pred-field input[type=number]{background:var(--card);border:1px solid var(--border2);border-radius:8px;padding:10px 12px;color:var(--text);font-size:.88rem;outline:none;transition:border-color .15s;width:100%;cursor:pointer}
.pred-field select:focus,.pred-field input:focus{border-color:var(--accent)}
.pred-field select option{background:var(--card)}
.likert-row{display:flex;gap:8px;align-items:center}
.likert-row span{font-size:.72rem;color:var(--muted2);white-space:nowrap}
.likert-row input[type=range]{flex:1;accent-color:var(--accent)}
.likert-val{font-weight:700;color:var(--accent);min-width:16px;text-align:center}
.pred-btn-row{display:flex;gap:12px;margin-bottom:4px}
.pred-submit{flex:1;padding:15px;border-radius:12px;border:none;background:#1F5C4A;color:#F5F7F4;font-size:1rem;font-weight:800;cursor:pointer;letter-spacing:.03em;transition:opacity .15s,transform .1s;box-shadow:0 4px 24px rgba(43,110,163,.35)}
.pred-submit:hover{opacity:.9;transform:translateY(-1px)}
.pred-submit:disabled{opacity:.5;cursor:default;transform:none}
.pred-retry{flex:0 0 auto;padding:15px 28px;border-radius:12px;border:1px solid var(--border2);background:var(--card);color:var(--text);font-size:.95rem;font-weight:700;cursor:pointer;transition:border-color .15s,background .15s;display:none}
.pred-retry:hover{border-color:var(--accent);background:rgba(43,110,163,.1)}
.pred-result-card{background:var(--card);border:1px solid var(--border2);border-radius:var(--r);padding:28px;margin-top:20px;display:none}
.pred-verdict{text-align:center;margin-bottom:28px}
.pred-verdict-text{font-size:2.8rem;font-weight:900;line-height:1}
.pred-verdict-sub{font-size:.95rem;color:var(--muted2);margin-top:8px}
.pred-gauge-wrap{position:relative;width:200px;height:200px;margin:0 auto 28px}
.pred-gauge-label{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;pointer-events:none}
.pred-pct{font-size:2.2rem;font-weight:900}
.pred-pct-sub{font-size:.68rem;color:var(--muted2)}
.pred-bar-wrap{margin-bottom:24px}
.pred-bar-labels{display:flex;justify-content:space-between;font-size:.72rem;color:var(--muted2);margin-bottom:6px}
.pred-bar-track{height:14px;background:var(--border);border-radius:7px;overflow:hidden;position:relative}
.pred-bar-fill{height:100%;border-radius:7px;transition:width .8s cubic-bezier(.4,0,.2,1)}
.pred-bar-mid{position:absolute;top:0;left:50%;width:2px;height:100%;background:rgba(23,35,31,.25)}
.pred-bar-caption{text-align:center;font-size:.7rem;color:var(--muted2);margin-top:5px}
.pred-factors-title{font-size:.7rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:12px}
.pred-factor-row{display:flex;align-items:center;gap:10px;margin-bottom:10px;font-size:.83rem}
.pred-factor-rank{flex:0 0 22px;font-size:.68rem;font-weight:800;color:var(--muted2);text-align:center}
.pred-factor-name{flex:1;color:var(--text)}
.pred-factor-bar-track{flex:0 0 90px;height:6px;background:var(--border);border-radius:3px;overflow:hidden}
.pred-factor-bar-fill{height:100%;border-radius:3px;background:var(--accent)}
.pred-factor-pct{flex:0 0 76px;text-align:right;color:var(--accent2);font-weight:700}
.pred-history{margin-top:28px}
.pred-history-title{font-size:.7rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:10px}
.pred-history-table{width:100%;border-collapse:collapse;font-size:.78rem}
.pred-history-table thead th{background:rgba(43,110,163,.1);color:var(--accent);padding:8px 10px;text-align:left;font-weight:600}
.pred-history-table tbody tr:nth-child(even){background:rgba(23,35,31,.02)}
.pred-history-table tbody td{padding:7px 10px;border-bottom:1px solid var(--border);color:var(--muted2)}
.pred-history-table .verdict-yes{color:var(--accent2);font-weight:700}
.pred-history-table .verdict-no{color:var(--danger);font-weight:700}
/* ── Validation errors ── */
.pred-field.has-error select,
.pred-field.has-error input[type=number]{border-color:var(--danger)!important;background:rgba(127,154,36,.07)!important}
.pred-field.has-error label{color:var(--danger)!important}
.pred-field-error{font-size:.72rem;color:var(--danger);margin-top:3px;font-weight:600;display:none}
.pred-field.has-error .pred-field-error{display:block}
.pred-error-banner{background:rgba(127,154,36,.1);border:1px solid rgba(127,154,36,.4);border-radius:8px;padding:12px 16px;margin-bottom:16px;font-size:.83rem;color:var(--danger);font-weight:600;display:none}
.pred-error-banner ul{margin:6px 0 0 16px;font-weight:400;color:var(--text)}
.pred-error-banner li{margin-bottom:3px}
"""
HTML_PAGE = HTML_PAGE.replace('</style>\n</head>', _PREDICTOR_CSS + '</style>\n</head>')

# ── Inject predictor screen HTML ─────────────────────────────────
_PREDICTOR_HTML = """
<!-- S-PREDICTOR: MEMBERSHIP WILLINGNESS PREDICTOR -->
<div id="s-predictor" class="screen">
  <div class="topbar">
    <button class="back-btn" onclick="goTo('s-nav')">← Back</button>
    <h2>Membership predictor</h2>
  </div>
  <div class="predictor-body">

    <div class="pred-intro">
      Fill in a student's profile using the fields below. Variables marked
      <span class="impact-tag impact-high">High Impact</span> are the strongest predictors
      of membership willingness based on your dataset's association rules and feature importances.
      Click <strong style="color:var(--accent2)">Predict</strong> to get an instant prediction
      from the trained Random Forest model.
    </div>

    <!-- Validation error banner -->
    <div class="pred-error-banner" id="pred-error-banner">
      <span>⚠Please fix the following before predicting:</span>
      <ul id="pred-error-list"></ul>
    </div>

    <!-- Section 1: Behavioural -->
    <div class="pred-section-label">Study habits and spending</div>
    <div class="pred-form-grid">

      <div class="pred-field" id="field-vf">
        <label>Visit Frequency / Week <span class="impact-tag impact-high">High Impact</span></label>
        <select id="inp-vf" onchange="clearFieldError('field-vf')">
          <option value="" selected disabled>— Select —</option>
          <option value="Rarely">Rarely</option>
          <option value="1">1× per week</option>
          <option value="2-3">2–3× per week</option>
          <option value="4-5">4–5× per week</option>
          <option value="Daily">Daily</option>
        </select>
        <span class="pred-field-error" id="err-vf"></span>
      </div>

      <div class="pred-field" id="field-spend">
        <label>Avg Spend per Visit <span class="impact-tag impact-high">High Impact</span></label>
        <select id="inp-spend" onchange="clearFieldError('field-spend')">
          <option value="" selected disabled>— Select —</option>
          <option value="100-200">₹100–200</option>
          <option value="200-400">₹200–400</option>
          <option value="400-600">₹400–600</option>
          <option value=">600">₹600+</option>
        </select>
        <span class="pred-field-error" id="err-spend"></span>
      </div>

      <div class="pred-field" id="field-sh">
        <label>Study Hours per Session <span class="impact-tag impact-high">High Impact</span></label>
        <select id="inp-sh" onchange="clearFieldError('field-sh')">
          <option value="" selected disabled>— Select —</option>
          <option value="<1">&lt; 1 hr</option>
          <option value="1-2">1–2 hrs</option>
          <option value="2-3">2–3 hrs</option>
          <option value="3-4">3–4 hrs</option>
          <option value=">4">&gt; 4 hrs</option>
        </select>
        <span class="pred-field-error" id="err-sh"></span>
      </div>

      <div class="pred-field" id="field-gs">
        <label>Group Study Frequency <span class="impact-tag impact-med">Med Impact</span></label>
        <select id="inp-gs" onchange="clearFieldError('field-gs')">
          <option value="" selected disabled>— Select —</option>
          <option value="Never">Never</option>
          <option value="Rarely">Rarely</option>
          <option value="Sometimes">Sometimes</option>
          <option value="Occasionally">Occasionally</option>
          <option value="Frequently">Frequently</option>
          <option value="Often">Often</option>
          <option value="Always">Always</option>
        </select>
        <span class="pred-field-error" id="err-gs"></span>
      </div>

    </div>

    <!-- Section 2: Demographic & Location -->
    <div class="pred-section-label">About the student</div>
    <div class="pred-form-grid">

      <div class="pred-field" id="field-age">
        <label>Age</label>
        <input type="number" id="inp-age" min="17" max="35" placeholder="e.g. 21"
               oninput="clearFieldError('field-age')"/>
        <span class="pred-field-error" id="err-age"></span>
      </div>

      <div class="pred-field" id="field-dist">
        <label>Distance from Campus <span class="impact-tag impact-med">Med Impact</span></label>
        <select id="inp-dist" onchange="clearFieldError('field-dist')">
          <option value="" selected disabled>— Select —</option>
          <option value="<1 km">&lt; 1 km</option>
          <option value="1-2 km">1–2 km</option>
          <option value="2-3 km">2–3 km</option>
          <option value=">3 km">&gt; 3 km</option>
        </select>
        <span class="pred-field-error" id="err-dist"></span>
      </div>

      <div class="pred-field" id="field-time">
        <label>Preferred Study Time</label>
        <select id="inp-time" onchange="onTimeOrLateNightChange()">
          <option value="" selected disabled>— Select —</option>
          <option value="Morning">Morning</option>
          <option value="Afternoon">Afternoon</option>
          <option value="Evening">Evening</option>
          <option value="Late Night">Late Night</option>
        </select>
        <span class="pred-field-error" id="err-time"></span>
      </div>

      <div class="pred-field" id="field-ln">
        <label>Late Night Study Preference</label>
        <select id="inp-ln" onchange="onTimeOrLateNightChange()">
          <option value="" selected disabled>— Select —</option>
          <option value="No">No</option>
          <option value="Yes">Yes</option>
        </select>
        <span class="pred-field-error" id="err-ln"></span>
      </div>

    </div>

    <!-- Section 3: Facility Preferences (Likert) -->
    <div class="pred-section-label">What matters in the café (1 = not important, 5 = very important)</div>
    <div class="pred-form-grid">

      <div class="pred-field" id="field-power">
        <label>Power outlets</label>
        <div class="likert-row">
          <span>1</span>
          <input type="range" id="inp-power" min="1" max="5" value="3"
                 oninput="document.getElementById('pow-val').textContent=this.value"/>
          <span>5</span>
          <span class="likert-val" id="pow-val">3</span>
        </div>
      </div>

      <div class="pred-field" id="field-wifi">
        <label>Wi-Fi</label>
        <div class="likert-row">
          <span>1</span>
          <input type="range" id="inp-wifi" min="1" max="5" value="3"
                 oninput="document.getElementById('wifi-val').textContent=this.value"/>
          <span>5</span>
          <span class="likert-val" id="wifi-val">3</span>
        </div>
      </div>

      <div class="pred-field" id="field-prod">
        <label>Productivity tools</label>
        <div class="likert-row">
          <span>1</span>
          <input type="range" id="inp-prod" min="1" max="5" value="3"
                 oninput="document.getElementById('prod-val').textContent=this.value"/>
          <span>5</span>
          <span class="likert-val" id="prod-val">3</span>
        </div>
      </div>

    </div>

    <!-- Buttons -->
    <div class="pred-btn-row">
      <button class="pred-submit" id="pred-submit-btn" onclick="runPrediction()">Predict membership</button>
      <button class="pred-retry" id="pred-retry-btn" onclick="resetPredictor()">Retry</button>
    </div>

    <!-- Result card -->
    <div class="pred-result-card" id="pred-result"></div>

    <!-- Prediction history -->
    <div class="pred-history" id="pred-history" style="display:none">
      <div class="pred-history-title">Prediction History (this session)</div>
      <div class="tbl-wrap">
        <table class="pred-history-table">
          <thead><tr>
            <th>#</th><th>Visit Freq</th><th>Spend</th><th>Study Hrs</th>
            <th>Distance</th><th>Probability</th><th>Verdict</th><th>Confidence</th>
          </tr></thead>
          <tbody id="pred-history-tbody"></tbody>
        </table>
      </div>
    </div>

  </div>
</div>
"""
_WHATIF_HTML = '''
<!-- S-WHATIF: WHAT-IF SCENARIO MODELLING -->
<div id="s-whatif" class="screen">
  <div class="topbar">
    <button class="back-btn" onclick="goTo(\'s-nav\')">← Back</button>
    <h2>What-if scenarios</h2>
  </div>
  <div class="wi-body">

    <!-- Parameters panel -->
    <div class="wi-controls">
      <h3>Scenario settings</h3>
      <div style="font-size:.75rem;color:var(--muted2);margin-bottom:14px">
        Change any setting, then select <strong>Run scenario</strong> to update all six results at once.
      </div>
      <fieldset class="wi-group"><legend>Market</legend><div class="wi-grid">
        <div class="wi-field"><label for="wi-catchment">Students in the catchment</label>
          <input type="number" id="wi-catchment" value="1000" min="100" max="50000" step="100"/></div>
        <div class="wi-field"><label for="wi-growth">Monthly growth (%)</label>
          <input type="number" id="wi-growth" value="5" min="-20" max="50" step="0.5"/></div>
        <div class="wi-field"><label for="wi-competitor">Students lost to a competitor (%)</label>
          <input type="number" id="wi-competitor" value="0" min="0" max="60" step="5"/></div>
      </div></fieldset>
<fieldset class="wi-group"><legend>Capacity and hours</legend><div class="wi-grid">
        <div class="wi-field"><label for="wi-seats">Seats</label>
          <input type="number" id="wi-seats" value="60" min="10" max="500" step="5"/></div>
        <div class="wi-field"><label for="wi-hours">Opening hours a day</label>
          <input type="number" id="wi-hours" value="12" min="4" max="24" step="1"/></div>
        <div class="wi-field"><label for="wi-occupancy">Seats in use on average (%)</label>
          <input type="number" id="wi-occupancy" value="70" min="10" max="100" step="5"/></div>
      </div></fieldset>
<fieldset class="wi-group"><legend>Prices and membership</legend><div class="wi-grid">
        <div class="wi-field"><label for="wi-price">Price change on today's spend (%)</label>
          <input type="number" id="wi-price" value="0" min="-30" max="100" step="5"/></div>
        <div class="wi-field"><label for="wi-memtarget">Interested students who sign up (%)</label>
          <input type="number" id="wi-memtarget" value="50" min="5" max="95" step="5"/></div>
        <div class="wi-field"><label for="wi-tier1price">Basic membership, per month (₹)</label>
          <input type="number" id="wi-tier1price" value="750" min="200" max="5000" step="50"/></div>
        <div class="wi-field"><label for="wi-tier1uptake">Members on Basic (%)</label>
          <input type="number" id="wi-tier1uptake" value="60" min="10" max="90" step="5"/></div>
        <div class="wi-field"><label for="wi-tier2price">Premium membership, per month (₹)</label>
          <input type="number" id="wi-tier2price" value="1500" min="500" max="10000" step="100"/></div>
      </div></fieldset>
<fieldset class="wi-group"><legend>Costs</legend><div class="wi-grid">
        <div class="wi-field"><label for="wi-fixedcosts">Fixed costs, per month (₹)</label>
          <input type="number" id="wi-fixedcosts" value="150000" min="10000" max="2000000" step="10000"/></div>
        <div class="wi-field"><label for="wi-staffcost">Staff cost, per shift (₹)</label>
          <input type="number" id="wi-staffcost" value="800" min="200" max="5000" step="100"/></div>
      </div></fieldset>
      <div style="display:flex;gap:10px;margin-top:18px;flex-wrap:wrap">
        <button class="wi-run-btn" id="wi-run-btn" onclick="runWhatIf()">Run scenario</button>
        <button class="wi-run-btn" id="wi-save-btn"
                style="background:var(--accent3);display:none" onclick="saveScenario()">
          Save and compare
        </button>
      </div>
    </div>

    <!-- Results -->
    <div class="wi-results" id="wi-results">

      <!-- KPI strip -->
      <div class="mini-kpi-row" id="wi-kpi-strip"></div>

      <!-- Tab strip -->
      <div class="tab-strip" style="margin-top:24px">
        <button class="tab-btn active" onclick="wiTab(\'revenue\')">Revenue</button>
        <button class="tab-btn" onclick="wiTab(\'breakeven\')">Break-Even</button>
        <button class="tab-btn" onclick="wiTab(\'pricing\')">Pricing</button>
        <button class="tab-btn" onclick="wiTab(\'competitor\')">Competitor</button>
        <button class="tab-btn" onclick="wiTab(\'staffing\')">Staffing</button>
        <button class="tab-btn" onclick="wiTab(\'tiers\')">Tiers</button>
        <button class="tab-btn" onclick="wiTab(\'compare\')">Compare</button>
      </div>

      <!-- TAB: Revenue -->
      <div id="wi-tab-revenue" class="wi-tab-pane">
        <div class="chart-grid-2" style="margin-top:18px">
          <div class="chart-card">
            <div class="cc-title">12-Month Revenue Projection</div>
            <div class="cc-sub" id="wi-proj-sub">Monthly revenue with growth rate applied</div>
            <canvas id="wi-ch-projection" height="260"></canvas>
          </div>
          <div class="chart-card">
            <div class="cc-title">Revenue by Time Slot</div>
            <div class="cc-sub">Daily revenue breakdown across study time slots</div>
            <canvas id="wi-ch-slots" height="260"></canvas>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px">
          <div class="cc-title">Footfall Sensitivity Analysis (±20%)</div>
          <div class="cc-sub">How revenue changes if actual footfall deviates from projection</div>
          <div class="tbl-wrap" style="margin-top:12px">
            <table class="sens-table" id="wi-sens-table">
              <thead><tr><th>Footfall Scenario</th><th>Daily Visitors</th>
                <th>Daily Revenue (₹)</th><th>Monthly Revenue (₹)</th></tr></thead>
              <tbody id="wi-sens-tbody"></tbody>
            </table>
          </div>
          <canvas id="wi-ch-sensitivity" height="180" style="margin-top:16px"></canvas>
        </div>
        <div class="chart-grid-2" style="margin-top:18px">
          <div class="chart-card">
            <div class="cc-title">Membership Revenue (Single Tier)</div>
            <div class="cc-sub">Monthly vs Annual walk-in and membership revenue</div>
            <canvas id="wi-ch-membership" height="240"></canvas>
          </div>
          <div class="chart-card">
            <div class="cc-title">Seating Utilisation</div>
            <div class="cc-sub">Daily demand vs what the seats can host vs visitors actually served</div>
            <canvas id="wi-ch-seating" height="240"></canvas>
          </div>
        </div>
        <div class="wi-note" id="wi-arima-note" style="margin-top:18px">
          <strong>On ARIMA Applicability</strong><br>
          <span id="wi-arima-text">Loading...</span>
        </div>
      </div>

      <!-- TAB: Break-Even -->
      <div id="wi-tab-breakeven" class="wi-tab-pane" style="display:none">
        <div class="mini-kpi-row" id="wi-be-kpis" style="margin-top:18px"></div>
        <div class="chart-grid-2" style="margin-top:18px">
          <div class="chart-card">
            <div class="cc-title">12-Month Profit Projection</div>
            <div class="cc-sub">Revenue minus fixed and variable costs. Green = profitable, red = loss.</div>
            <canvas id="wi-ch-profit" height="260"></canvas>
          </div>
          <div class="chart-card">
            <div class="cc-title">Cost vs Revenue (Month 1)</div>
            <div class="cc-sub">Fixed costs + variable costs vs total revenue at base scenario</div>
            <canvas id="wi-ch-breakeven-bar" height="260"></canvas>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px">
          <div class="cc-title">Break-Even Summary</div>
          <div id="wi-be-summary" style="padding:16px;font-size:.85rem;color:var(--muted2);line-height:1.8"></div>
        </div>
      </div>

      <!-- TAB: Pricing Sensitivity -->
      <div id="wi-tab-pricing" class="wi-tab-pane" style="display:none">
        <div class="chart-grid-2" style="margin-top:18px">
          <div class="chart-card">
            <div class="cc-title">Revenue vs Price Adjustment</div>
            <div class="cc-sub">Monthly revenue across -30% to +50% price changes</div>
            <canvas id="wi-ch-price-sweep-rev" height="280"></canvas>
          </div>
          <div class="chart-card">
            <div class="cc-title">Profit vs Price Adjustment</div>
            <div class="cc-sub">Monthly profit (after costs) — bars turn red when loss-making</div>
            <canvas id="wi-ch-price-sweep-profit" height="280"></canvas>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px">
          <div class="cc-title">Pricing Sweep Table</div>
          <div class="tbl-wrap" style="margin-top:12px">
            <table class="sens-table" id="wi-price-sweep-table">
              <thead><tr><th>Price Adjustment</th><th>Adj. Spend (₹)</th>
                <th>Revenue for the year (₹)</th><th>Profit for the year (₹)</th><th>Profit in month 12 (₹)</th><th>Status</th></tr></thead>
              <tbody id="wi-price-sweep-tbody"></tbody>
            </table>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px" id="wi-sgrid-card">
          <div class="cc-title">Profit for the year: price against one other setting</div>
          <div class="cc-sub" id="wi-sgrid-sub">Each cell runs the full scenario with that price change and that value</div>
          <div class="seg-tabs" role="group" aria-label="Compare price with" id="wi-sgrid-tabs"></div>
          <p class="card-explain" data-title="Reading this table" id="wi-sgrid-read"></p>
          <div class="tbl-wrap" style="margin-top:10px"><table class="sens-table sgrid" id="wi-sgrid"></table></div>
          <p class="sgrid-key" id="wi-sgrid-key"></p>
        </div>
      </div>

      <!-- TAB: Competitor -->
      <div id="wi-tab-competitor" class="wi-tab-pane" style="display:none">
        <div class="mini-kpi-row" id="wi-comp-kpis" style="margin-top:18px"></div>
        <div class="chart-grid-2" style="margin-top:18px">
          <div class="chart-card">
            <div class="cc-title">Annual Revenue vs Competitor Impact</div>
            <div class="cc-sub">How annual revenue drops as a competitor draws away students</div>
            <canvas id="wi-ch-comp-rev" height="280"></canvas>
          </div>
          <div class="chart-card">
            <div class="cc-title">Revenue Loss vs Competitor Impact</div>
            <div class="cc-sub">Absolute annual revenue lost at each competitor scenario</div>
            <canvas id="wi-ch-comp-loss" height="280"></canvas>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px">
          <div class="cc-title">Competitor Scenario Table</div>
          <div class="tbl-wrap" style="margin-top:12px">
            <table class="sens-table" id="wi-comp-table">
              <thead><tr><th>Scenario</th><th>Effective Catchment</th>
                <th>Annual Revenue (₹)</th><th>Revenue Loss (₹)</th></tr></thead>
              <tbody id="wi-comp-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- TAB: Staffing -->
      <div id="wi-tab-staffing" class="wi-tab-pane" style="display:none">
        <div class="mini-kpi-row" id="wi-staff-kpis" style="margin-top:18px"></div>
        <div class="chart-grid-2" style="margin-top:18px">
          <div class="chart-card">
            <div class="cc-title">Staff Required per Time Slot</div>
            <div class="cc-sub">Minimum staff based on peak concurrent visitors per slot</div>
            <canvas id="wi-ch-staff-count" height="260"></canvas>
          </div>
          <div class="chart-card">
            <div class="cc-title">Labour Cost vs Slot Revenue</div>
            <div class="cc-sub">Labour cost and revenue per slot — profit after labour shown</div>
            <canvas id="wi-ch-staff-profit" height="260"></canvas>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px">
          <div class="cc-title">Staffing Detail by Slot</div>
          <div class="tbl-wrap" style="margin-top:12px">
            <table class="sens-table" id="wi-staff-table">
              <thead><tr><th>Slot</th><th>Footfall</th><th>Staff Needed</th>
                <th>Labour Cost (₹)</th><th>Slot Revenue (₹)</th><th>Profit After Labour (₹)</th></tr></thead>
              <tbody id="wi-staff-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- TAB: Tiers -->
      <div id="wi-tab-tiers" class="wi-tab-pane" style="display:none">
        <div class="mini-kpi-row" id="wi-tier-kpis" style="margin-top:18px"></div>
        <div class="chart-grid-2" style="margin-top:18px">
          <div class="chart-card">
            <div class="cc-title">Tier Revenue Split (Monthly)</div>
            <div class="cc-sub">Basic vs Premium monthly membership revenue</div>
            <canvas id="wi-ch-tier-split" height="260"></canvas>
          </div>
          <div class="chart-card">
            <div class="cc-title">Two-Tier vs Single-Tier Comparison</div>
            <div class="cc-sub">Monthly revenue comparison — two tiers vs single average price</div>
            <canvas id="wi-ch-tier-compare" height="260"></canvas>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px">
          <div class="cc-title">Membership Tier Summary</div>
          <div id="wi-tier-summary" style="padding:16px;font-size:.85rem;color:var(--muted2);line-height:1.8"></div>
        </div>
      </div>

      <!-- TAB: Compare Scenarios -->
      <div id="wi-tab-compare" class="wi-tab-pane" style="display:none">
        <div style="margin-top:18px;font-size:.83rem;color:var(--muted2);margin-bottom:12px">
          Run different parameter combinations and click <strong style="color:var(--accent3)">Save and compare</strong>
          to build a side-by-side comparison. Up to 5 scenarios saved.
        </div>
        <div class="chart-card" style="margin-top:4px">
          <div class="cc-title">Saved Scenario Comparison</div>
          <div class="tbl-wrap" style="margin-top:12px">
            <table class="sens-table" id="wi-compare-table">
              <thead><tr>
                <th>#</th><th>Catchment</th><th>Growth %</th><th>Price Adj %</th>
                <th>Competitor %</th><th>Fixed Costs</th>
                <th>Monthly Rev (₹)</th><th>Annual Rev (₹)</th>
                <th>Profit M1 (₹)</th><th>Break-Even Mo.</th>
                <th>Members</th><th>Tier Rev/Mo (₹)</th>
                <th></th>
              </tr></thead>
              <tbody id="wi-compare-tbody">
                <tr><td colspan="13" style="text-align:center;color:var(--muted2);padding:20px">
                  No scenarios saved yet. Run a scenario and click Save and compare.
                </td></tr>
              </tbody>
            </table>
          </div>
        </div>
        <div class="chart-card" style="margin-top:18px">
          <div class="cc-title">Annual Revenue — Scenario Comparison</div>
          <div class="cc-sub">Side-by-side annual revenue for each saved scenario</div>
          <canvas id="wi-ch-compare" height="240"></canvas>
        </div>
      </div>

    </div><!-- /wi-results -->
  </div><!-- /wi-body -->
</div>
'''
_DESIGN_CSS = r"""
html[data-view="surface"] .deep-only{display:none!important}
html[data-view="deep"] .surface-only{display:none!important}
#app-nav .view-switch button[aria-pressed="true"]{background:var(--nav-on-bg);color:var(--nav-on-fg);font-weight:600}
.view-banner{display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap;margin:0 0 20px;padding:12px 16px;border:1px solid var(--border);border-left:3px solid var(--accent);border-radius:6px;background:var(--panel)}
.view-banner p{margin:0;font-size:14px;color:var(--muted2);max-width:80ch}
.view-banner strong{color:var(--text)}
.owner-rules{margin-bottom:18px}
.or-table td{vertical-align:top}
.or-table td:first-child{min-width:260px}
.or-rule{font-weight:600;color:var(--text);line-height:1.45}
.or-rule span{color:var(--accent);font-weight:700}
.or-plain{font-size:12.5px;color:var(--muted2);margin-top:3px;line-height:1.4}
.or-empty{color:var(--muted2);font-size:14px}

.ps-controls{display:flex;flex-wrap:wrap;align-items:center;gap:10px 16px;margin:18px 0 0}
.ps-seg{display:inline-flex;border:1px solid var(--border2);border-radius:7px;overflow:hidden}
.ps-seg-btn{font:inherit;font-size:13px;font-weight:600;padding:7px 14px;border:0;background:var(--panel);color:var(--text);cursor:pointer}
.ps-seg-btn+.ps-seg-btn{border-left:1px solid var(--border2)}
.ps-seg-btn.on{background:var(--accent);color:var(--panel)}
.ps-seg-btn:focus-visible{outline:2px solid var(--accent);outline-offset:-2px}
.ps-chips{display:flex;flex-wrap:wrap;align-items:center;gap:6px}
.ps-best{margin-top:14px}
.sens-table tr.best-row td{background:rgba(31,122,77,.08)}

/* v48 interactions */
.chart-card{position:relative}
.card-tools{position:absolute;top:16px;right:16px;display:flex;gap:6px;align-items:flex-start;z-index:3}
.tool-btn{font:inherit;font-size:12.5px;font-weight:600;line-height:1;padding:7px 10px;border-radius:6px;border:1px solid var(--border2);background:var(--panel);color:var(--text);cursor:pointer;white-space:nowrap}
.tool-btn:hover:not(:disabled){border-color:var(--accent);color:var(--accent)}
.tool-btn:disabled{opacity:.45;cursor:default}
.tool-btn:focus-visible,.drill-chip:focus-visible,.lev-goto:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.menu-wrap{position:relative}
.tool-menu{position:absolute;right:0;top:calc(100% + 6px);min-width:220px;background:var(--panel);border:1px solid var(--border2);border-radius:8px;padding:8px;box-shadow:0 8px 24px rgba(0,0,0,.14);z-index:20}
.tm-head{font-size:12px;font-weight:600;color:var(--muted2);padding:4px 6px 6px}
.tm-row{display:flex;align-items:center;gap:8px;padding:7px 6px;border-radius:5px;font-size:13.5px;color:var(--text);cursor:pointer}
.tm-row:hover{background:var(--bg)}
.tm-row input{accent-color:var(--accent);width:15px;height:15px;margin:0}
.tm-sw{width:12px;height:12px;border-radius:3px;flex:none}
.tm-link{display:block;width:100%;margin-top:4px;padding:7px 6px;border:0;border-top:1px solid var(--border);background:none;color:var(--accent);font:inherit;font-size:13px;font-weight:600;text-align:left;cursor:pointer}
.drill-bar{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin:4px 0 10px}
.drill-label{font-size:12.5px;font-weight:600;color:var(--muted2);margin-right:2px}
.drill-chip{font:inherit;font-size:12.5px;padding:5px 10px;border-radius:999px;border:1px solid var(--border2);background:var(--panel);color:var(--text);cursor:pointer;transition:background .15s,color .15s,border-color .15s}
.drill-chip:hover{border-color:var(--accent)}
.drill-chip.on{background:var(--accent);border-color:var(--accent);color:var(--panel)}
.drill-hint{flex-basis:100%;font-size:12.5px;color:var(--muted2)}
.cv-box{transition:opacity .18s ease,transform .18s ease,height .3s ease}
.cv-box.drill-out{opacity:0;transform:scale(.985)}
.zoom-hint{font-size:12.5px;color:var(--muted2);margin-top:6px}
.rule-callout{position:absolute;width:300px;max-width:calc(100% - 8px);background:var(--panel);color:var(--text);border:1px solid var(--border2);border-radius:8px;padding:12px 14px;box-shadow:0 10px 28px rgba(0,0,0,.16);font-size:13px;line-height:1.45;z-index:4;opacity:0;transform:translateY(4px);transition:opacity .2s ease,transform .2s ease}
.rule-callout.show{opacity:1;transform:none}
.rule-callout::before{content:"";position:absolute;width:10px;height:10px;background:var(--panel);border:1px solid var(--border2);transform:rotate(45deg)}
.rule-callout[data-side="right"]::before{left:-6px;top:calc(var(--arrow-y) - 5px);border-top:0;border-right:0}
.rule-callout[data-side="left"]::before{right:-6px;top:calc(var(--arrow-y) - 5px);border-bottom:0;border-left:0}
.rule-callout[data-side="below"]::before{top:-6px;left:calc(var(--arrow-x) - 5px);border-right:0;border-bottom:0}
.rc-close{position:absolute;top:6px;right:8px;border:0;background:none;color:var(--muted2);font-size:18px;line-height:1;cursor:pointer;padding:4px}
.rc-rule{font-weight:600;padding-right:18px;margin-bottom:6px}
.rc-rule span{color:var(--accent);font-weight:700}
.rc-plain{margin:0 0 8px;color:var(--muted2)}
.rc-dl{margin:0;display:flex;flex-direction:column;gap:4px}
.rc-dl div{display:grid;grid-template-columns:86px 1fr;gap:8px}
.rc-dl dt{font-weight:600}
.rc-dl dd{margin:0;color:var(--muted2)}
.hm-legend{display:flex;flex-wrap:wrap;gap:6px 14px;margin-top:10px;font-size:12.5px;color:var(--muted2)}
.hm-key{display:inline-flex;align-items:center;gap:6px}
.hm-sw{width:14px;height:14px;border-radius:3px;border:1px solid var(--border)}
.hm-cap{font-weight:600;color:var(--text)}
.chart-stack{display:flex;flex-direction:column;gap:18px}
.card-explain{margin:6px 0 16px;padding:14px 16px;border:1px solid var(--border);border-left:3px solid var(--accent3);border-radius:6px;font-size:13.5px;line-height:1.55;color:var(--muted2);max-width:960px}
.card-explain h4{font-size:14px;font-weight:600;color:var(--text);margin:12px 0 4px}
.card-explain h4:first-child{margin-top:0}
.card-explain p{margin:0 0 6px}
.ax-dl{margin:4px 0 6px;display:flex;flex-direction:column;gap:6px}
.ax-dl div{display:grid;grid-template-columns:110px 1fr;gap:10px}
.ax-dl dt{font-weight:600;color:var(--text)}
.ax-dl dd{margin:0}
.ax-tbl{max-width:520px;margin:4px 0 6px}
.sig{display:inline-block;font-size:12px;font-weight:600;padding:2px 8px;border-radius:999px}
.sig-yes{background:rgba(31,122,77,.12);color:var(--accent2)}
.sig-no{background:rgba(85,99,93,.12);color:var(--muted2)}
#pred-levers{scroll-margin-top:20px}
@media (max-width:860px){#pred-levers{scroll-margin-top:76px}}
.lev-head{display:flex;justify-content:space-between;align-items:flex-start;gap:16px}
.lev-intro{font-size:14px;color:var(--muted2);max-width:78ch;margin:-4px 0 10px}
.lev-status{font-size:13.5px;color:var(--muted2);margin:0 0 14px;padding:10px 14px;border-radius:6px;border:1px solid var(--border);background:var(--panel);max-width:960px}
.lev-status.on{border-left:3px solid var(--accent);color:var(--text)}
.lev-status.err{border-left:3px solid var(--danger);color:var(--danger)}
.lev-prep{font-size:12.5px;color:var(--muted2);margin:-6px 0 14px}
.lev-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:18px}
.lev-card-head{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}
.sec-head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
.sec-head .section-title{margin-bottom:14px}
.lev-goto{font:inherit;font-size:12.5px;font-weight:600;padding:6px 10px;border-radius:6px;border:1px solid var(--border2);background:var(--panel);color:var(--accent);cursor:pointer;margin-bottom:14px}
.lev-goto:hover{border-color:var(--accent)}
.lever-note{margin:10px 0 0;padding:8px 12px;border-radius:6px;background:rgba(138,90,0,.08);border-left:3px solid var(--warn);font-size:13px;line-height:1.45;color:var(--text)}
@media (max-width:980px){.lev-grid{grid-template-columns:1fr}}
@media (max-width:640px){.card-tools{position:static;justify-content:flex-end;margin:-4px 0 8px}.has-tools .cc-title{padding-right:0!important}.ax-dl div,.rc-dl div{grid-template-columns:1fr;gap:0}.lev-head{flex-direction:column}}
@media (prefers-reduced-motion:reduce){.cv-box,.rule-callout,.drill-chip{transition:none}}

.chart-guide{margin-top:16px;padding-top:12px;border-top:1px solid var(--border);font-size:13px;line-height:1.5;color:var(--muted2)}
.chart-guide .cg-head{font-size:13px;font-weight:600;color:var(--text);margin-bottom:4px}
.chart-guide .cg-p{margin:0 0 8px}
.chart-guide .cg-axes{margin:0;display:flex;flex-direction:column;gap:4px}
.chart-guide .cg-axes div{display:grid;grid-template-columns:128px 1fr;gap:10px}
.chart-guide dt{font-weight:600;color:var(--text)}
.chart-guide dd{margin:0}
@media (max-width:640px){.chart-guide .cg-axes div{grid-template-columns:1fr;gap:0}}
.wi-help{font-size:12.5px;line-height:1.45;color:var(--muted2);margin-top:2px}
.wi-assumptions{margin:18px 0 0;padding-top:12px;border-top:1px solid var(--border);font-size:13px;line-height:1.5;color:var(--muted2)}

/* ════════════════════════════════════════════════════════════
   v42 DESIGN LAYER — "Reading room"  (see DESIGN.md)
   ════════════════════════════════════════════════════════════ */
:root{
  --bg:#F5F7F4;--panel:#FFFFFF;--card:#FFFFFF;
  --border:#DCE3DD;--border2:#C5D0C8;
  --accent:#1F5C4A;--accent2:#1F7A4D;--accent3:#8F6212;
  --warn:#8A5A00;--danger:#B3261E;--info:#17231F;--purple:#17231F;
  --text:#17231F;--muted:#5F6C66;--muted2:#55635D;
  --r:6px;--r2:4px;
  --c0:#2b6ea3;--c1:#d2692c;--c2:#2f8f6a;--c3:#cfae2a;
  --nav-w:232px;
  --nav-bg:#1F5C4A;--nav-fg:#FFFFFF;--nav-on-bg:#FFFFFF;--nav-on-fg:#1F5C4A;
  --on-accent:#FFFFFF;--accent-hover:#17483A;--focus-ring:rgba(31,92,74,.18);
  --hover:#F1F5F1;--chip:#EEF2EE;--chip2:#E2E8E2;--chip-good:#E6F2EA;--chip-warn:#F7EEDC;
  --base-row:#F7F3E8;--err-bg:#FBEDEC;--err-border:#E8BDB9;--tip-bg:#17231F;--tip-fg:#FFFFFF;
  color-scheme:light;
  --f-head:'Bricolage Grotesque','Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif;
  --f-body:'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif;
}
html{font-size:16px}
body{background:var(--bg);color:var(--text);font-family:var(--f-body);font-size:14px;line-height:1.5;font-variant-numeric:tabular-nums;-webkit-font-smoothing:antialiased}
h1,h2,h3,h4{font-family:var(--f-head);font-weight:600;letter-spacing:-.01em;color:var(--text)}
*{text-transform:none!important;letter-spacing:normal!important}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-track{background:var(--bg)}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:5px}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
button{font-family:var(--f-body)}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}

/* one screen at a time — fixes sections opening out of sight below the menu */
.screen{display:none!important;animation:none!important}
.screen.active{display:block!important}
#s-upload.active,#s-progress.active{display:flex!important}

/* ── App shell: persistent navigation ── */
#app-nav{display:none;position:fixed;left:0;top:0;bottom:0;width:var(--nav-w);background:var(--nav-bg);color:var(--nav-fg);padding:24px 14px 18px;flex-direction:column;gap:2px;z-index:150;overflow-y:auto}
body.has-nav #app-nav{display:flex}
body.has-nav .screen.active{margin-left:var(--nav-w)}
#app-nav .an-brand{font-family:var(--f-head);font-weight:600;font-size:20px;line-height:1.15;padding:0 10px 4px}
#app-nav .an-sub{font-size:12px;color:rgba(255,255,255,.72);padding:0 10px 20px}
#app-nav button{all:unset;box-sizing:border-box;display:block;width:100%;padding:9px 10px;border-radius:6px;font-family:var(--f-body);font-size:14px;color:rgba(255,255,255,.86);cursor:pointer}
#app-nav button:hover{background:rgba(255,255,255,.10);color:var(--nav-fg)}
#app-nav button:focus-visible{outline:2px solid var(--nav-fg);outline-offset:-2px}
#app-nav button.on{background:var(--nav-on-bg);color:var(--nav-on-fg);font-weight:600}
#app-nav .an-gap{flex:1;min-height:20px}
#app-nav .an-rule{height:1px;background:rgba(255,255,255,.22);margin:10px 6px}
#app-nav .an-facts{font-size:12px;line-height:1.5;color:rgba(255,255,255,.78);padding:0 10px 10px}
#app-nav .an-facts b{display:block;font-family:var(--f-head);font-size:20px;font-weight:600;color:var(--nav-fg)}
@media (max-width:860px){
  #app-nav{position:sticky;top:0;bottom:auto;width:auto;flex-direction:row;align-items:center;gap:4px;padding:8px 10px;overflow-x:auto;white-space:nowrap}
  body.has-nav #app-nav{display:flex}
  body.has-nav .screen.active{margin-left:0}
  #app-nav .an-brand{font-size:15px;padding:0 8px 0 2px}
  #app-nav .an-sub,#app-nav .an-gap,#app-nav .an-rule,#app-nav .an-facts{display:none}
  #app-nav button{width:auto;padding:8px 10px;flex:0 0 auto}
}

/* ── Upload & progress ── */
#s-upload,#s-progress{background:var(--bg)}
.upload-header h1{font-size:34px;font-weight:600}
.upload-header h1 em{color:inherit;font-style:normal}
.upload-header p,.prog-sub{color:var(--muted2);font-size:16px}
.drop-card{background:var(--panel);border:1.5px dashed var(--border2);border-radius:8px;padding:44px 28px}
.drop-card:hover,.drop-card.over{border-color:var(--accent);background:var(--panel)}
.drop-card .icon{display:none}
.drop-card h3{font-size:16px}
.drop-card strong{color:var(--accent);text-decoration:underline}
.file-chosen{color:var(--accent2);background:var(--chip-good);border-radius:6px}
.up-prog-track,.big-prog-track,#export-prog-track,.score-bar-bg,.fi-track,.pred-bar-track,.pred-factor-bar-track{background:var(--border)!important}
.up-prog-fill,.big-prog-fill,#export-prog-fill,#dl-progress-bar{background:var(--accent)!important;box-shadow:none!important}
.prog-title{font-family:var(--f-head);font-size:26px;font-weight:600}
.prog-pct{font-family:var(--f-head);color:var(--text)}
.step-row{color:var(--muted)}.step-row.active{color:var(--text)}.step-row.done{color:var(--accent2)}
.step-row.active .step-dot{background:var(--accent);box-shadow:none}.step-row.done .step-dot{background:var(--accent2)}
.reconnect-spinner,.spinner{border-color:var(--border2);border-top-color:var(--accent)}

/* ── Overview (was the tile menu) ── */
#s-nav{display:none;min-height:0!important;padding:40px 40px 64px!important;align-items:stretch!important;justify-content:flex-start!important}
.nav-header{text-align:left;margin-bottom:28px;max-width:720px}
.nav-header h1{font-size:34px;font-weight:600}
.nav-header h1 em{color:inherit;font-style:normal}
.nav-header p{font-size:16px;color:var(--muted2)}
.nav-grid{display:block!important;max-width:760px;background:var(--panel);border:1px solid var(--border);border-radius:var(--r)}
.nav-tile{display:grid;grid-template-columns:220px 1fr;gap:16px;align-items:baseline;text-align:left;background:none;border:0!important;border-bottom:1px solid var(--border)!important;border-radius:0;padding:16px 20px;transition:none}
.nav-tile::before{display:none}
.nav-tile:hover{transform:none;box-shadow:none;background:var(--hover)}
.nav-tile .t-icon{display:none}
.nav-tile h3{font-size:16px;margin:0;color:var(--accent);text-decoration:underline;text-underline-offset:3px}
.nav-tile p{font-size:14px;color:var(--muted2)}
.nav-tile-utility{display:flex!important;flex-direction:row!important;gap:10px;background:none;border:0;border-radius:0;padding:16px 20px}
.util-btn{background:var(--panel);border:1px solid var(--border2);border-radius:6px;padding:9px 14px;color:var(--text);transition:none}
.util-btn:hover{background:var(--hover);border-color:var(--accent);transform:none;box-shadow:none}
.util-btn .ub-icon{display:none}
.util-btn .ub-title{font-size:14px;font-weight:600;color:var(--text)}
.util-btn .ub-sub{font-size:12px;color:var(--muted2)}
.util-btn.danger .ub-title{color:var(--danger)}
.util-btn.danger:hover,.util-btn.success:hover{border-color:var(--accent);background:var(--hover)}
@media (max-width:860px){#s-nav{padding:24px 16px 48px!important}.nav-tile{grid-template-columns:1fr;gap:2px}.nav-tile-utility{flex-direction:column!important}}

/* ── Section header bar ── */
.topbar{position:sticky;top:0;z-index:20;background:var(--bg);border-bottom:1px solid var(--border);padding:18px 40px 14px;gap:12px}
.topbar h2{font-size:26px;font-weight:600;line-height:1.2}
body.has-nav .back-btn{display:none}
.back-btn,.dl-btn,.view-more-btn,.pred-retry,#export-modal-cancel{background:var(--panel)!important;color:var(--text)!important;border:1px solid var(--border2)!important;border-radius:6px!important;font-size:14px;font-weight:600;box-shadow:none!important;transition:none}
.back-btn:hover,.dl-btn:hover,.view-more-btn:hover,.pred-retry:hover,#export-modal-cancel:hover{border-color:var(--accent)!important;background:var(--hover)!important;transform:none}
.wi-run-btn,.pred-submit{background:var(--accent)!important;color:var(--on-accent)!important;border:1px solid var(--accent)!important;border-radius:6px!important;font-size:14px;font-weight:600;box-shadow:none!important;transition:none;padding:10px 18px}
.wi-run-btn:hover,.pred-submit:hover{background:var(--accent-hover)!important;transform:none!important;opacity:1}
.wi-run-btn:disabled,.pred-submit:disabled{opacity:.55}
#wi-save-btn{background:var(--panel)!important;color:var(--accent)!important;border:1px solid var(--accent)!important}
@media (max-width:860px){.topbar{position:static;padding:14px 16px 10px;flex-wrap:wrap}.topbar h2{font-size:20px}}

/* ── Page bodies ── */
.sum-body,.dash-body,.pred-body,.wi-body,.predictor-body{max-width:1240px;margin:0;padding:24px 40px 72px}
@media (max-width:860px){.sum-body,.dash-body,.pred-body,.wi-body,.predictor-body{padding:16px 16px 56px}}
.section-title,.sum-section h3,.pred-section-label{font-family:var(--f-head);font-size:20px;font-weight:600;color:var(--text);margin:8px 0 14px;gap:0}
.section-title::after,.sum-section h3::after,.pred-section-label::after{display:none}
.dash-section,.sum-section{margin-bottom:36px}

/* ── Key figures: one ledger row, not a row of coloured boxes ── */
.mini-kpi-row,.kpi-grid{display:flex;flex-wrap:wrap;gap:0;background:var(--panel);border:1px solid var(--border);border-radius:var(--r);margin-bottom:24px;overflow:hidden}
.mini-kpi,.kpi{flex:1 1 150px;background:none!important;border:0!important;border-right:1px solid var(--border)!important;border-radius:0!important;padding:16px 20px;text-align:left;box-shadow:none!important;margin:0 -1px -1px 0;border-bottom:1px solid var(--border)!important}
.mini-kpi .mk-val,.kpi .k-val{font-family:var(--f-head);font-size:26px;font-weight:600;line-height:1.15;color:var(--text)}
.mini-kpi .mk-lab,.kpi .k-label,.kpi .k-sub{font-size:13px;color:var(--muted2);margin-top:4px}
.kpi .k-label{margin:0 0 4px}
/* figures stay ink unless the value is a status */
.mk-val[style*="--accent)"],.mk-val[style*="--accent2)"],.mk-val[style*="--accent3)"],.mk-val[style*="--info)"],.mk-val[style*="--purple)"],.mk-val[style*="--warn)"],.mk-val[style*="--muted2)"]{color:var(--text)!important}
.kpi .k-val[class*="c-"]{color:var(--text)!important}
.clean-table tbody td:nth-child(3) *{color:var(--text)!important}

/* ── Cards, charts ── */
.chart-card,.rel-card,.col-card,.model-card,.reg-insight,.wi-controls,.wi-note,.pred-result-card,.pred-history,.pred-intro,.tbl-wrap{background:var(--panel);border:1px solid var(--border);border-radius:var(--r);box-shadow:none}
.chart-card{padding:20px}
.chart-card .cc-title{font-family:var(--f-body);font-size:14px;font-weight:600;color:var(--text)}
.chart-card .cc-sub{font-size:13px;color:var(--muted2);margin-bottom:14px}
.chart-grid-2,.chart-grid-3,.chart-grid-4,.chart-grid-1-2,.chart-grid-2-1{gap:16px;margin-bottom:16px}
@media (max-width:860px){.chart-grid-2,.chart-grid-3,.chart-grid-4,.chart-grid-1-2,.chart-grid-2-1,.rel-grid,.col-grid,.model-grid{grid-template-columns:1fr!important}.chart-card.span2{grid-column:auto}}
.reg-insight{border-left:3px solid var(--accent3)}
.reg-insight h4{font-family:var(--f-body);font-size:14px;font-weight:600;color:var(--text)}
.reg-insight p,.wi-note{font-size:14px;color:var(--muted2);line-height:1.55;max-width:78ch}
.wi-note strong{color:var(--text)}
.rel-card .rc-title{font-size:14px;font-weight:600;color:var(--text)}
.rel-card .rc-desc{font-size:13px;color:var(--muted2)}
.rel-card .rc-val{font-family:var(--f-head);font-size:26px;font-weight:600;color:var(--text)!important}
.rc-row,.col-card .cc-row{border-color:var(--border);font-size:13px}
.rc-row .rk,.col-card .cc-k{color:var(--muted2)}.rc-row .rv,.col-card .cc-v{color:var(--text);font-weight:600}
.col-card .cc-name{font-size:14px;font-weight:600;color:var(--text)}
.cluster-legend{display:none}

/* ── Chips ── */
.pill,.badge,.impact-tag,.tag-static,.tag-live,.export-option .eo-tag,.nav-tile .coming{background:var(--chip)!important;color:var(--muted2)!important;border:0!important;border-radius:4px!important;font-size:12px!important;font-weight:600;padding:2px 8px}
.pill-green{background:var(--chip-good)!important;color:var(--accent2)!important}
.pill-orange,.impact-high{background:var(--chip-warn)!important;color:var(--warn)!important}
.c-accent,.c-info,.c-orange,.c-warn{color:var(--text)}.c-green{color:var(--accent2)}.c-danger{color:var(--danger)}

/* ── Tables ── */
.clean-table,.sens-table,.arm-table,.scorecard-table,.pred-history-table{background:var(--panel);font-size:14px;border-collapse:collapse;width:100%}
.clean-table thead th,.sens-table thead th,.arm-table thead th,.scorecard-table thead th,.pred-history-table thead th{background:var(--panel)!important;color:var(--muted2)!important;font-size:13px;font-weight:600;text-align:left;border-bottom:1px solid var(--border2)!important;padding:10px 14px}
.clean-table tbody td,.sens-table tbody td,.arm-table tbody td,.scorecard-table tbody td,.pred-history-table tbody td{border-bottom:1px solid var(--border)!important;padding:10px 14px;color:var(--text);background:none}
.clean-table tbody tr:nth-child(even),.sens-table tbody tr:nth-child(even),.scorecard-table tbody tr:nth-child(even),.pred-history-table tbody tr:nth-child(even){background:none!important}
.arm-table tbody tr:hover,.scorecard-table tbody tr:hover{background:var(--hover)!important}
.clean-table tbody td:nth-child(2){color:var(--text)!important;font-weight:600;text-align:left}
.clean-table tbody td:nth-child(3){font-family:var(--f-head);font-weight:600;font-size:16px;color:var(--text)!important}
.clean-table tbody td:nth-child(4){color:var(--muted2);font-size:13px}
.sens-table .base-row{background:var(--base-row)!important}
.sens-table .base-row td{font-weight:600}
.arm-table .lift-bar,.score-bar-fill,.fi-fill,.pred-factor-bar-fill{border-radius:2px}
.arm-table .lift-bar{background:var(--accent)}
.rank-badge{background:var(--chip)!important;color:var(--muted2)!important;border-radius:4px;box-shadow:none}
.rank-badge.gold{background:var(--accent3)!important;color:var(--on-accent)!important}
.rank-badge.silver,.rank-badge.bronze{background:var(--chip2)!important;color:var(--text)!important}
.tbl-wrap{overflow-x:auto}

/* ── Tabs ── */
.tab-strip{gap:4px;border-bottom:1px solid var(--border);margin-bottom:18px;flex-wrap:wrap}
.tab-btn{background:none!important;border:0!important;border-bottom:2px solid transparent!important;border-radius:0!important;color:var(--muted2)!important;font-size:14px;font-weight:600;padding:9px 12px;margin-bottom:-1px;box-shadow:none!important;transition:none}
.tab-btn.active{color:var(--text)!important;border-bottom-color:var(--accent)!important}
.tab-btn:hover:not(.active){color:var(--text)!important;background:none!important}

/* ── Forms ── */
.wi-controls{padding:20px}
.wi-controls h3{font-size:20px}
.wi-field label,.pred-field label{font-size:13px;font-weight:600;color:var(--muted2)}
.wi-field input,.wi-field select,.pred-field select,.pred-field input[type=number]{background:var(--panel)!important;color:var(--text)!important;border:1px solid var(--border2)!important;border-radius:6px!important;font-family:var(--f-body);font-size:14px;padding:9px 10px}
.wi-field input:focus,.wi-field select:focus,.pred-field select:focus,.pred-field input:focus{border-color:var(--accent)!important;box-shadow:0 0 0 2px var(--focus-ring)!important;outline:none}
.pred-field select option{background:var(--panel);color:var(--text)}
.likert-row input[type=range]{accent-color:var(--accent)}
.likert-val{color:var(--text);font-family:var(--f-head);font-weight:600}
.pred-field.has-error label{color:var(--danger)}
.pred-field-error{color:var(--danger);font-size:13px}
.pred-error-banner{background:var(--err-bg);border:1px solid var(--err-border);color:var(--danger);border-radius:6px}

/* ── Predictor result ── */
.pred-verdict-text{font-family:var(--f-head);font-size:34px;font-weight:600}
.pred-verdict-sub,.pred-pct-sub,.pred-bar-labels,.pred-bar-caption{color:var(--muted2);font-size:13px}
.pred-pct{font-family:var(--f-head);font-weight:600}
.pred-bar-fill{background:var(--accent)!important}
.pred-bar-mid{background:var(--text)}
.pred-factors-title,.pred-history-title{font-size:14px;font-weight:600;color:var(--text)}
.pred-history-table .verdict-yes{color:var(--accent2)}.pred-history-table .verdict-no{color:var(--danger)}
.conf-cell{border-radius:6px}.conf-cell .cv{font-family:var(--f-head)}
.coef-row .cr-name,.fi-name{color:var(--text);font-size:13px}

/* ── Modals, tooltip ── */
#export-modal,#export-prog-box{background:var(--panel);border:1px solid var(--border);border-radius:8px;box-shadow:0 12px 40px rgba(23,35,31,.18);color:var(--text)}
#export-modal h3,#export-prog-box h3{font-size:20px}
#export-modal p,#export-prog-step,#export-prog-log{color:var(--muted2)}
.export-option{background:var(--panel);border:1px solid var(--border2);border-radius:6px;transition:none}
.export-option:hover{border-color:var(--accent);background:var(--hover);transform:none}
.export-option .eo-icon{display:none}
.export-option .eo-title{color:var(--text);font-weight:600}.export-option .eo-desc{color:var(--muted2)}
#tooltip{background:var(--tip-bg);color:var(--tip-fg);border:0;border-radius:6px;box-shadow:0 6px 20px rgba(23,35,31,.25);font-size:13px}
#tooltip strong{color:var(--tip-fg)}
#sum-loading p{color:var(--muted2)}

/* ── polish ── */
.pred-gauge-wrap{width:auto!important;height:auto!important;margin:8px 0 20px!important}
.pred-gauge-wrap canvas{display:none}
.pred-gauge-label{position:static!important;display:block!important;text-align:center}
.pred-pct{font-size:48px!important;line-height:1.1}
.nav-tile:focus-visible{outline:2px solid var(--accent);outline-offset:-2px}
.topbar h2{letter-spacing:-.01em!important}

/* ════════ DARK THEME — the same design, reading room at night ════════ */
:root[data-theme="dark"]{
  --bg:#101614;--panel:#18211E;--card:#18211E;
  --border:#2A3631;--border2:#3B4B44;
  --accent:#6FCFAE;--accent2:#5CC08A;--accent3:#D9A441;
  --warn:#E0B25A;--danger:#F0766C;--info:#E9EFEA;--purple:#E9EFEA;
  --text:#E9EFEA;--muted:#8C9992;--muted2:#A9B5AE;
  --c0:#3987e5;--c1:#d95926;--c2:#199e70;--c3:#c98500;
  --nav-bg:#13302A;--nav-fg:#FFFFFF;--nav-on-bg:#6FCFAE;--nav-on-fg:#0F1513;
  --on-accent:#0F1513;--accent-hover:#8ADBC0;--focus-ring:rgba(111,207,174,.28);
  --hover:#1F2A26;--chip:#243029;--chip2:#2D3B34;--chip-good:#173327;--chip-warn:#3A2F14;
  --base-row:#2B2718;--err-bg:#3A1C1A;--err-border:#6B2F2B;--tip-bg:#E9EFEA;--tip-fg:#101614;
  color-scheme:dark;
}
:root[data-theme="dark"] #app-nav{border-right:1px solid var(--border)}
:root[data-theme="dark"] #export-modal,:root[data-theme="dark"] #export-prog-box{box-shadow:0 12px 40px rgba(0,0,0,.5)}
/* older rules that paint fixed light or dark surfaces follow the theme */
.drop-card:hover,.drop-card.over{background:var(--panel)}
#export-modal-backdrop,#export-prog-backdrop{background:rgba(0,0,0,.45)}
#pred-loading,#analytics-loading,#sum-loading{background:var(--bg)!important;color:var(--text)}
code,.code-chip{color:var(--text)}

/* ── theme switch ── */
.theme-switch{display:flex;gap:2px;padding:2px;margin:0 6px 10px;border-radius:7px;background:rgba(255,255,255,.14)}
#app-nav .theme-switch button{flex:1;width:auto;text-align:center;padding:6px 8px;font-size:13px;border-radius:5px;color:rgba(255,255,255,.86)}
#app-nav .theme-switch button[aria-pressed="true"]{background:var(--nav-on-bg);color:var(--nav-on-fg);font-weight:600}
#theme-float{position:fixed;top:14px;right:16px;z-index:160;display:flex;gap:2px;padding:2px;border-radius:7px;background:var(--panel);border:1px solid var(--border2)}
#theme-float button{all:unset;box-sizing:border-box;cursor:pointer;font-family:var(--f-body);font-size:13px;padding:6px 12px;border-radius:5px;color:var(--muted2)}
#theme-float button[aria-pressed="true"]{background:var(--accent);color:var(--on-accent);font-weight:600}
#theme-float button:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
body.has-nav #theme-float{display:none}
@media (max-width:860px){.theme-switch{margin:0 0 0 6px;flex:0 0 auto}#app-nav .theme-switch button{padding:6px 10px}}
"""
HTML_PAGE = HTML_PAGE.replace('</style>\n</head>', _DESIGN_CSS + '</style>\n</head>')
HTML_PAGE = HTML_PAGE.replace('</body>\n</html>', _PREDICTOR_HTML + _WHATIF_HTML + '\n</body>\n</html>')

# ════════════════════════════════════════════════════════════
#  SERVER START
# ════════════════════════════════════════════════════════════
# browser edition: no server start; the Web Worker calls the app directly
