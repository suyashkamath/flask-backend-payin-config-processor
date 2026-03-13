
"""
Payin-Config — Payout Recalculator
===================================
Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

Usage:
    python recalculate_payout.py
"""

import pandas as pd
import math
import os
import json
import re
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
#  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# ─────────────────────────────────────────────────────────────────────────────

VEHICLE_TYPE_MASTER = {
    1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
    2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
    3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
    4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
    5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
    6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
    7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
    8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
    9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
    10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
    11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
    12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
    13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
    14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
    15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
    16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
    17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
    18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
    19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
    20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
    21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
    22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
    23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
    24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
    25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
    26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
    27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
    28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
    29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
    30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
    31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
    32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
}

# IDs considered as TAXI
TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# IDs considered as STAFF BUS
STAFF_BUS_IDS = {28}  # Staff Bus

# IDs considered as SCHOOL BUS
SCHOOL_BUS_IDS = {11}  # School Bus

# IDs considered as BUS (any bus — route/passenger)
ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# IDs considered as GCV 3-Wheeler goods
GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# IDs considered as Passenger 3-Wheeler (auto etc.)
PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# Insurers to match for Upto 2.5 GVW special rule
SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}



# ─


# ─────────────────────────────────────────────────────────────────────────────
#  FLOOR
# ─────────────────────────────────────────────────────────────────────────────

def floor_payout(value):
    return float(math.floor(float(value)))


# ─────────────────────────────────────────────────────────────────────────────
#  PO STRING PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_po_to_payout(po_str, p):
    po_str = str(po_str).strip().upper()

    if re.search(r'\d+%\s*OF PAYIN', po_str):
        percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
        return floor_payout(p * (percent / 100))

    if "PAYIN + 1" in po_str:
        return floor_payout(p + 1)

    if "LESS 2% OF PAYIN" in po_str:
        return floor_payout(p - 2)

    # "-3%", "-4%", "-5%"
    m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
    if m:
        ded = float(m.group(1))
        return floor_payout(p - ded)

    # "21% PO" — fixed payout
    m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
    if m:
        return floor_payout(float(m.group(1)))

    # Fallback
    return floor_payout(p * 0.90)


# ─────────────────────────────────────────────────────────────────────────────
#  RULE SELECTION — slab-based matching on REMARKS
# ─────────────────────────────────────────────────────────────────────────────

def select_po(rules_list, p):
    """
    Match payin p against slab REMARKS. Confirmed semantics:
      "Payin Below X%"  -> p <= X       (inclusive)
      "Payin A% to B%"  -> A <= p <= B  (inclusive both ends)
      "Payin Above X%"  -> p > X        (strict — p=50 stays in 31-50, p=51 hits Above 50)
      NIL / ALL FUEL / empty -> always match
    """
    if not rules_list:
        return None
    for r in rules_list:
        rem = str(r.get("REMARKS", "")).upper().strip()
        if not rem or rem == "NIL" or rem == "ALL FUEL":
            return r["PO"]
        m_below = re.search(r'BELOW\s+(\d+(?:\.\d+)?)%', rem)
        if m_below and p <= float(m_below.group(1)):
            return r["PO"]
        m_range = re.search(r'(\d+(?:\.\d+)?)%\s+TO\s+(\d+(?:\.\d+)?)%', rem)
        if m_range:
            lo, hi = float(m_range.group(1)), float(m_range.group(2))
            if lo <= p <= hi:          # inclusive both ends
                return r["PO"]
        m_above = re.search(r'ABOVE\s+(\d+(?:\.\d+)?)%', rem)
        if m_above and p > float(m_above.group(1)):    # strict
            return r["PO"]
    return rules_list[-1]["PO"]


def insurer_matches(rule_insurer_str, ins):
    """
    Check if the row insurer matches a rule INSURER field.
    Handles entries like "Tata- Comp" where the rule token contains
    the insurer name plus extra words/punctuation.

    Per comma-separated token in the rule:
      1. Exact match after normalisation.
      2. First word of token equals insurer -> "TATA- COMP" first word = "TATA".
      3. Insurer is a substring of the token as final fallback.
    """
    ri = str(rule_insurer_str).strip().upper()
    if ri == "ALL COMPANIES":
        return True

    ins_norm = ins.strip().upper()
    for token in [x.strip().upper() for x in ri.split(",")]:
        if ins_norm == token:
            return True
        token_words = re.sub(r'[^A-Z0-9 ]', ' ', token).split()
        if token_words and token_words[0] == ins_norm:
            return True
        if ins_norm in token:
            return True

    return False


def filter_by_insurer(rules, ins):
    """
    Return best matching rules for given insurer:
    1. Specific match (not 'All Companies', not 'Rest of Companies')
    2. 'All Companies'
    3. 'Rest of Companies'
    """
    specific = [r for r in rules
                if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
                and insurer_matches(r.get("INSURER",""), ins)]
    if specific:
        return specific

    all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
    if all_co:
        return all_co

    rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
    return rest


# ─────────────────────────────────────────────────────────────────────────────
#  DETERMINE JSON SEGMENT from row data
# ─────────────────────────────────────────────────────────────────────────────

def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
                              from_wt, to_wt, company_code, is_od):
    """
    Returns (lob, json_segment) tuple for rule lookup.
    lob is the JSON LOB string.
    json_segment is the JSON SEGMENT string.
    Returns (None, None) if not mappable.
    """
    sp = str(sub_product_name).strip()
    seg = str(segment).strip().upper()
    vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
    ins_upper = str(company_code).strip().upper()

    # ── TWO WHEELER ──────────────────────────────────────────────────────────
    if sp == "Two Wheeler":
        lob = "TW"
        if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
            return lob, "TW SAOD + COMP"
        else:  # TP Only
            return lob, "TW TP"

    # ── PRIVATE CAR ──────────────────────────────────────────────────────────
    if sp == "Private Car":
        lob = "PVT CAR"
        if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
            return lob, "PVT CAR COMP + SAOD"
        else:
            return lob, "PVT CAR TP"

    # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
    if sp == "Passenger Vehicle":
        # TAXI
        if vt_id in TAXI_VEHICLE_IDS:
            return "TAXI", "TAXI"

        # STAFF BUS
        if vt_id in STAFF_BUS_IDS:
            return "BUS", "STAFF BUS"

        # SCHOOL BUS
        if vt_id in SCHOOL_BUS_IDS:
            return "BUS", "SCHOOL BUS"

        # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
        if vt_id in ROUTE_BUS_IDS:
            return "BUS", "STAFF BUS"

        # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
        if vt_id in PCV_3W_IDS:
            return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

        # Tempo Traveller — treat as staff bus
        return "BUS", "STAFF BUS"

    # ── GOODS VEHICLE ────────────────────────────────────────────────────────
    if sp == "Goods Vehicle":
        from_w = float(from_wt) if pd.notna(from_wt) else 0
        to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

        # 3-Wheeler goods
        if vt_id in GCV_3W_IDS:
            return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

        # Upto 2.5T GVW + special insurers
        if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
            return "GCV, PCV 3W", "Upto 2.5 GVW"

        # Everything else (inc. upto 2.5T with other insurers)
        return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

    # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
    if sp == "Miscellaneous Vehicle":
        return "MISD", "Misd, Tractor"

    return None, None


# ─────────────────────────────────────────────────────────────────────────────
#  CORE FORMULA ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def compute_payout(payin, sub_product_name, segment, company_code,
                   vehicle_type_id, from_wt, to_wt,
                   rules, is_od=True):
    """
    Returns (payout_value, explanation_dict).
    explanation_dict has keys: lob, segment, insurer_matched, po_formula,
    remarks_slab, calculation_note
    """
    explanation = {
        "lob": "",
        "segment": "",
        "insurer_matched": "",
        "po_formula": "",
        "remarks_slab": "",
        "calculation_note": "",
    }

    try:
        p = float(payin)
    except (TypeError, ValueError):
        explanation["calculation_note"] = "Invalid payin value"
        return 0.0, explanation
    if p == 0:
        explanation["calculation_note"] = "Payin is 0 — no payout"
        return 0.0, explanation

    # ── RULE 1: Payin <= 5 → Payout is 0 ─────────────────────────────────────
    if p <= 5:
        explanation["calculation_note"] = f"Payin {p} <= 5 — payout is 0"
        explanation["po_formula"] = "Payin <= 5 → 0"
        return 0.0, explanation

    # ── RULE 2: IRDA rates (OD only) → Payout is 0 ───────────────────────────
    # Private Car OD payin = 15, Two Wheeler OD payin = 17.5 are IRDA fixed rates.
    # We do not process payout for IRDA rates — calculated_od_rate = 0.
    if is_od:
        sp_upper = str(sub_product_name).strip()
        if (sp_upper == "Private Car" and p == 15) or (sp_upper == "Two Wheeler" and p == 17.5):
            explanation["calculation_note"] = f"IRDA rate (payin={p}) — no payout processed"
            explanation["po_formula"] = "IRDA rate → 0"
            return 0.0, explanation

    ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
    lob, json_seg = get_json_lob_and_segment(
        sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
    )

    explanation["lob"]     = lob     if lob     else "NOT MAPPED"
    explanation["segment"] = json_seg if json_seg else "NOT MAPPED"

    if lob is None:
        p_out = floor_payout(p * 0.90)
        explanation["po_formula"]        = "90% of Payin (fallback)"
        explanation["insurer_matched"]   = "N/A"
        explanation["remarks_slab"]      = "N/A"
        explanation["calculation_note"]  = f"No LOB mapping found. Fallback: floor({p} × 0.90) = {p_out}"
    else:
        seg_rules      = [r for r in rules if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]
        candidate_rules = filter_by_insurer(seg_rules, ins)
        selected_po    = select_po(candidate_rules, p)

        if candidate_rules:
            explanation["insurer_matched"] = str(candidate_rules[0].get("INSURER", ""))
            # Find which slab was picked
            for r in candidate_rules:
                rem = str(r.get("REMARKS","")).upper().strip()
                if not rem or rem == "NIL" or rem == "ALL FUEL":
                    explanation["remarks_slab"] = r.get("REMARKS","NIL")
                    break
                m_below = re.search(r'BELOW\s+(\d+(?:\.\d+)?)%', rem)
                if m_below and p <= float(m_below.group(1)):
                    explanation["remarks_slab"] = r.get("REMARKS","")
                    break
                m_range = re.search(r'(\d+(?:\.\d+)?)%\s+TO\s+(\d+(?:\.\d+)?)%', rem)
                if m_range:
                    lo, hi = float(m_range.group(1)), float(m_range.group(2))
                    if lo <= p <= hi:          # inclusive both ends
                        explanation["remarks_slab"] = r.get("REMARKS","")
                        break
                m_above = re.search(r'ABOVE\s+(\d+(?:\.\d+)?)%', rem)
                if m_above and p > float(m_above.group(1)):
                    explanation["remarks_slab"] = r.get("REMARKS","")
                    break
        else:
            explanation["insurer_matched"] = "No matching rule"
            explanation["remarks_slab"]    = "N/A"

        if selected_po is None:
            p_out = floor_payout(p * 0.90)
            explanation["po_formula"]       = "90% of Payin (fallback — no rule matched)"
            explanation["calculation_note"] = f"Fallback: floor({p} × 0.90) = {p_out}"
        else:
            explanation["po_formula"] = selected_po
            p_out = parse_po_to_payout(selected_po, p)
            # Build human-readable calculation note
            po_up = str(selected_po).strip().upper()
            if re.search(r'\d+%\s*OF PAYIN', po_up):
                pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
                explanation["calculation_note"] = f"floor({p} × {pct}/100) = {p_out}"
            elif "PAYIN + 1" in po_up:
                explanation["calculation_note"] = f"floor({p} + 1) = {p_out}"
            elif "LESS 2% OF PAYIN" in po_up:
                explanation["calculation_note"] = f"floor({p} - 2) = {p_out}"
            elif re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up):
                ded = float(re.search(r'-(\d+(?:\.\d+)?)%', po_up).group(1))
                explanation["calculation_note"] = f"floor({p} - {ded}) = {p_out}"
            elif re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up):
                explanation["calculation_note"] = f"Fixed PO = {p_out}"
            else:
                explanation["calculation_note"] = f"= {p_out}"

    result = max(0.0, p_out)
    return result, explanation


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*70)
    print("  Payin-Config — Payout Recalculator")
    print("="*70)

    json_path   = input("\nEnter path to payout_rules.json : ").strip().strip('"')
    input_path  = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
    output_path = input("Enter output file path (blank=auto): ").strip().strip('"')

    try:
        with open(json_path) as f:
            rules = json.load(f)
        print(f"  Loaded {len(rules)} rules from {json_path}")
    except Exception as e:
        print(f"\n[ERROR] Failed to load JSON: {e}"); return

    if not output_path:
        base, ext = os.path.splitext(input_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"{base}_recalculated_{ts}{ext}"

    print(f"\n  Reading: {input_path}")
    df = pd.read_excel(input_path)
    df.columns = [c.strip() for c in df.columns]
    total = len(df)
    print(f"  Rows   : {total}")

    # Capture original payout values before any changes
    original_od = df['payout_od_rate'].tolist()
    original_tp = df['payout_tp_rate'].tolist()

    required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
                'sub_product_name','segment','company_code',
                'vehicle_type_id','from_weightage_kg','to_weightage_kg']
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"\n[ERROR] Missing columns: {missing}"); return

    new_od, new_tp = [], []
    changed_od = changed_tp = processed_od = processed_tp = 0

    # Explanation column lists
    od_lob, od_seg, od_ins, od_po, od_slab, od_note = [], [], [], [], [], []
    tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_note = [], [], [], [], [], []

    for _, row in df.iterrows():
        sp   = row['sub_product_name']
        seg  = row['segment']
        ins  = row['company_code']
        vt   = row['vehicle_type_id']
        f_wt = row['from_weightage_kg']
        t_wt = row['to_weightage_kg']
        # OD
        payin_od = row['payin_od_rate']
        old_od   = row['payout_od_rate']
        if pd.notna(payin_od) and float(payin_od) != 0:
            processed_od += 1
            calc_od, expl_od = compute_payout(payin_od, sp, seg, ins, vt, f_wt, t_wt, rules, is_od=True)
            new_od.append(calc_od)
            if abs(float(old_od) - calc_od) > 0.001:
                changed_od += 1
        else:
            calc_od = 0.0 if pd.isna(old_od) else old_od
            new_od.append(calc_od)
            expl_od = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
                       "remarks_slab":"","calculation_note":"Payin is 0"}
        od_lob.append(expl_od["lob"])
        od_seg.append(expl_od["segment"])
        od_ins.append(expl_od["insurer_matched"])
        od_po.append(expl_od["po_formula"])
        od_slab.append(expl_od["remarks_slab"])
        od_note.append(expl_od["calculation_note"])

        # TP
        payin_tp = row['payin_tp_rate']
        old_tp   = row['payout_tp_rate']
        if pd.notna(payin_tp) and float(payin_tp) != 0:
            processed_tp += 1
            calc_tp, expl_tp = compute_payout(payin_tp, sp, seg, ins, vt, f_wt, t_wt, rules, is_od=False)
            new_tp.append(calc_tp)
            if abs(float(old_tp) - calc_tp) > 0.001:
                changed_tp += 1
        else:
            calc_tp = 0.0 if pd.isna(old_tp) else old_tp
            new_tp.append(calc_tp)
            expl_tp = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
                       "remarks_slab":"","calculation_note":"Payin is 0"}
        tp_lob.append(expl_tp["lob"])
        tp_seg.append(expl_tp["segment"])
        tp_ins.append(expl_tp["insurer_matched"])
        tp_po.append(expl_tp["po_formula"])
        tp_slab.append(expl_tp["remarks_slab"])
        tp_note.append(expl_tp["calculation_note"])

    # Keep original payout_od_rate and payout_tp_rate UNCHANGED
    # Add calculated rates and changed flag as new columns
    df['calculated_od_rate'] = new_od
    df['calculated_tp_rate'] = new_tp
    def _safe_changed(o_od, c_od, o_tp, c_tp):
        try:
            od_changed = abs(float(o_od if pd.notna(o_od) else 0) - float(c_od)) > 0.001
        except: od_changed = False
        try:
            tp_changed = abs(float(o_tp if pd.notna(o_tp) else 0) - float(c_tp)) > 0.001
        except: tp_changed = False
        return od_changed or tp_changed

    df['changed_payout'] = [
        _safe_changed(o_od, c_od, o_tp, c_tp)
        for o_od, c_od, o_tp, c_tp in zip(original_od, new_od, original_tp, new_tp)
    ]

    # Append OD explanation columns
    df['od_rule_lob']            = od_lob
    df['od_rule_segment']        = od_seg
    df['od_rule_insurer']        = od_ins
    df['od_rule_po_formula']     = od_po
    df['od_rule_slab']           = od_slab
    df['od_rule_calculation']    = od_note

    # Append TP explanation columns
    df['tp_rule_lob']            = tp_lob
    df['tp_rule_segment']        = tp_seg
    df['tp_rule_insurer']        = tp_ins
    df['tp_rule_po_formula']     = tp_po
    df['tp_rule_slab']           = tp_slab
    df['tp_rule_calculation']    = tp_note

    # Reorder columns: insert calculated_od_rate, calculated_tp_rate, changed_payout
    # right after payout_tp_rate
    all_cols = list(df.columns)
    new_cols = ['calculated_od_rate', 'calculated_tp_rate', 'changed_payout']
    # Remove new cols from wherever they currently are
    base_cols = [c for c in all_cols if c not in new_cols]
    # Find position right after payout_tp_rate
    try:
        insert_pos = base_cols.index('payout_tp_rate') + 1
    except ValueError:
        insert_pos = len(base_cols)
    final_cols = base_cols[:insert_pos] + new_cols + base_cols[insert_pos:]
    df = df[final_cols]

    df.to_excel(output_path, index=False)

    print(f"\n{'='*70}")
    print(f"  COMPLETED")
    print(f"  Total rows           : {total}")
    print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
    print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
    print(f"  Output saved to      : {output_path}")
    print(f"{'='*70}")

    # Sample preview
    sample_cols = ['sub_product_name','segment','company_code',
                   'payin_od_rate','payout_od_rate','calculated_od_rate',
                   'payin_tp_rate','payout_tp_rate','calculated_tp_rate','changed_payout']
    sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
    print("\n  Sample output (first 15 non-zero OD rows):\n")
    print(sample.to_string(index=False))
    print()


if __name__ == "__main__":
    main()
