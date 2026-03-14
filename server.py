# """
# Payout Recalculator — Flask Server
# ====================================
# Serves the UI and processes PayinConfig uploads against JSON payout rules.
# """

# import os
# import io
# import json
# import tempfile
# from datetime import datetime
# from pathlib import Path

# import pandas as pd
# from flask import Flask, render_template, request, jsonify, send_file
# from werkzeug.utils import secure_filename

# # Import all logic from recalculate_payout.py (same folder)
# import recalculate_payout as calc

# app = Flask(__name__)
# app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB

# UPLOAD_FOLDER = Path('uploads')
# OUTPUT_FOLDER = Path('outputs')
# UPLOAD_FOLDER.mkdir(exist_ok=True)
# OUTPUT_FOLDER.mkdir(exist_ok=True)

# ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


# def allowed_file(filename):
#     return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# # ============================================================================
# # INDEX
# # ============================================================================

# @app.route('/')
# def index():
#     return render_template('index.html')


# # ============================================================================
# # API: PROCESS — receives files + rules JSON, returns processed Excel
# # ============================================================================

# @app.route('/process', methods=['POST'])
# def process():
#     try:
#         # ── 1. Validate inputs ────────────────────────────────────────────────
#         payin_file   = request.files.get('payin_file')
#         rules_json   = request.form.get('rules_json', '').strip()

#         payin_min    = float(request.form.get('payin_min', 5))   # payin <= this → 0
#         irda_pvt_car = float(request.form.get('irda_pvt_car', 15))
#         irda_tw      = float(request.form.get('irda_tw', 17.5))

#         # Segment filters (empty = apply to all segments)
#         def _parse_segs(key):
#             raw = request.form.get(key, '').strip()
#             if not raw or raw == 'ALL': return []
#             return [s.strip().upper() for s in raw.split(',') if s.strip()]
#         thresh_segs   = _parse_segs('thresh_segs')
#         irda_pvt_segs = _parse_segs('irda_pvt_segs')
#         irda_tw_segs  = _parse_segs('irda_tw_segs')
#         print(f"[/process] thresh_segs={thresh_segs or 'all'}, irda_pvt_segs={irda_pvt_segs or 'all'}, irda_tw_segs={irda_tw_segs or 'all'}")

#         if not payin_file or payin_file.filename == '':
#             return jsonify({'success': False, 'error': 'PayinConfig file is required'}), 400
#         if not rules_json:
#             return jsonify({'success': False, 'error': 'Payout rules JSON is required'}), 400

#         # ── 2. Parse rules JSON ───────────────────────────────────────────────
#         try:
#             rules = json.loads(rules_json)
#         except json.JSONDecodeError as e:
#             return jsonify({'success': False, 'error': f'Invalid JSON: {e}'}), 400

#         # ── 3. Save uploads to temp files ─────────────────────────────────────
#         with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_payin:
#             payin_file.save(tmp_payin.name)
#             payin_path = tmp_payin.name

#         # ── 5. Load PayinConfig — ALL rows, no filtering ─────────────────────
#         df = pd.read_excel(payin_path)
#         df.columns = [c.strip() for c in df.columns]
#         total = len(df)
#         print(f"[/process] Total rows loaded: {total} — no state row filtering applied")

#         required_cols = [
#             'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
#             'sub_product_name', 'segment', 'company_code',
#             'vehicle_type_id', 'from_weightage_kg', 'to_weightage_kg'
#         ]
#         missing = [c for c in required_cols if c not in df.columns]
#         if missing:
#             return jsonify({'success': False, 'error': f'Missing columns: {missing}'}), 400

#         # ── 7. Monkey-patch IRDA thresholds if user customised them ──────────
#         # We temporarily override the module-level check by wrapping compute_payout
#         orig_compute = calc.compute_payout

#         def patched_compute(payin, sub_product_name, segment, company_code,
#                             vehicle_type_id, from_wt, to_wt,
#                             rules_arg, is_od=True):
#             explanation = {
#                 "lob": "", "segment": "", "insurer_matched": "",
#                 "po_formula": "", "remarks_slab": "",
#                 "calculation_note": "",
#             }
#             try:
#                 p = float(payin)
#             except (TypeError, ValueError):
#                 explanation["calculation_note"] = "Invalid payin value"
#                 return 0.0, explanation

#             if p == 0:
#                 explanation["calculation_note"] = "Payin is 0 — no payout"
#                 return 0.0, explanation

#             # Map sub_product_name → chip label for segment filter matching
#             SP_TO_CHIP = {
#                 'private car':   'PVT CAR',
#                 'two wheeler':   'TW',
#                 'goods vehicle': 'CV / GCV',
#                 'passenger':     'PCV',
#                 'taxi':          'TAXI',
#                 'bus':           'BUS',
#                 'misd':          'MISD',
#                 'tractor':       'MISD',
#             }
#             sp_lower = str(sub_product_name).strip().lower()
#             sp_chip  = SP_TO_CHIP.get(sp_lower, sp_lower.upper())

#             def seg_match(filters):
#                 """True if no filter set, or sp_chip matches any selected label."""
#                 if not filters:
#                     return True
#                 return any(f.upper() in sp_chip.upper() or sp_chip.upper() in f.upper()
#                            for f in filters)

#             # Rule 1: payin <= payin_min → 0  (check segment filter)
#             if p <= payin_min and seg_match(thresh_segs):
#                 explanation["calculation_note"] = f"Payin {p} <= {payin_min} — payout is 0"
#                 explanation["po_formula"] = f"Payin <= {payin_min} → 0"
#                 return 0.0, explanation

#             # Rule 2: IRDA rates (OD only) → 0  (check segment filter per rate)
#             if is_od:
#                 sp = str(sub_product_name).strip()
#                 if sp == "Private Car" and p == irda_pvt_car and seg_match(irda_pvt_segs):
#                     explanation["calculation_note"] = f"IRDA rate (payin={p}) — no payout processed"
#                     explanation["po_formula"] = "IRDA rate → 0"
#                     return 0.0, explanation
#                 if sp == "Two Wheeler" and p == irda_tw and seg_match(irda_tw_segs):
#                     explanation["calculation_note"] = f"IRDA rate (payin={p}) — no payout processed"
#                     explanation["po_formula"] = "IRDA rate → 0"
#                     return 0.0, explanation

#             return orig_compute(payin, sub_product_name, segment, company_code,
#                                 vehicle_type_id, from_wt, to_wt,
#                                 rules_arg, is_od=is_od)

#         # ── 7. Process rows ───────────────────────────────────────────────────
#         original_od = df['payout_od_rate'].tolist()
#         original_tp = df['payout_tp_rate'].tolist()

#         new_od, new_tp = [], []
#         changed_od = changed_tp = processed_od = processed_tp = 0

#         od_lob, od_seg, od_ins, od_po, od_slab, od_note = [], [], [], [], [], []
#         tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_note = [], [], [], [], [], []

#         blank = {
#             'lob': '', 'segment': '', 'insurer_matched': '', 'po_formula': '',
#             'remarks_slab': '', 'calculation_note': 'Payin is 0'
#         }

#         for _, row in df.iterrows():
#             sp         = row['sub_product_name']
#             seg        = row['segment']
#             ins        = row['company_code']
#             vt         = row['vehicle_type_id']
#             f_wt       = row['from_weightage_kg']
#             t_wt       = row['to_weightage_kg']
#             # OD
#             payin_od = row['payin_od_rate']
#             old_od   = row['payout_od_rate']
#             if pd.notna(payin_od) and float(payin_od) != 0:
#                 processed_od += 1
#                 calc_od, expl_od = patched_compute(
#                     payin_od, sp, seg, ins, vt, f_wt, t_wt,
#                     rules, is_od=True
#                 )
#                 new_od.append(calc_od)
#                 if abs(float(old_od) - calc_od) > 0.001:
#                     changed_od += 1
#             else:
#                 new_od.append(0.0 if pd.isna(old_od) else old_od)
#                 expl_od = blank
#             od_lob.append(expl_od['lob']); od_seg.append(expl_od['segment'])
#             od_ins.append(expl_od['insurer_matched']); od_po.append(expl_od['po_formula'])
#             od_slab.append(expl_od['remarks_slab'])
#             od_note.append(expl_od['calculation_note'])

#             # TP
#             payin_tp = row['payin_tp_rate']
#             old_tp   = row['payout_tp_rate']
#             if pd.notna(payin_tp) and float(payin_tp) != 0:
#                 processed_tp += 1
#                 calc_tp, expl_tp = patched_compute(
#                     payin_tp, sp, seg, ins, vt, f_wt, t_wt,
#                     rules, is_od=False
#                 )
#                 new_tp.append(calc_tp)
#                 if abs(float(old_tp) - calc_tp) > 0.001:
#                     changed_tp += 1
#             else:
#                 new_tp.append(0.0 if pd.isna(old_tp) else old_tp)
#                 expl_tp = blank
#             tp_lob.append(expl_tp['lob']); tp_seg.append(expl_tp['segment'])
#             tp_ins.append(expl_tp['insurer_matched']); tp_po.append(expl_tp['po_formula'])
#             tp_slab.append(expl_tp['remarks_slab'])
#             tp_note.append(expl_tp['calculation_note'])

#         # ── 8. Build output DataFrame ─────────────────────────────────────────
#         df['calculated_od_rate'] = new_od
#         df['calculated_tp_rate'] = new_tp

#         def _safe_changed(o_od, c_od, o_tp, c_tp):
#             try:
#                 a = abs(float(o_od if pd.notna(o_od) else 0) - float(c_od)) > 0.001
#             except: a = False
#             try:
#                 b = abs(float(o_tp if pd.notna(o_tp) else 0) - float(c_tp)) > 0.001
#             except: b = False
#             return a or b

#         df['changed_payout'] = [
#             _safe_changed(o_od, c_od, o_tp, c_tp)
#             for o_od, c_od, o_tp, c_tp in zip(original_od, new_od, original_tp, new_tp)
#         ]

#         df['od_rule_lob']         = od_lob
#         df['od_rule_segment']     = od_seg
#         df['od_rule_insurer']     = od_ins
#         df['od_rule_po_formula']  = od_po
#         df['od_rule_slab']        = od_slab
#         df['od_rule_calculation'] = od_note
#         df['tp_rule_lob']         = tp_lob
#         df['tp_rule_segment']     = tp_seg
#         df['tp_rule_insurer']     = tp_ins
#         df['tp_rule_po_formula']  = tp_po
#         df['tp_rule_slab']        = tp_slab
#         df['tp_rule_calculation'] = tp_note

#         # Reorder: insert new cols right after payout_tp_rate
#         new_cols  = ['calculated_od_rate', 'calculated_tp_rate', 'changed_payout']
#         base_cols = [c for c in df.columns if c not in new_cols]
#         try:
#             insert_pos = base_cols.index('payout_tp_rate') + 1
#         except ValueError:
#             insert_pos = len(base_cols)
#         final_cols = base_cols[:insert_pos] + new_cols + base_cols[insert_pos:]
#         df = df[final_cols]

#         # ── 9. Write output Excel to memory buffer ────────────────────────────
#         # Replace NaN/inf with empty string so blank cells stay blank (not #NUM!)
#         df = df.where(pd.notnull(df), '')

#         buf = io.BytesIO()
#         with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
#             df.to_excel(writer, sheet_name='Recalculated', index=False)
#             wb  = writer.book
#             ws  = writer.sheets['Recalculated']

#             hdr_fmt = wb.add_format({
#                 'bold': True, 'bg_color': '#FFC000', 'font_color': '#000000',
#                 'border': 1, 'align': 'center', 'valign': 'vcenter',
#                 'font_name': 'Arial', 'font_size': 9,
#             })
#             data_fmt = wb.add_format({
#                 'border': 1, 'align': 'center', 'valign': 'vcenter',
#                 'font_name': 'Arial', 'font_size': 9,
#             })
#             changed_fmt = wb.add_format({
#                 'border': 1, 'align': 'center', 'valign': 'vcenter',
#                 'font_name': 'Arial', 'font_size': 9,
#                 'bg_color': '#FFE0E0',
#             })

#             for col_idx, col_name in enumerate(df.columns):
#                 ws.write(0, col_idx, col_name, hdr_fmt)

#             changed_col_idx = list(df.columns).index('changed_payout')
#             for row_idx, row_data in enumerate(df.itertuples(index=False), start=1):
#                 is_changed = row_data[changed_col_idx]
#                 fmt = changed_fmt if is_changed else data_fmt
#                 for col_idx, val in enumerate(row_data):
#                     ws.write(row_idx, col_idx, val, fmt)

#             for col_idx, col_name in enumerate(df.columns):
#                 max_len = max(
#                     len(str(col_name)),
#                     df.iloc[:, col_idx].astype(str).str.len().max() if len(df) > 0 else 0
#                 )
#                 ws.set_column(col_idx, col_idx, min(max_len + 3, 35))

#             ws.freeze_panes(1, 0)

#         buf.seek(0)

#         # ── 10. Clean up temp files ───────────────────────────────────────────
#         os.unlink(payin_path)
#         # ── 11. Build summary stats ───────────────────────────────────────────
#         changed_total = int(df['changed_payout'].sum())
#         stats = {
#             'total_rows':     total,
#             'od_processed':   processed_od,
#             'od_changed':     changed_od,
#             'tp_processed':   processed_tp,
#             'tp_changed':     changed_tp,
#             'changed_payout': changed_total,
#         }

#         # Generate output filename
#         base_name = secure_filename(payin_file.filename).rsplit('.', 1)[0]
#         ts = datetime.now().strftime('%Y%m%d_%H%M%S')
#         out_name = f'{base_name}_recalculated_{ts}.xlsx'

#         import urllib.parse
#         response = send_file(
#             buf,
#             as_attachment=True,
#             download_name=out_name,
#             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#         )
#         response.headers['X-Stats'] = urllib.parse.quote(json.dumps(stats))
#         response.headers['Access-Control-Expose-Headers'] = 'X-Stats'
#         return response

#     except Exception as e:
#         import traceback
#         traceback.print_exc()
#         return jsonify({'success': False, 'error': str(e)}), 500


# # ============================================================================
# # API: VALIDATE JSON — quick check that rules JSON is parseable
# # ============================================================================

# @app.route('/validate-rules', methods=['POST'])
# def validate_rules():
#     try:
#         data = request.get_json(force=True)
#         raw  = data.get('rules_json', '')
#         rules = json.loads(raw)
#         lobs = sorted(set(r.get('LOB', '') for r in rules if r.get('LOB')))
#         return jsonify({
#             'valid': True,
#             'total_rules': len(rules),
#             'lobs': lobs,
#         })
#     except json.JSONDecodeError as e:
#         return jsonify({'valid': False, 'error': str(e)})
#     except Exception as e:
#         return jsonify({'valid': False, 'error': str(e)})


# if __name__ == '__main__':
#     app.run(debug=True, port=5001)


# version 2 by suyash on saturday  

"""
Payout Recalculator — Flask Server
====================================
Serves the UI and processes PayinConfig uploads against JSON payout rules.
Supports state-wise payout overrides via rto_group_id matching.
"""

import os
import io
import json
import tempfile
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

import recalculate_payout as calc

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB

UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('outputs')
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# ── Load RTO master once at startup ──────────────────────────────────────────
RTO_MASTER_PATH = Path('distinct_rto_output.json')
_RTO_MASTER = []
if RTO_MASTER_PATH.exists():
    try:
        with open(RTO_MASTER_PATH) as f:
            _RTO_MASTER = json.load(f)
        print(f"[startup] Loaded {len(_RTO_MASTER)} RTO records from {RTO_MASTER_PATH}")
    except Exception as e:
        print(f"[startup] WARNING: Could not load RTO master: {e}")
else:
    print(f"[startup] WARNING: {RTO_MASTER_PATH} not found — state overrides won't resolve company_id lookups")


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_rto_group_ids(rto_code, company_id):
    """
    Return list of rto_group_ids matching rto_code + company_id from master.
    company_id can be int or str. rto_code is case-insensitive.
    """
    rto_upper = str(rto_code).strip().upper()
    try:
        cid = int(company_id)
    except (TypeError, ValueError):
        return []
    return [
        r['rto_group_id']
        for r in _RTO_MASTER
        if str(r.get('rto_code', '')).strip().upper() == rto_upper
        and r.get('company_id') == cid
    ]


def apply_state_override_po(po_str, p):
    """Apply a PO formula string to payin value p. Returns float."""
    import math
    po_up = str(po_str).strip().upper()

    if re.search(r'\d+%\s*OF PAYIN', po_up):
        pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
        return float(math.floor(p * pct / 100))

    if 'PAYIN + 1' in po_up:
        return float(math.floor(p + 1))

    if 'LESS 2% OF PAYIN' in po_up:
        return float(math.floor(p - 2))

    m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up)
    if m:
        return float(math.floor(p - float(m.group(1))))

    m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up)
    if m:
        return float(math.floor(float(m.group(1))))

    return float(math.floor(p * 0.90))


def payin_in_range(p, remarks):
    """Check if payin p falls within the REMARKS range string."""
    rem = str(remarks).strip().upper()
    if not rem or rem in ('NIL', 'ALL FUEL'):
        return True
    m = re.search(r'BELOW\s+(\d+(?:\.\d+)?)%', rem)
    if m and p <= float(m.group(1)):
        return True
    m = re.search(r'(\d+(?:\.\d+)?)%\s+TO\s+(\d+(?:\.\d+)?)%', rem)
    if m and float(m.group(1)) <= p <= float(m.group(2)):
        return True
    m = re.search(r'ABOVE\s+(\d+(?:\.\d+)?)%', rem)
    if m and p > float(m.group(1)):
        return True
    return False


# Segment chips → sub_product_name mapping
CHIP_TO_SP = {
    'PVT CAR':  ['Private Car'],
    'TW':       ['Two Wheeler'],
    'CV / GCV': ['Goods Vehicle'],
    'PCV':      ['Passenger Vehicle'],
    'TAXI':     ['Passenger Vehicle'],   # vehicle_type_id filter handled separately
    'BUS':      ['Passenger Vehicle'],
    'MISD':     ['Miscellaneous Vehicle'],
}

ALL_SEGMENTS = [
    'TW SAOD + COMP', 'TW TP',
    'PVT CAR COMP + SAOD', 'PVT CAR TP',
    'All GVW & PCV 3W, GCV 3W', 'Upto 2.5 GVW',
    'SCHOOL BUS', 'STAFF BUS', 'TAXI',
    'Misd, Tractor',
]


def insurer_matches_override(rule_insurer, row_insurer):
    """Match insurer field in override rule against row company_code."""
    ri = str(rule_insurer).strip().upper()
    if ri == 'ALL COMPANIES':
        return True
    row_ins = str(row_insurer).strip().upper()
    if ri == 'REST OF COMPANIES':
        # We treat rest-of-companies as catch-all (no specific match set)
        return True
    # Comma-separated tokens
    for token in [x.strip().upper() for x in ri.split(',')]:
        if row_ins == token:
            return True
        token_words = re.sub(r'[^A-Z0-9 ]', ' ', token).split()
        if token_words and token_words[0] == row_ins:
            return True
        if row_ins in token:
            return True
    return False


def resolve_override_insurer(override):
    """
    Always company-specific: look up rto_group_ids by (state_code, company_id).
    Returns set of rto_group_ids to match against rows.
    """
    rto_code   = override.get('state_code', '')
    company_id = override.get('company_id', '')
    group_ids  = get_rto_group_ids(rto_code, company_id)
    return 'custom', set(group_ids)


# ============================================================================
# INDEX
# ============================================================================

@app.route('/')
def index():
    return render_template('index.html')


# ============================================================================
# API: RTO LOOKUP — resolve rto_group_ids for a (state, company_id) pair
# ============================================================================

@app.route('/rto-lookup', methods=['POST'])
def rto_lookup():
    try:
        data = request.get_json(force=True)
        rto_code = data.get('rto_code', '').strip().upper()
        company_id = data.get('company_id', '')
        if not rto_code:
            return jsonify({'success': False, 'error': 'rto_code required'}), 400
        group_ids = get_rto_group_ids(rto_code, company_id)
        return jsonify({
            'success': True,
            'rto_code': rto_code,
            'company_id': company_id,
            'group_ids': group_ids,
            'count': len(group_ids),
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ============================================================================
# API: PROCESS
# ============================================================================

@app.route('/process', methods=['POST'])
def process():
    try:
        # ── 1. Validate inputs ────────────────────────────────────────────────
        payin_file   = request.files.get('payin_file')
        rules_json   = request.form.get('rules_json', '').strip()

        payin_min    = float(request.form.get('payin_min', 5))
        irda_pvt_car = float(request.form.get('irda_pvt_car', 15))
        irda_tw      = float(request.form.get('irda_tw', 17.5))

        # State overrides JSON
        state_overrides_raw = request.form.get('state_overrides_json', '[]').strip()

        def _parse_segs(key):
            raw = request.form.get(key, '').strip()
            if not raw or raw == 'ALL': return []
            return [s.strip().upper() for s in raw.split(',') if s.strip()]
        thresh_segs   = _parse_segs('thresh_segs')
        irda_pvt_segs = _parse_segs('irda_pvt_segs')
        irda_tw_segs  = _parse_segs('irda_tw_segs')

        if not payin_file or payin_file.filename == '':
            return jsonify({'success': False, 'error': 'PayinConfig file is required'}), 400
        if not rules_json:
            return jsonify({'success': False, 'error': 'Payout rules JSON is required'}), 400

        # ── 2. Parse rules & state overrides ─────────────────────────────────
        try:
            rules = json.loads(rules_json)
        except json.JSONDecodeError as e:
            return jsonify({'success': False, 'error': f'Invalid JSON: {e}'}), 400

        try:
            state_overrides = json.loads(state_overrides_raw)
        except json.JSONDecodeError:
            state_overrides = []

        # Pre-resolve insurer group_ids for each override
        resolved_overrides = []
        for ov in state_overrides:
            ins_type, group_ids = resolve_override_insurer(ov)
            resolved_overrides.append({**ov, '_ins_type': ins_type, '_group_ids': group_ids})

        # ── 3. Save uploads ───────────────────────────────────────────────────
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_payin:
            payin_file.save(tmp_payin.name)
            payin_path = tmp_payin.name

        # ── 4. Load PayinConfig ───────────────────────────────────────────────
        df = pd.read_excel(payin_path)
        df.columns = [c.strip() for c in df.columns]
        total = len(df)
        print(f"[/process] Total rows: {total}")

        required_cols = [
            'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
            'sub_product_name', 'segment', 'company_code',
            'vehicle_type_id', 'from_weightage_kg', 'to_weightage_kg'
        ]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return jsonify({'success': False, 'error': f'Missing columns: {missing}'}), 400

        # Check if rto_group_id column exists (needed for state overrides)
        has_rto_col = 'rto_group_id' in df.columns
        if state_overrides and not has_rto_col:
            print("[/process] WARNING: state_overrides provided but 'rto_group_id' column not found in input")

        # ── 5. Monkey-patch with IRDA/threshold overrides ────────────────────
        orig_compute = calc.compute_payout

        def patched_compute(payin, sub_product_name, segment, company_code,
                            vehicle_type_id, from_wt, to_wt, rules_arg, is_od=True):
            explanation = {
                "lob": "", "segment": "", "insurer_matched": "",
                "po_formula": "", "remarks_slab": "", "calculation_note": "",
            }
            try:
                p = float(payin)
            except (TypeError, ValueError):
                explanation["calculation_note"] = "Invalid payin value"
                return 0.0, explanation

            if p == 0:
                explanation["calculation_note"] = "Payin is 0 — no payout"
                return 0.0, explanation

            SP_TO_CHIP = {
                'private car': 'PVT CAR', 'two wheeler': 'TW',
                'goods vehicle': 'CV / GCV', 'passenger': 'PCV',
                'taxi': 'TAXI', 'bus': 'BUS', 'misd': 'MISD',
                'tractor': 'MISD',
            }
            sp_lower = str(sub_product_name).strip().lower()
            sp_chip  = SP_TO_CHIP.get(sp_lower, sp_lower.upper())

            def seg_match(filters):
                if not filters: return True
                return any(f.upper() in sp_chip.upper() or sp_chip.upper() in f.upper()
                           for f in filters)

            if p <= payin_min and seg_match(thresh_segs):
                explanation["calculation_note"] = f"Payin {p} <= {payin_min} — payout is 0"
                explanation["po_formula"] = f"Payin <= {payin_min} → 0"
                return 0.0, explanation

            if is_od:
                sp = str(sub_product_name).strip()
                if sp == "Private Car" and p == irda_pvt_car and seg_match(irda_pvt_segs):
                    explanation["calculation_note"] = f"IRDA rate (payin={p}) — no payout processed"
                    explanation["po_formula"] = "IRDA rate → 0"
                    return 0.0, explanation
                if sp == "Two Wheeler" and p == irda_tw and seg_match(irda_tw_segs):
                    explanation["calculation_note"] = f"IRDA rate (payin={p}) — no payout processed"
                    explanation["po_formula"] = "IRDA rate → 0"
                    return 0.0, explanation

            return orig_compute(payin, sub_product_name, segment, company_code,
                                vehicle_type_id, from_wt, to_wt, rules_arg, is_od=is_od)

        # ── 6. Process rows ───────────────────────────────────────────────────
        original_od = df['payout_od_rate'].tolist()
        original_tp = df['payout_tp_rate'].tolist()

        new_od, new_tp = [], []
        changed_od = changed_tp = processed_od = processed_tp = 0

        od_lob, od_seg, od_ins, od_po, od_slab, od_note = [], [], [], [], [], []
        tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_note = [], [], [], [], [], []

        # State override tracking
        state_override_applied = []
        state_override_notes   = []
        state_override_states  = []
        state_override_po_col  = []

        blank = {
            'lob': '', 'segment': '', 'insurer_matched': '', 'po_formula': '',
            'remarks_slab': '', 'calculation_note': 'Payin is 0'
        }

        for _, row in df.iterrows():
            sp   = row['sub_product_name']
            seg  = row['segment']
            ins  = row['company_code']
            vt   = row['vehicle_type_id']
            f_wt = row['from_weightage_kg']
            t_wt = row['to_weightage_kg']
            row_rto_gid = row.get('rto_group_id', None) if has_rto_col else None

            # --- OD ---
            payin_od = row['payin_od_rate']
            old_od   = row['payout_od_rate']
            if pd.notna(payin_od) and float(payin_od) != 0:
                processed_od += 1
                calc_od, expl_od = patched_compute(
                    payin_od, sp, seg, ins, vt, f_wt, t_wt, rules, is_od=True
                )
                new_od.append(calc_od)
                if abs(float(old_od if pd.notna(old_od) else 0) - calc_od) > 0.001:
                    changed_od += 1
            else:
                new_od.append(0.0 if pd.isna(old_od) else old_od)
                expl_od = blank
            od_lob.append(expl_od['lob']); od_seg.append(expl_od['segment'])
            od_ins.append(expl_od['insurer_matched']); od_po.append(expl_od['po_formula'])
            od_slab.append(expl_od['remarks_slab']); od_note.append(expl_od['calculation_note'])

            # --- TP ---
            payin_tp = row['payin_tp_rate']
            old_tp   = row['payout_tp_rate']
            if pd.notna(payin_tp) and float(payin_tp) != 0:
                processed_tp += 1
                calc_tp, expl_tp = patched_compute(
                    payin_tp, sp, seg, ins, vt, f_wt, t_wt, rules, is_od=False
                )
                new_tp.append(calc_tp)
                if abs(float(old_tp if pd.notna(old_tp) else 0) - calc_tp) > 0.001:
                    changed_tp += 1
            else:
                new_tp.append(0.0 if pd.isna(old_tp) else old_tp)
                expl_tp = blank
            tp_lob.append(expl_tp['lob']); tp_seg.append(expl_tp['segment'])
            tp_ins.append(expl_tp['insurer_matched']); tp_po.append(expl_tp['po_formula'])
            tp_slab.append(expl_tp['remarks_slab']); tp_note.append(expl_tp['calculation_note'])

            # --- State Override ---
            override_applied = False
            override_notes_list = []
            override_state_list = []
            override_po_list    = []

            if resolved_overrides and has_rto_col and pd.notna(row_rto_gid):
                row_gid = int(row_rto_gid)
                sp_str  = str(sp).strip()
                ins_str = str(ins).strip().upper()

                for ov in resolved_overrides:
                    state_code  = ov.get('state_code', '')
                    ov_segment  = ov.get('segment', 'ALL')   # 'ALL' or specific JSON segment
                    group_ids   = ov['_group_ids']  # set of rto_group_ids for this state+company
                    po_formula  = ov.get('po_formula', '')
                    remarks     = ov.get('remarks', 'NIL')
                    apply_to    = ov.get('apply_to', 'BOTH')  # 'OD', 'TP', 'BOTH'
                    company_name = ov.get('company_name', str(ov.get('company_id', '')))

                    # Only apply if this row's rto_group_id is in the resolved set
                    if not group_ids or row_gid not in group_ids:
                        continue

                    # Check segment match
                    if ov_segment != 'ALL':
                        # Map the JSON segment back to sub_product_name
                        seg_sp_map = {
                            'TW SAOD + COMP': 'Two Wheeler',
                            'TW TP':          'Two Wheeler',
                            'PVT CAR COMP + SAOD': 'Private Car',
                            'PVT CAR TP':     'Private Car',
                            'All GVW & PCV 3W, GCV 3W': 'Goods Vehicle',
                            'Upto 2.5 GVW':   'Goods Vehicle',
                            'SCHOOL BUS':     'Passenger Vehicle',
                            'STAFF BUS':      'Passenger Vehicle',
                            'TAXI':           'Passenger Vehicle',
                            'Misd, Tractor':  'Miscellaneous Vehicle',
                        }
                        expected_sp = seg_sp_map.get(ov_segment, '')
                        if expected_sp and sp_str != expected_sp:
                            continue

                    # group_id filter already done above — no further insurer check needed

                    # Check payin range
                    od_ok = (apply_to in ('OD', 'BOTH')) and pd.notna(payin_od) and float(payin_od) != 0
                    tp_ok = (apply_to in ('TP', 'BOTH')) and pd.notna(payin_tp) and float(payin_tp) != 0

                    applied_this = False
                    # Threshold dominates — never override a row that was zeroed by threshold
                    od_above_thresh = od_ok and float(payin_od) > payin_min
                    tp_above_thresh = tp_ok and float(payin_tp) > payin_min

                    if od_above_thresh and payin_in_range(float(payin_od), remarks):
                        new_val = max(0.0, apply_state_override_po(po_formula, float(payin_od)))
                        new_od[-1] = new_val
                        applied_this = True
                        override_po_list.append(f"OD={po_formula}")

                    if tp_above_thresh and payin_in_range(float(payin_tp), remarks):
                        new_val = max(0.0, apply_state_override_po(po_formula, float(payin_tp)))
                        new_tp[-1] = new_val
                        applied_this = True
                        override_po_list.append(f"TP={po_formula}")

                    if applied_this:
                        override_applied = True
                        override_notes_list.append(
                            f"State={state_code} Seg={ov_segment} Co={company_name} PO={po_formula} Range={remarks}"
                        )
                        override_state_list.append(state_code)

            state_override_applied.append(override_applied)
            state_override_notes.append(' | '.join(override_notes_list) if override_notes_list else '')
            state_override_states.append(', '.join(sorted(set(override_state_list))) if override_state_list else '')
            state_override_po_col.append(', '.join(override_po_list) if override_po_list else '')

        # ── 7. Build output DataFrame ─────────────────────────────────────────
        df['calculated_od_rate'] = new_od
        df['calculated_tp_rate'] = new_tp

        def _safe_changed(o_od, c_od, o_tp, c_tp):
            try:
                a = abs(float(o_od if pd.notna(o_od) else 0) - float(c_od)) > 0.001
            except: a = False
            try:
                b = abs(float(o_tp if pd.notna(o_tp) else 0) - float(c_tp)) > 0.001
            except: b = False
            return a or b

        df['changed_payout'] = [
            _safe_changed(o_od, c_od, o_tp, c_tp)
            for o_od, c_od, o_tp, c_tp in zip(original_od, new_od, original_tp, new_tp)
        ]

        # State override columns
        df['state_override_applied'] = state_override_applied
        df['state_override_states']  = state_override_states
        df['state_override_po']      = state_override_po_col
        df['state_override_note']    = state_override_notes

        # Rule explanation columns
        df['od_rule_lob']         = od_lob
        df['od_rule_segment']     = od_seg
        df['od_rule_insurer']     = od_ins
        df['od_rule_po_formula']  = od_po
        df['od_rule_slab']        = od_slab
        df['od_rule_calculation'] = od_note
        df['tp_rule_lob']         = tp_lob
        df['tp_rule_segment']     = tp_seg
        df['tp_rule_insurer']     = tp_ins
        df['tp_rule_po_formula']  = tp_po
        df['tp_rule_slab']        = tp_slab
        df['tp_rule_calculation'] = tp_note

        # Reorder: new cols right after payout_tp_rate
        core_new = ['calculated_od_rate', 'calculated_tp_rate', 'changed_payout',
                    'state_override_applied', 'state_override_states',
                    'state_override_po', 'state_override_note']
        base_cols = [c for c in df.columns if c not in core_new]
        try:
            insert_pos = base_cols.index('payout_tp_rate') + 1
        except ValueError:
            insert_pos = len(base_cols)
        final_cols = base_cols[:insert_pos] + core_new + base_cols[insert_pos:]
        df = df[final_cols]

        # ── 8. Write Excel ────────────────────────────────────────────────────
        df = df.where(pd.notnull(df), '')
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Recalculated', index=False)
            wb  = writer.book
            ws  = writer.sheets['Recalculated']

            hdr_fmt = wb.add_format({
                'bold': True, 'bg_color': '#FFC000', 'font_color': '#000000',
                'border': 1, 'align': 'center', 'valign': 'vcenter',
                'font_name': 'Arial', 'font_size': 9,
            })
            data_fmt = wb.add_format({
                'border': 1, 'align': 'center', 'valign': 'vcenter',
                'font_name': 'Arial', 'font_size': 9,
            })
            changed_fmt = wb.add_format({
                'border': 1, 'align': 'center', 'valign': 'vcenter',
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#FFE0E0',
            })
            state_ov_fmt = wb.add_format({
                'border': 1, 'align': 'center', 'valign': 'vcenter',
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#E8F4FF',   # light blue for state-overridden rows
            })

            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, hdr_fmt)

            changed_col_idx    = list(df.columns).index('changed_payout')
            state_ov_col_idx   = list(df.columns).index('state_override_applied')

            for row_idx, row_data in enumerate(df.itertuples(index=False), start=1):
                is_changed        = row_data[changed_col_idx]
                is_state_override = row_data[state_ov_col_idx]
                if is_state_override:
                    fmt = state_ov_fmt
                elif is_changed:
                    fmt = changed_fmt
                else:
                    fmt = data_fmt
                for col_idx, val in enumerate(row_data):
                    ws.write(row_idx, col_idx, val, fmt)

            for col_idx, col_name in enumerate(df.columns):
                max_len = max(
                    len(str(col_name)),
                    df.iloc[:, col_idx].astype(str).str.len().max() if len(df) > 0 else 0
                )
                ws.set_column(col_idx, col_idx, min(max_len + 3, 35))

            ws.freeze_panes(1, 0)

        buf.seek(0)
        os.unlink(payin_path)

        # ── 9. Stats ──────────────────────────────────────────────────────────
        changed_total          = int(df['changed_payout'].sum())
        state_override_count   = int(df['state_override_applied'].sum())
        stats = {
            'total_rows':            total,
            'od_processed':          processed_od,
            'od_changed':            changed_od,
            'tp_processed':          processed_tp,
            'tp_changed':            changed_tp,
            'changed_payout':        changed_total,
            'state_override_count':  state_override_count,
        }

        base_name = secure_filename(payin_file.filename).rsplit('.', 1)[0]
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_name = f'{base_name}_recalculated_{ts}.xlsx'

        import urllib.parse
        response = send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response.headers['X-Stats'] = urllib.parse.quote(json.dumps(stats))
        response.headers['Access-Control-Expose-Headers'] = 'X-Stats'
        return response

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


# ============================================================================
# API: VALIDATE JSON
# ============================================================================

@app.route('/validate-rules', methods=['POST'])
def validate_rules():
    try:
        data  = request.get_json(force=True)
        raw   = data.get('rules_json', '')
        rules = json.loads(raw)
        lobs  = sorted(set(r.get('LOB', '') for r in rules if r.get('LOB')))
        return jsonify({'valid': True, 'total_rules': len(rules), 'lobs': lobs})
    except json.JSONDecodeError as e:
        return jsonify({'valid': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'valid': False, 'error': str(e)})


# ============================================================================
# API: RTO MASTER INFO
# ============================================================================

@app.route('/rto-info', methods=['GET'])
def rto_info():
    states = sorted(set(
        str(r.get('rto_code', '')).strip().upper()
        for r in _RTO_MASTER if r.get('rto_code')
    ))
    return jsonify({
        'total_records': len(_RTO_MASTER),
        'states': states,
    })


if __name__ == '__main__':
    app.run(debug=True, port=5001)
