"""
Payout Recalculator — Flask Server
====================================
Serves the UI and processes PayinConfig uploads against JSON payout rules.
"""

import os
import io
import json
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

# Import all logic from recalculate_payout.py (same folder)
import recalculate_payout as calc

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB

UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('outputs')
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ============================================================================
# INDEX
# ============================================================================

@app.route('/')
def index():
    return render_template('index.html')


# ============================================================================
# API: PROCESS — receives files + rules JSON, returns processed Excel
# ============================================================================

@app.route('/process', methods=['POST'])
def process():
    try:
        # ── 1. Validate inputs ────────────────────────────────────────────────
        payin_file   = request.files.get('payin_file')
        rules_json   = request.form.get('rules_json', '').strip()

        payin_min    = float(request.form.get('payin_min', 5))   # payin <= this → 0
        irda_pvt_car = float(request.form.get('irda_pvt_car', 15))
        irda_tw      = float(request.form.get('irda_tw', 17.5))

        # Segment filters (empty = apply to all segments)
        def _parse_segs(key):
            raw = request.form.get(key, '').strip()
            if not raw or raw == 'ALL': return []
            return [s.strip().upper() for s in raw.split(',') if s.strip()]
        thresh_segs   = _parse_segs('thresh_segs')
        irda_pvt_segs = _parse_segs('irda_pvt_segs')
        irda_tw_segs  = _parse_segs('irda_tw_segs')
        print(f"[/process] thresh_segs={thresh_segs or 'all'}, irda_pvt_segs={irda_pvt_segs or 'all'}, irda_tw_segs={irda_tw_segs or 'all'}")

        if not payin_file or payin_file.filename == '':
            return jsonify({'success': False, 'error': 'PayinConfig file is required'}), 400
        if not rules_json:
            return jsonify({'success': False, 'error': 'Payout rules JSON is required'}), 400

        # ── 2. Parse rules JSON ───────────────────────────────────────────────
        try:
            rules = json.loads(rules_json)
        except json.JSONDecodeError as e:
            return jsonify({'success': False, 'error': f'Invalid JSON: {e}'}), 400

        # ── 3. Save uploads to temp files ─────────────────────────────────────
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_payin:
            payin_file.save(tmp_payin.name)
            payin_path = tmp_payin.name

        # ── 5. Load PayinConfig — ALL rows, no filtering ─────────────────────
        df = pd.read_excel(payin_path)
        df.columns = [c.strip() for c in df.columns]
        total = len(df)
        print(f"[/process] Total rows loaded: {total} — no state row filtering applied")

        required_cols = [
            'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
            'sub_product_name', 'segment', 'company_code',
            'vehicle_type_id', 'from_weightage_kg', 'to_weightage_kg'
        ]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return jsonify({'success': False, 'error': f'Missing columns: {missing}'}), 400

        # ── 7. Monkey-patch IRDA thresholds if user customised them ──────────
        # We temporarily override the module-level check by wrapping compute_payout
        orig_compute = calc.compute_payout

        def patched_compute(payin, sub_product_name, segment, company_code,
                            vehicle_type_id, from_wt, to_wt,
                            rules_arg, is_od=True):
            explanation = {
                "lob": "", "segment": "", "insurer_matched": "",
                "po_formula": "", "remarks_slab": "",
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

            # Map sub_product_name → chip label for segment filter matching
            SP_TO_CHIP = {
                'private car':   'PVT CAR',
                'two wheeler':   'TW',
                'goods vehicle': 'CV / GCV',
                'passenger':     'PCV',
                'taxi':          'TAXI',
                'bus':           'BUS',
                'misd':          'MISD',
                'tractor':       'MISD',
            }
            sp_lower = str(sub_product_name).strip().lower()
            sp_chip  = SP_TO_CHIP.get(sp_lower, sp_lower.upper())

            def seg_match(filters):
                """True if no filter set, or sp_chip matches any selected label."""
                if not filters:
                    return True
                return any(f.upper() in sp_chip.upper() or sp_chip.upper() in f.upper()
                           for f in filters)

            # Rule 1: payin <= payin_min → 0  (check segment filter)
            if p <= payin_min and seg_match(thresh_segs):
                explanation["calculation_note"] = f"Payin {p} <= {payin_min} — payout is 0"
                explanation["po_formula"] = f"Payin <= {payin_min} → 0"
                return 0.0, explanation

            # Rule 2: IRDA rates (OD only) → 0  (check segment filter per rate)
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
                                vehicle_type_id, from_wt, to_wt,
                                rules_arg, is_od=is_od)

        # ── 7. Process rows ───────────────────────────────────────────────────
        original_od = df['payout_od_rate'].tolist()
        original_tp = df['payout_tp_rate'].tolist()

        new_od, new_tp = [], []
        changed_od = changed_tp = processed_od = processed_tp = 0

        od_lob, od_seg, od_ins, od_po, od_slab, od_note = [], [], [], [], [], []
        tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_note = [], [], [], [], [], []

        blank = {
            'lob': '', 'segment': '', 'insurer_matched': '', 'po_formula': '',
            'remarks_slab': '', 'calculation_note': 'Payin is 0'
        }

        for _, row in df.iterrows():
            sp         = row['sub_product_name']
            seg        = row['segment']
            ins        = row['company_code']
            vt         = row['vehicle_type_id']
            f_wt       = row['from_weightage_kg']
            t_wt       = row['to_weightage_kg']
            # OD
            payin_od = row['payin_od_rate']
            old_od   = row['payout_od_rate']
            if pd.notna(payin_od) and float(payin_od) != 0:
                processed_od += 1
                calc_od, expl_od = patched_compute(
                    payin_od, sp, seg, ins, vt, f_wt, t_wt,
                    rules, is_od=True
                )
                new_od.append(calc_od)
                if abs(float(old_od) - calc_od) > 0.001:
                    changed_od += 1
            else:
                new_od.append(0.0 if pd.isna(old_od) else old_od)
                expl_od = blank
            od_lob.append(expl_od['lob']); od_seg.append(expl_od['segment'])
            od_ins.append(expl_od['insurer_matched']); od_po.append(expl_od['po_formula'])
            od_slab.append(expl_od['remarks_slab'])
            od_note.append(expl_od['calculation_note'])

            # TP
            payin_tp = row['payin_tp_rate']
            old_tp   = row['payout_tp_rate']
            if pd.notna(payin_tp) and float(payin_tp) != 0:
                processed_tp += 1
                calc_tp, expl_tp = patched_compute(
                    payin_tp, sp, seg, ins, vt, f_wt, t_wt,
                    rules, is_od=False
                )
                new_tp.append(calc_tp)
                if abs(float(old_tp) - calc_tp) > 0.001:
                    changed_tp += 1
            else:
                new_tp.append(0.0 if pd.isna(old_tp) else old_tp)
                expl_tp = blank
            tp_lob.append(expl_tp['lob']); tp_seg.append(expl_tp['segment'])
            tp_ins.append(expl_tp['insurer_matched']); tp_po.append(expl_tp['po_formula'])
            tp_slab.append(expl_tp['remarks_slab'])
            tp_note.append(expl_tp['calculation_note'])

        # ── 8. Build output DataFrame ─────────────────────────────────────────
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

        # Reorder: insert new cols right after payout_tp_rate
        new_cols  = ['calculated_od_rate', 'calculated_tp_rate', 'changed_payout']
        base_cols = [c for c in df.columns if c not in new_cols]
        try:
            insert_pos = base_cols.index('payout_tp_rate') + 1
        except ValueError:
            insert_pos = len(base_cols)
        final_cols = base_cols[:insert_pos] + new_cols + base_cols[insert_pos:]
        df = df[final_cols]

        # ── 9. Write output Excel to memory buffer ────────────────────────────
        # Replace NaN/inf with empty string so blank cells stay blank (not #NUM!)
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

            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, hdr_fmt)

            changed_col_idx = list(df.columns).index('changed_payout')
            for row_idx, row_data in enumerate(df.itertuples(index=False), start=1):
                is_changed = row_data[changed_col_idx]
                fmt = changed_fmt if is_changed else data_fmt
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

        # ── 10. Clean up temp files ───────────────────────────────────────────
        os.unlink(payin_path)
        # ── 11. Build summary stats ───────────────────────────────────────────
        changed_total = int(df['changed_payout'].sum())
        stats = {
            'total_rows':     total,
            'od_processed':   processed_od,
            'od_changed':     changed_od,
            'tp_processed':   processed_tp,
            'tp_changed':     changed_tp,
            'changed_payout': changed_total,
        }

        # Generate output filename
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
# API: VALIDATE JSON — quick check that rules JSON is parseable
# ============================================================================

@app.route('/validate-rules', methods=['POST'])
def validate_rules():
    try:
        data = request.get_json(force=True)
        raw  = data.get('rules_json', '')
        rules = json.loads(raw)
        lobs = sorted(set(r.get('LOB', '') for r in rules if r.get('LOB')))
        return jsonify({
            'valid': True,
            'total_rules': len(rules),
            'lobs': lobs,
        })
    except json.JSONDecodeError as e:
        return jsonify({'valid': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'valid': False, 'error': str(e)})


if __name__ == '__main__':
    app.run(debug=True, port=5001)
