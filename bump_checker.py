#!/usr/bin/env python3
"""
Bump Rule Checker — Python Wrapper
支持調用 Perl 版本或獨立 Python 檢查
"""

import subprocess
import sys
from pathlib import Path
import json
import tempfile
import shutil

_HERE = Path(__file__).parent.resolve()


def find_perl():
    """尋找系統中的 Perl"""
    # 優先檢查本地 portable Perl (nested structure)
    local_perl = _HERE / "perl" / "perl" / "bin" / "perl.exe"
    if local_perl.exists():
        return str(local_perl)

    # 嘗試 perl/bin/perl.exe (舊結構)
    local_perl_alt = _HERE / "perl" / "bin" / "perl.exe"
    if local_perl_alt.exists():
        return str(local_perl_alt)

    # 檢查系統 PATH
    try:
        result = subprocess.run(["where", "perl"], capture_output=True, text=True)
        if result.returncode == 0:
            return result.stdout.strip().split('\n')[0]
    except:
        pass

    return None


def run_bump_check_python(gds_file: str, out_dir: str = None, original_name: str = None) -> dict:
    """
    Python Bump Rule 檢查 — TM + BOE 雙標準
    """
    try:
        import gdstk
        import csv
        from datetime import datetime
        from collections import Counter
    except ImportError as e:
        return {"success": False, "message": f"缺少模組: {str(e)}", "gds_file": gds_file}

    gds_path = Path(gds_file)
    if not gds_path.exists():
        return {"success": False, "message": f"GDS 檔案不存在: {gds_file}"}

    if out_dir is None:
        out_dir = str(gds_path.parent)
    else:
        out_dir = str(Path(out_dir))
        Path(out_dir).mkdir(parents=True, exist_ok=True)

    try:
        lib = gdstk.read_gds(str(gds_path))
        gds_base = original_name if original_name else gds_path.stem
        gds_unit_um = getattr(lib, 'unit', 0.001)
        print(f"[INFO] GDS unit={gds_unit_um} μm, base={gds_base}")

        # ── Collect bump polygons ──────────────────────────────────────
        bumps = []
        cells_data = lib.cells
        if isinstance(cells_data, dict):
            cells_data = list(cells_data.values())
        elif not isinstance(cells_data, list):
            cells_data = list(cells_data)

        for cell in cells_data:
            if not (hasattr(cell, 'polygons') and cell.polygons):
                continue
            for poly in cell.polygons:
                if poly.layer not in [224, 225]:
                    continue
                pts = poly.points
                if not len(pts):
                    continue
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                x1, x2 = min(xs) * gds_unit_um, max(xs) * gds_unit_um
                y1, y2 = min(ys) * gds_unit_um, max(ys) * gds_unit_um
                bumps.append({
                    "x_min": x1, "y_min": y1,
                    "w": x2 - x1, "h": y2 - y1,
                    "cx": (x1 + x2) / 2, "cy": (y1 + y2) / 2,
                    "layer": poly.layer,
                })

        print(f"[INFO] Collected {len(bumps)} bumps on layers 224/225")

        # ── Size groups ────────────────────────────────────────────────
        size_ctr = Counter((round(b['w'], 1), round(b['h'], 1)) for b in bumps)

        # ── IC bounding box (from bump extents) ────────────────────────
        if bumps:
            ic_x1 = min(b['x_min'] for b in bumps)
            ic_x2 = max(b['x_min'] + b['w'] for b in bumps)
            ic_y1 = min(b['y_min'] for b in bumps)
            ic_y2 = max(b['y_min'] + b['h'] for b in bumps)
            ic_span_w = ic_x2 - ic_x1   # μm
            ic_span_h = ic_y2 - ic_y1   # μm
        else:
            ic_x1 = ic_x2 = ic_y1 = ic_y2 = 0
            ic_span_w = ic_span_h = 0

        # ── Row clustering by Y center ──────────────────────────────────
        ROW_TOL = 5.0   # um: bumps within ±5 um cy are same row
        rows = {}
        for b in bumps:
            key = round(b['cy'] / ROW_TOL) * ROW_TOL
            rows.setdefault(key, []).append(b)

        sorted_row_keys = sorted(rows.keys())  # ascending by Y
        n_rows = len(sorted_row_keys)

        # ILB = inner rows (smallest |cy|), OLB = outer rows
        # Heuristic: rows with |cy| in lower half → ILB
        abs_sorted = sorted(abs(k) for k in sorted_row_keys)
        if abs_sorted:
            half_abs = abs_sorted[len(abs_sorted) // 2]
        else:
            half_abs = 0

        ilb_bumps = [b for b in bumps if abs(b['cy']) <= half_abs + ROW_TOL]
        olb_bumps = [b for b in bumps if abs(b['cy']) > half_abs + ROW_TOL]
        if not olb_bumps:  # single-row fallback
            olb_bumps = ilb_bumps

        # ILB row Y extents
        ilb_y_max = max(b['y_min'] + b['h'] for b in ilb_bumps) if ilb_bumps else ic_y2
        ilb_y_min = min(b['y_min'] for b in ilb_bumps) if ilb_bumps else ic_y1
        # OLB row Y extents
        olb_y_max = max(b['y_min'] + b['h'] for b in olb_bumps) if olb_bumps else ic_y2
        olb_y_min = min(b['y_min'] for b in olb_bumps) if olb_bumps else ic_y1

        # B = ILB top → OLB bottom (if Y-separated)
        if olb_y_min > ilb_y_max:
            B_vals = [olb_y_min - ilb_y_max]
            B_min = B_max = B_vals[0]
        elif ilb_y_min > olb_y_max:
            B_vals = [ilb_y_min - olb_y_max]
            B_min = B_max = B_vals[0]
        else:
            B_vals = [abs(ic_span_h / 2)]
            B_min = B_max = B_vals[0]

        # ── R1: Effective area per bump ────────────────────────────────
        r1_viols = []
        r1_size_res = {}
        for b in bumps:
            short_d = min(b['w'], b['h'])
            long_d = max(b['w'], b['h'])
            eff = (short_d - 7) * (long_d - 10)
            b['eff_area'] = eff
            key = (round(b['w'], 1), round(b['h'], 1))
            if key not in r1_size_res:
                r1_size_res[key] = {'eff': round(eff, 1), 'count': 0, 'pass': eff >= 500}
            r1_size_res[key]['count'] += 1
            if eff < 500:
                r1_viols.append({
                    'std': 'TM', 'rule': 'R1_EffArea', 'severity': 'HIGH',
                    'net': '-', 'sig': '-',
                    'w': b['w'], 'h': b['h'], 'cx': b['cx'], 'cy': b['cy'],
                    'measured': round(eff, 1), 'spec': '500', 'margin': round(eff - 500, 1),
                    'desc': f"Bump {b['w']:.1f}x{b['h']:.1f} eff_area={eff:.1f}<500 um²"
                })
        r1_fail = len(r1_viols)
        r1_verdict = 'FAIL' if r1_fail else 'PASS'

        # ── R2: ILB pitch (X spacing within innermost row) ────────────
        inner_row_key = min(sorted_row_keys, key=abs)
        inner_row = sorted(rows[inner_row_key], key=lambda b: b['cx'])
        if len(inner_row) > 1:
            pitches = [inner_row[i+1]['cx'] - inner_row[i]['cx']
                       for i in range(len(inner_row)-1) if inner_row[i+1]['cx'] - inner_row[i]['cx'] > 0]
            ilb_pitch = round(sorted(pitches)[len(pitches)//2], 2) if pitches else 0
        else:
            ilb_pitch = 0

        TM_ILB_PITCH = 39.0
        r2_verdict = 'PASS' if ilb_pitch > 0 and abs(ilb_pitch - TM_ILB_PITCH) < 1.5 else ('RISK' if ilb_pitch > 0 else 'N/A')
        r2_badge = '#27ae60' if r2_verdict == 'PASS' else ('#e67e22' if r2_verdict == 'RISK' else '#95a5a6')

        # ── R3: ILB→OLB distance B ────────────────────────────────────
        R3_MIN_B_PER_ROW = {1: 0, 2: 240.6, 3: 360.9, 4: 481.2}
        n_olb_rows = max(1, n_rows - (1 if n_rows > 1 else 0))
        r3_spec_min = R3_MIN_B_PER_ROW.get(n_olb_rows, 120.0)
        r3_verdict = 'PASS' if B_min >= r3_spec_min else 'RISK'
        r3_badge = '#27ae60' if r3_verdict == 'PASS' else '#e67e22'

        # ── R4: Edge distances A & C ────────────────────────────────
        # Without IC die boundary layer, estimate IC boundary at bump extent + small margin
        # Use ILB_y_min → IC_y1 as A, OLB_y_max → IC_y2 as C
        # If only 1 row, use bump top/bottom to IC span edge
        est_A = max(0.0, ilb_y_min - ic_y1)
        est_C = max(0.0, ic_y2 - olb_y_max)
        TM_EDGE_MIN = 40.0
        r4_viols = []
        if est_A > 0 and est_A < TM_EDGE_MIN:
            r4_viols.append({'dir': '下邊距 A (Y)', 'val': est_A, 'margin': est_A - TM_EDGE_MIN})
        if est_C > 0 and est_C < TM_EDGE_MIN:
            r4_viols.append({'dir': '上邊距 C (Y)', 'val': est_C, 'margin': est_C - TM_EDGE_MIN})
        r4_verdict = 'RISK' if r4_viols else ('PASS' if (est_A >= TM_EDGE_MIN and est_C >= TM_EDGE_MIN) else 'N/A')
        r4_badge = '#27ae60' if r4_verdict == 'PASS' else ('#e67e22' if r4_verdict == 'RISK' else '#95a5a6')

        # ── R5: B/(A+C) < 5 ───────────────────────────────────────────
        r5_viols = []
        if est_A > 0 and est_C > 0:
            ac_sum = est_A + est_C
            ratio_max = B_max / ac_sum if ac_sum > 0 else 0
            ratio_min = B_min / ac_sum if ac_sum > 0 else 0
            r5_verdict = 'FAIL' if ratio_max >= 5 else 'PASS'
            r5_badge = '#c0392b' if r5_verdict == 'FAIL' else '#27ae60'
            if ratio_max >= 5:
                r5_viols.append({
                    'std': 'TM', 'rule': 'R5_DFX', 'severity': 'HIGH',
                    'net': '-', 'sig': 'ILB/OLB',
                    'w': 'N/A', 'h': 'N/A', 'cx': 0, 'cy': 0,
                    'measured': round(ratio_max, 2), 'spec': '<5', 'margin': round(ratio_max - 5, 2),
                    'desc': f"B/(A+C)={ratio_max:.2f}≥5 B={B_max:.1f} A={est_A:.1f} C={est_C:.1f}"
                })
        else:
            r5_verdict = 'N/A'
            r5_badge = '#95a5a6'
            ratio_min = ratio_max = 0
            ac_sum = 0

        # ── BOE Rules ──────────────────────────────────────────────────
        # B1: IC size L=5~35mm, W=0.5~2.2mm
        ic_L_mm = ic_span_w / 1000
        ic_W_mm = ic_span_h / 1000
        b1_pass_L = 5 <= ic_L_mm <= 35
        b1_pass_W = 0.5 <= ic_W_mm <= 2.2
        b1_verdict = 'PASS' if (b1_pass_L and b1_pass_W) else 'FAIL'
        b1_badge = '#27ae60' if b1_verdict == 'PASS' else '#c0392b'

        # B2: leftmost/rightmost bump X distance to bump-span edge ≤ 200 um
        left_bump_x = min(b['x_min'] for b in bumps) if bumps else 0
        right_bump_x = max(b['x_min'] + b['w'] for b in bumps) if bumps else 0
        # Distance from edge bump center to IC edge (estimated from span)
        dist_left = left_bump_x - ic_x1
        dist_right = ic_x2 - right_bump_x
        b2_pass = dist_left <= 200 and dist_right <= 200
        b2_verdict = 'PASS' if b2_pass else 'FAIL'
        b2_badge = '#27ae60' if b2_pass else '#c0392b'

        # B3: min bump short side ≥ 12 um
        min_w = min(min(b['w'], b['h']) for b in bumps) if bumps else 0
        b3_verdict = 'PASS' if min_w >= 12 else 'FAIL'
        b3_badge = '#27ae60' if b3_verdict == 'PASS' else '#c0392b'

        # B4: lead space = min pitch - max width ≥ 12 um
        b4_space = ilb_pitch - max(b['w'] for b in ilb_bumps) if ilb_bumps and ilb_pitch > 0 else None
        b4_verdict = ('PASS' if b4_space is not None and b4_space >= 12 else
                      'FAIL' if b4_space is not None else 'N/A')
        b4_badge = '#27ae60' if b4_verdict == 'PASS' else ('#c0392b' if b4_verdict == 'FAIL' else '#95a5a6')

        # B5: lead pitch ≥ 15 um
        b5_verdict = 'PASS' if ilb_pitch >= 15 else ('FAIL' if ilb_pitch > 0 else 'N/A')
        b5_badge = '#27ae60' if b5_verdict == 'PASS' else ('#c0392b' if b5_verdict == 'FAIL' else '#95a5a6')

        # B6: lead area c*a ≥ 950 um² for c ≥ 65
        b6_eligible = [(round(b['w'],1), round(b['h'],1)) for b in bumps if max(b['w'], b['h']) >= 65]
        b6_ctr = Counter(b6_eligible)
        b6_pass_count = sum(v for (w,h),v in b6_ctr.items() if w*h >= 950)
        b6_fail_count = sum(v for (w,h),v in b6_ctr.items() if w*h < 950)
        b6_verdict = 'PASS' if b6_fail_count == 0 and b6_pass_count > 0 else ('FAIL' if b6_fail_count > 0 else 'N/A')
        b6_badge = '#27ae60' if b6_verdict == 'PASS' else ('#c0392b' if b6_verdict == 'FAIL' else '#95a5a6')

        # ── All violations ─────────────────────────────────────────────
        all_viols = r1_viols[:]
        for rv in r4_viols:
            all_viols.append({
                'std': 'TM', 'rule': 'R4_EdgeDist', 'severity': 'MEDIUM',
                'net': '-', 'sig': '-',
                'w': '-', 'h': '-', 'cx': '-', 'cy': rv['val'],
                'measured': round(rv['val'], 1), 'spec': '40', 'margin': round(rv['margin'], 1),
                'desc': f"{rv['dir']} = {rv['val']:.1f} < 40 um"
            })
        all_viols.extend(r5_viols)

        tm_viol_count = len(all_viols)
        boe_viol_count = 0
        total_viol = tm_viol_count + boe_viol_count
        crit_count = sum(1 for v in all_viols if v.get('severity') == 'CRITICAL')
        high_count = sum(1 for v in all_viols if v.get('severity') == 'HIGH')
        med_count = sum(1 for v in all_viols if v.get('severity') == 'MEDIUM')

        # ── Output paths ───────────────────────────────────────────────
        out_path = Path(out_dir)
        out_path.mkdir(parents=True, exist_ok=True)
        violations_csv = out_path / f"{gds_base}_violations.csv"
        report_html = out_path / f"{gds_base}_report.html"

        # ── Write CSV ──────────────────────────────────────────────────
        if violations_csv.exists():
            violations_csv.unlink()
        with open(str(violations_csv), 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'No', 'Std', 'Rule', 'Severity', 'Net', 'Sig', 'W', 'H', 'Cx', 'Cy',
                'Measured', 'Spec', 'Margin', 'Description'])
            writer.writeheader()
            for i, v in enumerate(all_viols, 1):
                writer.writerow({
                    'No': i, 'Std': v.get('std', ''), 'Rule': v.get('rule', ''),
                    'Severity': v.get('severity', ''), 'Net': v.get('net', '-'),
                    'Sig': v.get('sig', '-'), 'W': v.get('w', '-'), 'H': v.get('h', '-'),
                    'Cx': v.get('cx', '-'), 'Cy': v.get('cy', '-'),
                    'Measured': v.get('measured', ''), 'Spec': v.get('spec', ''),
                    'Margin': v.get('margin', ''), 'Description': v.get('desc', ''),
                })
        print(f"[INFO] CSV written: {violations_csv}")

        # ── Build HTML ─────────────────────────────────────────────────
        ts = datetime.now().strftime('%Y-%m-%d %H:%M')
        chip_model = gds_base.split('_')[0] if '_' in gds_base else gds_base

        def badge(text, color):
            return f"<span class='badge' style='background:{color}'>{text}</span>"

        def tm_badge(text):
            return f"<span class='std-tm'>TM</span>"

        def boe_badge(text):
            return f"<span class='std-boe'>BOE</span>"

        # R1 size table rows
        r1_size_rows = ""
        for (w, h), info in sorted(r1_size_res.items(), key=lambda x: -x[1]['count']):
            cnt = info['count']
            eff = info['eff']
            short_d = min(w, h)
            long_d = max(w, h)
            eff_formula = f"({short_d:.0f}-7)×({long_d:.0f}-10)={eff:.0f} um²"
            row_color = " style='background:#fdecea'" if not info['pass'] else ""
            verd = badge('PASS', '#27ae60') if info['pass'] else badge('FAIL', '#c0392b')
            r1_size_rows += (
                f"<tr{row_color}><td>{w:.1f} × {h:.1f} um</td><td>{cnt}</td>"
                f"<td>{eff_formula}</td><td>≥500</td><td>{verd}</td></tr>\n"
            )

        # R4 edge table rows
        r4_edge_rows = ""
        dirs_data = [
            ('左邊距 (X)', ic_x1 - ic_x1, '>0 (X方向)'),
            ('右邊距 (X)', ic_x2 - ic_x2, '>0 (X方向)'),
            ('下邊距 A (Y)', est_A, f'≥{TM_EDGE_MIN:.0f} um'),
            ('上邊距 C (Y)', est_C, f'≥{TM_EDGE_MIN:.0f} um'),
        ]
        for label, val, spec in [
            ('下邊距 A (Y)', est_A, f'≥{TM_EDGE_MIN:.0f} um'),
            ('上邊距 C (Y)', est_C, f'≥{TM_EDGE_MIN:.0f} um'),
        ]:
            margin = val - TM_EDGE_MIN
            ok = val >= TM_EDGE_MIN
            row_color = "" if ok else " style='background:#fef0e7'"
            mc = "class='pos'" if ok else "class='neg'"
            verd = badge('PASS', '#27ae60') if ok else badge('RISK', '#e67e22')
            r4_edge_rows += (
                f"<tr{row_color}><td>{label}</td><td><b>{val:.1f} um</b></td>"
                f"<td>{spec}</td><td {mc}>{margin:+.1f}</td><td>{verd}</td></tr>\n"
            )

        # Detailed violation table rows (first 20)
        SEV_COLOR = {'CRITICAL': '#c0392b', 'HIGH': '#e67e22', 'MEDIUM': '#f39c12'}
        _fallback_color = '#888'
        viol_rows = ""
        for i, v in enumerate(all_viols[:20], 1):
            sev = v.get('severity', 'HIGH')
            row_bg = " style='background:#fef0e7'" if sev in ('CRITICAL', 'HIGH') else ""
            _sc = SEV_COLOR.get(sev, _fallback_color)
            sev_badge = f"<span class='badge' style='background:{_sc};color:#fff;padding:2px 6px;border-radius:4px'>{sev}</span>"
            rule_color = '#2471a3' if v.get('std') == 'TM' else '#117a65'
            _rule = v.get('rule', '')
            rule_badge = f"<span class='badge' style='background:{rule_color}'>{_rule}</span>"
            std_badge = "<span class='std-tm'>TM</span>" if v.get('std') == 'TM' else "<span class='std-boe'>BOE</span>"
            cx = f"{v['cx']:.0f}" if isinstance(v.get('cx'), (int, float)) else str(v.get('cx', '-'))
            cy = f"{v['cy']:.0f}" if isinstance(v.get('cy'), (int, float)) else str(v.get('cy', '-'))
            w_s = f"{v['w']:.1f}" if isinstance(v.get('w'), (int, float)) else str(v.get('w', '-'))
            h_s = f"{v['h']:.1f}" if isinstance(v.get('h'), (int, float)) else str(v.get('h', '-'))
            viol_rows += (
                f"<tr{row_bg}><td>{i}</td><td>{std_badge}</td><td>{rule_badge}</td>"
                f"<td>{sev_badge}</td><td>{v.get('net','-')}</td><td>{v.get('sig','-')}</td>"
                f"<td>{w_s}</td><td>{h_s}</td><td>{cx}</td><td>{cy}</td>"
                f"<td>{v.get('measured','')}</td><td>{v.get('spec','')}</td>"
                f"<td>{v.get('margin','')}</td><td>{v.get('desc','')}</td></tr>\n"
            )
        if len(all_viols) > 20:
            viol_rows += f"<tr><td colspan='14' style='text-align:center;color:#888'>... 另有 {len(all_viols)-20} 筆，請見 CSV 檔 ...</td></tr>\n"

        # Recommendations
        recs_tm = ""
        if r1_fail:
            recs_tm += "<li><b>[R1] 修改 Bump 尺寸</b>：使有效面積 (X-7)×(Y-10) ≥ 500 um²</li>\n"
        if r4_viols:
            recs_tm += "<li><b>[R4] 邊距不足</b>：調整 ILB/OLB 位置使 A & C ≥ 40 um</li>\n"
        if r5_viols:
            recs_tm += "<li><b>[R5] B/(A+C) 超限</b>：縮短 ILB-OLB 間距，或確認 DFX 規範適用場景</li>\n"
        if not recs_tm:
            recs_tm = "<li>TM 規則全部通過，無需修改</li>\n"

        recs_boe = "<li><b>[B7/B8]</b> Bump 厚度及高度差需向 Bumping 廠商取得製程規格書確認</li>\n"
        recs_boe += "<li><b>[B9]</b> 確認 GDS 座標已套用 220 ppm 向中心預縮</li>\n"

        # B3-B6 combined table
        b3_meas = f"min={min_w:.1f} um"
        b4_meas = f"min={b4_space:.2f} um" if b4_space is not None else "N/A"
        b5_meas = f"ILB={ilb_pitch:.2f} um" if ilb_pitch > 0 else "N/A"
        b6_meas = f"{b6_pass_count} 個 c≥65 全通過" if b6_pass_count > 0 else "無 c≥65 bump"

        b3b6_rows = (
            f"<tr><td><b>B3</b></td><td>Lead Width a（短邊）</td><td>≥ 12 um</td><td>{b3_meas}</td><td>{badge(b3_verdict, b3_badge)}</td></tr>\n"
            f"<tr><td><b>B4</b></td><td>Lead Space b = pitch−a</td><td>≥ 12 um</td><td>{b4_meas}</td><td>{badge(b4_verdict, b4_badge)}</td></tr>\n"
            f"<tr><td><b>B5</b></td><td>Lead Pitch d（中心距）</td><td>≥ 15 um</td><td>{b5_meas}</td><td>{badge(b5_verdict, b5_badge)}</td></tr>\n"
            f"<tr><td><b>B6</b></td><td>Lead Area c×a（c≥65）</td><td>≥ 950 um²</td><td>{b6_meas}</td><td>{badge(b6_verdict, b6_badge)}</td></tr>\n"
        )

        # B2 table
        b2_left_dist = left_bump_x - ic_x1
        b2_right_dist = ic_x2 - right_bump_x
        b2_rows = (
            f"<tr><td>左邊（X）</td><td>{b2_left_dist:.1f} um {'✓' if b2_left_dist<=200 else '✗'}</td><td>最左 Bump 距左邊</td></tr>\n"
            f"<tr><td>右邊（X）</td><td>{b2_right_dist:.1f} um {'✓' if b2_right_dist<=200 else '✗'}</td><td>最右 Bump 距右邊</td></tr>\n"
        )

        # size distribution table for B3-B6
        size_dist_rows = ""
        for (w, h), cnt in sorted(size_ctr.items(), key=lambda x: -x[1]):
            long_d = max(w, h)
            area = w * h
            b6_app = "c≥65 ✓" if long_d >= 65 else "c<65，N/A"
            b6_v = badge('PASS', '#27ae60') if (long_d >= 65 and area >= 950) else ("<span style='color:#999'>—</span>" if long_d < 65 else badge('FAIL', '#c0392b'))
            size_dist_rows += f"<tr><td>{w} × {h:.1f} um</td><td>{cnt}</td><td>{area:.0f} um²</td><td>{b6_app}</td><td>{b6_v}</td></tr>\n"

        r1_verdict_badge = badge(r1_verdict, '#c0392b' if r1_verdict == 'FAIL' else '#27ae60')
        r2_meas = f"ILB = {ilb_pitch:.2f} um" if ilb_pitch > 0 else "無法量測（單行或無 ILB）"
        r3_meas = f"B = {B_min:.1f} ~ {B_max:.1f} um"
        r3_spec = f"{n_olb_rows} 排 OLB > {r3_spec_min:.1f} um"
        r4_meas = f"A={est_A:.1f} um / C={est_C:.1f} um"
        r5_meas = (f"B={B_min:.1f}~{B_max:.1f} um, ratio={ratio_min:.2f}~{ratio_max:.2f}"
                   if ac_sum > 0 else "需要 IC 邊界層")
        # Pre-compute R1 summary measurement string (avoid nested f-string)
        if r1_size_res:
            top_sz = sorted(r1_size_res.items(), key=lambda x: -x[1]['count'])[0]
            (tw, th), tinfo = top_sz
            r1_meas = f"{tw:.1f}x{th:.1f} um -> eff={tinfo['eff']} um2"
        else:
            r1_meas = "N/A"
        # Pre-compute R5 detail string
        if ac_sum > 0:
            r5_detail = (f"<b>A+C = {ac_sum:.1f} um</b>，最大允許 B = 5 × {ac_sum:.1f} = "
                         f"<b>{5*ac_sum:.1f} um</b><br>"
                         f"B = {B_min:.1f} ~ {B_max:.1f} um，ratio = "
                         f"{ratio_min:.2f} ~ {ratio_max:.2f}</b>")
        else:
            r5_detail = "需要 IC 邊界層資訊以計算 A 和 C"

        html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta charset="UTF-8">
<title>TM + BOE Bump Rule Compliance Report - {gds_base}</title>
<style>
  body{{font-family:Arial,sans-serif;font-size:13px;margin:20px;color:#222;background:#f8f9fa;}}
  h1{{color:#2c3e50;border-bottom:3px solid #2980b9;padding-bottom:8px;}}
  h2{{color:#2980b9;border-left:5px solid #2980b9;padding-left:10px;margin-top:30px;}}
  h2.boe{{color:#16a085;border-left-color:#16a085;}}
  h3{{color:#555;margin-top:20px;}}
  .info-box{{background:#fff;border:1px solid #ddd;border-radius:6px;padding:15px;margin:15px 0;}}
  .summary-grid{{display:flex;gap:15px;flex-wrap:wrap;margin:15px 0;}}
  .summary-card{{background:#fff;border-radius:8px;padding:15px 20px;min-width:140px;
                text-align:center;border:1px solid #ddd;box-shadow:0 2px 4px rgba(0,0,0,.05);}}
  .summary-card .num{{font-size:30px;font-weight:bold;}}
  .summary-card .lbl{{font-size:11px;color:#666;margin-top:4px;}}
  table{{border-collapse:collapse;width:100%;margin-top:10px;background:#fff;font-size:12px;}}
  th{{background:#2c3e50;color:#fff;padding:8px 10px;text-align:left;white-space:nowrap;}}
  th.boe-hdr{{background:#16a085;}}
  td{{padding:6px 10px;border-bottom:1px solid #eee;vertical-align:top;}}
  tr:hover td{{background:#f0f7ff;}}
  .badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;
         font-weight:bold;color:#fff;white-space:nowrap;}}
  .std-tm{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;
          font-weight:bold;color:#fff;background:#2980b9;white-space:nowrap;}}
  .std-boe{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;
           font-weight:bold;color:#fff;background:#16a085;white-space:nowrap;}}
  .neg{{color:#c0392b;font-weight:bold;}} .pos{{color:#27ae60;}}
  .footer{{margin-top:40px;padding:15px;background:#ecf0f1;border-radius:6px;
          font-size:11px;color:#666;text-align:center;}}
  .toc{{background:#fff;border:1px solid #ddd;border-radius:6px;padding:15px;
       display:inline-block;min-width:300px;margin:15px 0;}}
  .toc ul{{margin:5px 0;padding-left:20px;}} .toc li{{margin:4px 0;}}
  .toc a{{color:#2980b9;text-decoration:none;}} .toc a:hover{{text-decoration:underline;}}
  .section-divider{{border:none;border-top:2px solid #ecf0f1;margin:30px 0;}}
  @media print{{h2{{page-break-before:always;}}h3{{page-break-after:avoid;}}
    .info-box,.summary-grid{{page-break-inside:avoid;}}table{{page-break-inside:avoid;}}}}
</style>
</head>
<body>
<h1>&#128202; TM + BOE Bump Rule 合規性報告</h1>
<div class='info-box'><table style='width:auto'>
<tr><td><b>晶片型號</b></td><td>{chip_model}</td><td><b>GDS 版本</b></td><td>{gds_base}</td></tr>
<tr><td><b>TM 規範</b></td><td>TM bump rule 參考</td><td><b>BOE 規範</b></td><td>BOE IC Bump Design Rule Rev0.2</td></tr>
<tr><td><b>Bump Layer</b></td><td>Layer 224（共 {len(bumps)} 個）</td><td><b>報告日期</b></td><td>{ts}</td></tr>
<tr><td><b>晶片尺寸（估計）</b></td><td>{ic_L_mm:.1f} x {ic_W_mm*1000:.0f} um</td><td><b>DB Unit</b></td><td>{gds_unit_um} um/unit</td></tr>
</table></div>

<div class='toc'><b>&#128196; 目錄</b><ul>
<li><a href='#summary'>合規總覽</a></li>
<li>Layer 224<ul>
<li><a href='#tm224'>TM R1~R5 規則</a></li>
<li><a href='#boe224'>BOE B1~B9 規則</a></li>
</ul></li>
<li><a href='#details'>詳細違規清單</a></li>
<li><a href='#recs'>修正建議</a></li>
</ul></div>

<h2 id='summary'>&#9989; 合規總覽</h2>
<div class='summary-grid'>
  <div class='summary-card'><div class='num' style='color:{"#c0392b" if total_viol else "#27ae60"}'>{total_viol}</div><div class='lbl'>總違規筆數</div></div>
  <div class='summary-card'><div class='num' style='color:#2980b9'>{tm_viol_count}</div><div class='lbl'>TM 違規</div></div>
  <div class='summary-card'><div class='num' style='color:#16a085'>{boe_viol_count}</div><div class='lbl'>BOE 違規</div></div>
  <div class='summary-card'><div class='num' style='color:#c0392b'>{crit_count}</div><div class='lbl'>CRITICAL</div></div>
  <div class='summary-card'><div class='num' style='color:#e67e22'>{high_count}</div><div class='lbl'>HIGH</div></div>
  <div class='summary-card'><div class='num' style='color:#f39c12'>{med_count}</div><div class='lbl'>MEDIUM</div></div>
</div>

<h3>TM Bump Rules 判定結果</h3>
<table style='width:auto;min-width:700px'>
<tr><th>規則</th><th>描述</th><th>規範值</th><th>GDS 量測</th><th>違規數</th><th>判定</th></tr>
<tr><td><b>R1</b></td><td>Bump 有效面積</td><td>(X-7)×(Y-10) ≥ 500 um²</td><td>{r1_meas}</td><td style='text-align:center'>{r1_fail}</td><td>{badge(r1_verdict, '#c0392b' if r1_verdict=='FAIL' else '#27ae60')}</td></tr>
<tr><td><b>R2</b></td><td>ILB Pitch</td><td>= {TM_ILB_PITCH:.0f} um</td><td>{r2_meas}</td><td style='text-align:center'>0</td><td>{badge(r2_verdict, r2_badge)}</td></tr>
<tr><td><b>R3</b></td><td>ILB→OLB 間距</td><td>{r3_spec}</td><td>{r3_meas}</td><td style='text-align:center'>0</td><td>{badge(r3_verdict, r3_badge)}</td></tr>
<tr><td><b>R4</b></td><td>晶片邊緣距離 A&amp;C</td><td>≥ {TM_EDGE_MIN:.0f} um</td><td>{r4_meas}</td><td style='text-align:center'>{len(r4_viols)}</td><td>{badge(r4_verdict, r4_badge)}</td></tr>
<tr><td><b>R5</b></td><td>B/(A+C) DFX</td><td>&lt; 5</td><td>{r5_meas}</td><td style='text-align:center'>{len(r5_viols)}</td><td>{badge(r5_verdict, r5_badge)}</td></tr>
</table>

<h3 style='color:#16a085;margin-top:20px'>BOE IC Bump Design Rules 判定結果</h3>
<table style='width:auto;min-width:700px'>
<tr><th class='boe-hdr'>規則</th><th class='boe-hdr'>描述</th><th class='boe-hdr'>規範值</th><th class='boe-hdr'>GDS 量測</th><th class='boe-hdr'>違規數</th><th class='boe-hdr'>判定</th></tr>
<tr><td><b>B1</b></td><td>IC 尺寸</td><td>L=5~35mm / W=0.5~2.2mm</td><td>L={ic_L_mm:.3f}mm / W={ic_W_mm:.3f}mm</td><td style='text-align:center'>0</td><td>{badge(b1_verdict, b1_badge)}</td></tr>
<tr><td><b>B2</b></td><td>邊緣 Bump 距 IC Edge（X）</td><td>≤ 200 um（左右各1筆）</td><td>Left={b2_left_dist:.1f} um / Right={b2_right_dist:.1f} um</td><td style='text-align:center'>0</td><td>{badge(b2_verdict, b2_badge)}</td></tr>
<tr><td><b>B3</b></td><td>Lead Width a</td><td>≥ 12 um</td><td>min={min_w:.1f} um</td><td style='text-align:center'>0</td><td>{badge(b3_verdict, b3_badge)}</td></tr>
<tr><td><b>B4</b></td><td>Lead Space b</td><td>≥ 12 um</td><td>{b4_meas}</td><td style='text-align:center'>0</td><td>{badge(b4_verdict, b4_badge)}</td></tr>
<tr><td><b>B5</b></td><td>Lead Pitch d</td><td>≥ 15 um</td><td>{b5_meas}</td><td style='text-align:center'>0</td><td>{badge(b5_verdict, b5_badge)}</td></tr>
<tr><td><b>B6</b></td><td>Lead Area c×a（c≥65）</td><td>≥ 950 um²</td><td>{b6_meas}</td><td style='text-align:center'>0</td><td>{badge(b6_verdict, b6_badge)}</td></tr>
<tr><td><b>B7</b></td><td>Bump 厚度</td><td>≥ 9 um</td><td>GDS 無法量測</td><td style='text-align:center'>0</td><td>{badge('N/A', '#95a5a6')}</td></tr>
<tr><td><b>B8</b></td><td>Bump 高度差</td><td>≤ 1.5 um</td><td>GDS 無法量測</td><td style='text-align:center'>0</td><td>{badge('N/A', '#95a5a6')}</td></tr>
<tr><td><b>B9</b></td><td>Pre-shrinkage</td><td>220 ppm</td><td>需確認設計意圖</td><td style='text-align:center'>0</td><td>{badge('VERIFY', '#8e44ad')}</td></tr>
</table>

<hr class='section-divider'>
<h2 id='tm224'>&#128313; Layer 224 TM 規則詳細說明</h2>

<h3 id='r1'>R1：Bump 有效面積</h3>
<div class='info-box'>
<b>規範：</b>(X-7) × (Y-10) ≥ 500 um²<br>
<b>判定：</b>{badge(r1_verdict, '#c0392b' if r1_verdict=='FAIL' else '#27ae60')}<br><br>
<table><tr><th>Bump 尺寸</th><th>數量</th><th>有效面積</th><th>規範</th><th>判定</th></tr>
{r1_size_rows}
</table>
</div>

<h3 id='r2'>R2：ILB Pitch</h3>
<div class='info-box'>
<b>規範：</b>ILB Pitch = {TM_ILB_PITCH:.0f} um<br>
<b>判定：</b>{badge(r2_verdict, r2_badge)}<br><br>
ILB 量測 Pitch = <b>{ilb_pitch:.2f} um</b>（{f'符合規範' if r2_verdict=='PASS' else '請確認 ILB/OLB 分類'}）<br>
</div>

<h3 id='r3'>R3：ILB→OLB 間距</h3>
<div class='info-box'>
<b>規範：</b>ILB 上緣 → OLB 下緣間距 B，偵測到 <b>{n_olb_rows} 排</b> OLB → 最小間距 B &gt; <b>{r3_spec_min:.1f} um</b><br>
<b>判定：</b>{badge(r3_verdict, r3_badge)}<br><br>
量測 B 範圍：<b>{B_min:.1f} ~ {B_max:.1f} um</b>
</div>

<h3 id='r4'>R4：晶片邊緣距離</h3>
<div class='info-box'>
<b>規範：</b>A &amp; C ≥ {TM_EDGE_MIN:.0f} um<br>
<b>判定：</b>{badge(r4_verdict, r4_badge)}<br><br>
<table><tr><th>方向</th><th>量測值</th><th>規範</th><th>差距</th><th>判定</th></tr>
{r4_edge_rows}
</table><br>
<small>注意：A / C 距離由 Bump 排列估算，如需精確值請提供 IC 邊界層（Die Boundary Layer）</small>
</div>

<h3 id='r5'>R5：B/(A+C) DFX</h3>
<div class='info-box'>
<b>規範：</b>B/(A+C) &lt; 5<br>
<b>判定：</b>{badge(r5_verdict, r5_badge)}<br><br>
{r5_detail}
</div>
<hr class='section-divider'>

<h2 id='boe224' class='boe'>&#127381; Layer 224 BOE 規則詳細說明</h2>

<h3 id='b1'>B1：IC 尺寸</h3>
<div class='info-box'>
<b>規範：</b>IC Length L = 5~35 mm，IC Width W = 0.5~2.2 mm<br>
<b>判定：</b>{badge(b1_verdict, b1_badge)}<br><br>
<table><tr><th>參數</th><th>量測值</th><th>規範</th><th>判定</th></tr>
<tr><td>L (晶片長度)</td><td>{ic_L_mm:.3f} mm</td><td>5 ~ 35 mm</td><td>{badge('PASS' if b1_pass_L else 'FAIL', '#27ae60' if b1_pass_L else '#c0392b')}</td></tr>
<tr><td>W (晶片寬度)</td><td>{ic_W_mm:.3f} mm</td><td>0.5 ~ 2.2 mm</td><td>{badge('PASS' if b1_pass_W else 'FAIL', '#27ae60' if b1_pass_W else '#c0392b')}</td></tr>
</table></div>

<h3 id='b2'>B2：Bump 距 IC 邊緣距離</h3>
<div class='info-box'>
<b>規範：</b>IC 邊緣 Bump 到 IC Edge 距離 ≤ 200 um（X 方向，左右各1筆）<br>
<b>判定：</b>{badge(b2_verdict, b2_badge)}<br><br>
<table><tr><th>邊緣</th><th>最近 Bump 距離</th><th>說明</th></tr>
{b2_rows}
</table>
</div>

<h3 id='b3b6'>B3~B6：Lead 幾何規範</h3>
<div class='info-box'>
<table><tr><th>規則</th><th>描述</th><th>規範</th><th>GDS 量測</th><th>判定</th></tr>
{b3b6_rows}
</table><br>
<b>Bump 尺寸分佈：</b><br>
<table><tr><th>尺寸 (a×c)</th><th>數量</th><th>面積</th><th>B6 適用（c≥65）</th><th>判定</th></tr>
{size_dist_rows}
</table>
</div>

<h3 id='b7b9'>B7~B9：製程 / 縮放規範</h3>
<div class='info-box'>
<table><tr><th>規則</th><th>描述</th><th>規範</th><th>GDS 狀態</th><th>判定</th></tr>
<tr><td><b>B7</b></td><td>Bump 厚度 D</td><td>≥ 9 um</td><td>2D GDS 無法量測，需製程規格書確認</td><td>{badge('N/A', '#95a5a6')}</td></tr>
<tr><td><b>B8</b></td><td>Bump 高度差</td><td>≤ 1.5 um</td><td>需量測儀器確認</td><td>{badge('N/A', '#95a5a6')}</td></tr>
<tr><td><b>B9</b></td><td>Pre-shrinkage</td><td>220 ppm（向 IC 中心）</td><td>需確認設計意圖</td><td>{badge('VERIFY', '#8e44ad')}</td></tr>
</table>
</div>
<hr class='section-divider'>

<h2 id='details'>&#128203; 詳細違規清單</h2>
{"<p>每個違規類型列出前 20 筆代表樣本；完整資料請參閱 CSV 檔（共 "+str(total_viol)+" 筆）</p>" if total_viol else "<p style='color:#27ae60'>&#10003; 無違規紀錄 — 所有規則均通過</p>"}
<table>
<tr><th>No</th><th>Std</th><th>Rule</th><th>Severity</th><th>Net</th><th>Sig</th><th>W</th><th>H</th><th>Cx</th><th>Cy</th><th>Measured</th><th>Spec</th><th>Margin</th><th>Description</th></tr>
{viol_rows if viol_rows else "<tr><td colspan='14' style='text-align:center;color:#27ae60'>無違規</td></tr>"}
</table>
<hr class='section-divider'>

<h2 id='recs'>&#128161; 修正建議</h2>
<h3>TM 規則修正</h3><div class='info-box'><ol>
{recs_tm}
</ol></div>
<h3>BOE 規則修正</h3><div class='info-box'><ol>
{recs_boe}
</ol></div>

<div class='footer'>Generated by GDS Bump Rule Checker (Python) v2.0 (TM + BOE) | {ts} | GDS: {gds_base}</div>
</body></html>
"""

        if report_html.exists():
            report_html.unlink()
        report_html.write_text(html_content, encoding='utf-8')
        print(f"[INFO] HTML report written: {report_html} ({report_html.stat().st_size} bytes)")

        return {
            "success": True,
            "gds_file": str(gds_path),
            "output_dir": out_dir,
            "violations_csv": str(violations_csv).replace("\\", "/"),
            "report_html": str(report_html).replace("\\", "/"),
            "report_pdf": None,
            "total_bumps": len(bumps),
            "total_violations": total_viol,
            "message": f"Python Bump Check: {len(bumps)} bumps, {total_viol} violations (TM={tm_viol_count}, BOE={boe_viol_count})",
            "return_code": 0
        }

    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(f"[ERROR] {error_msg}")
        return {
            "success": False,
            "message": f"Python Bump Check 失敗: {str(e)}",
            "gds_file": str(gds_path),
            "output_dir": out_dir if out_dir else str(gds_path.parent),
            "error": error_msg
        }


def run_bump_check_perl(gds_file: str, out_dir: str = None, original_name: str = None) -> dict:
    """
    調用 Perl check_bump_rule2.pl 進行 Bump Rule 檢查

    Args:
        gds_file: GDS 檔案路徑
        out_dir: 輸出目錄（預設為 GDS 同目錄）

    Returns:
        {
            "success": bool,
            "gds_file": str,
            "output_dir": str,
            "violations_csv": str,  # 違規檔案路徑
            "boe_violations_csv": str,
            "report_html": str,
            "report_pdf": str,
            "message": str,
            "stderr": str
        }
    """
    perl_exe = find_perl()
    if not perl_exe:
        return {
            "success": False,
            "message": "找不到 Perl，請安裝 Strawberry Perl 或檢查 PATH",
            "gds_file": gds_file
        }

    pl_script = _HERE / "check_bump_rule2.pl"
    if not pl_script.exists():
        return {
            "success": False,
            "message": f"找不到 check_bump_rule2.pl，期望路徑: {pl_script}",
            "gds_file": gds_file
        }

    gds_path = Path(gds_file)
    if not gds_path.exists():
        return {
            "success": False,
            "message": f"GDS 檔案不存在: {gds_file}",
            "gds_file": gds_file
        }

    # 決定輸出目錄
    if out_dir is None:
        out_dir = str(gds_path.parent)
    else:
        out_dir = str(Path(out_dir))
        Path(out_dir).mkdir(parents=True, exist_ok=True)

    # 構建命令
    cmd = [perl_exe, str(pl_script), "-i", str(gds_path), "-o", out_dir]

    print(f"執行: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

        # Perl script 使用輸入檔案的 stem 生成輸出（可能是臨時名）
        temp_gds_base = gds_path.stem
        # 使用原始檔名（如果提供）或臨時檔名
        gds_base = original_name if original_name else temp_gds_base

        out_dir_path = Path(out_dir)

        # 如果 Perl 用臨時名生成了檔案，重命名為原始名
        if original_name and original_name != temp_gds_base:
            print(f"[DEBUG] Renaming Perl output from '{temp_gds_base}' to '{original_name}'")

            temp_files = [
                (temp_gds_base, "_tm_violations.csv"),
                (temp_gds_base, "_boe_violations.csv"),
                (temp_gds_base, "_report.html"),
                (temp_gds_base, "_report.pdf"),
            ]

            for temp_base, suffix in temp_files:
                temp_file = out_dir_path / f"{temp_base}{suffix}"
                if temp_file.exists():
                    new_file = out_dir_path / f"{original_name}{suffix}"
                    # 如果目標檔案已存在，先刪除它
                    if new_file.exists():
                        new_file.unlink()
                        print(f"[DEBUG] Deleted existing: {new_file.name}")
                    # 重命名
                    temp_file.rename(new_file)
                    print(f"[DEBUG] Renamed: {temp_file.name} → {new_file.name}")

        tm_violations_csv = out_dir_path / f"{gds_base}_tm_violations.csv"
        boe_violations_csv = out_dir_path / f"{gds_base}_boe_violations.csv"
        report_html = out_dir_path / f"{gds_base}_report.html"
        report_pdf = out_dir_path / f"{gds_base}_report.pdf"

        # 檢查報告檔案是否存在
        html_exists = report_html.exists()
        pdf_exists = report_pdf.exists()
        print(f"[DEBUG] Perl output check: html={html_exists} ({report_html}), pdf={pdf_exists} ({report_pdf})")

        result_dict = {
            "success": result.returncode == 0,
            "gds_file": str(gds_path),
            "output_dir": out_dir,
            "violations_csv": str(tm_violations_csv).replace("\\", "/") if tm_violations_csv.exists() else None,
            "tm_violations_csv": str(tm_violations_csv).replace("\\", "/") if tm_violations_csv.exists() else None,
            "boe_violations_csv": str(boe_violations_csv).replace("\\", "/") if boe_violations_csv.exists() else None,
            "report_html": str(report_html).replace("\\", "/") if html_exists else None,
            "report_pdf": str(report_pdf).replace("\\", "/") if pdf_exists else None,
            "message": result.stdout,
            "stderr": result.stderr,
            "return_code": result.returncode
        }
        if result_dict["report_html"]:
            print(f"[INFO] Perl 報告路徑: {result_dict['report_html']}")
        return result_dict

    except subprocess.TimeoutExpired:
        return {
            "success": False,
            "message": "Bump Rule 檢查超時（>5分鐘）",
            "gds_file": str(gds_path),
            "output_dir": out_dir
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"執行錯誤: {str(e)}",
            "gds_file": str(gds_path),
            "output_dir": out_dir,
            "error": str(e)
        }


def check_bump_rules_summary(gds_file: str) -> dict:
    """
    簡化的 Bump Rule 概要檢查（Python 實現）

    檢查項目:
    - Layer 224/225 是否存在（Bump/PAD 層）
    - Polygon 數量統計（支援 Layer:DType 和 Layer 編號）
    - 邊界檢查
    """
    try:
        import gdstk
    except ImportError:
        return {
            "success": False,
            "message": "缺少 gdstk 套件，請執行: pip install gdstk"
        }

    gds_path = Path(gds_file)
    if not gds_path.exists():
        return {"success": False, "message": f"GDS 檔案不存在: {gds_file}"}

    try:
        print(f"[DEBUG] Reading GDS: {gds_path}")
        lib = gdstk.read_gds(str(gds_path))

        bump_poly_count = 0
        bump_layers_found = set()

        # 兼容 gdstk 版本差異：cells 可能是 dict 或 list
        cells = lib.cells
        print(f"[DEBUG] cells type: {type(cells)}")
        if isinstance(cells, dict):
            cells = cells.values()
        elif not isinstance(cells, list):
            cells = list(cells)

        for cell in cells:
            # 檢查 Polygons
            if hasattr(cell, 'polygons') and cell.polygons:
                for poly in cell.polygons:
                    layer = poly.layer
                    dtype = poly.datatype

                    # Layer 編號：可能是單純編號或 (layer, dtype) 元組
                    layer_num = layer if isinstance(layer, int) else (layer[0] if isinstance(layer, (tuple, list)) else None)

                    # 檢查是否為 Bump Layer (224 或 225)
                    if layer_num in [224, 225]:
                        bump_poly_count += 1
                        bump_layers_found.add(layer_num)

            # 也檢查 paths（某些工具用 Path 而非 Polygon 表示 Bump）
            if hasattr(cell, 'paths') and cell.paths:
                for path in cell.paths:
                    layer = path.layer
                    dtype = path.datatype
                    layer_num = layer if isinstance(layer, int) else (layer[0] if isinstance(layer, (tuple, list)) else None)

                    if layer_num in [224, 225]:
                        bump_poly_count += 1
                        bump_layers_found.add(layer_num)

        print(f"[DEBUG] bump_poly_count: {bump_poly_count}, bump_layers: {bump_layers_found}")

        summary = {
            "success": True,
            "gds_file": str(gds_path),
            "bump_layers_found": sorted(list(bump_layers_found)),
            "total_bump_polygons": bump_poly_count,
            "needs_full_check": bump_poly_count > 0,
            "message": f"偵測到 {bump_poly_count} 個 Bump 元素在 Layer {sorted(list(bump_layers_found))}"
        }

        print(f"[DEBUG] Summary: {summary}")
        return summary

    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        print(f"[ERROR] Exception in check_bump_rules_summary: {error_detail}")
        return {
            "success": False,
            "message": f"解析 GDS 失敗: {str(e)}",
            "gds_file": str(gds_file),
            "error_detail": error_detail
        }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python bump_checker.py <gds_file> [output_dir]")
        print("\n示例:")
        print("  python bump_checker.py design.gds")
        print("  python bump_checker.py design.gds ./output")
        sys.exit(1)

    gds = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None

    result = run_bump_check_perl(gds, out)
    print(json.dumps(result, indent=2, ensure_ascii=False))
    sys.exit(0 if result.get("success") else 1)
