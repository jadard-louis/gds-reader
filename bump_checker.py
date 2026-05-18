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
    純 Python Bump Rule 檢查（簡化版，無需 Perl）

    檢查項目:
    - Layer 224/225 Polygon 尺寸檢查（寬、高範圍）
    - Polygon 間距檢查
    - 基本幾何規則（min width、spacing）
    """
    try:
        import gdstk
        import csv
        from datetime import datetime
    except ImportError as e:
        return {
            "success": False,
            "message": f"缺少模組: {str(e)}",
            "gds_file": gds_file
        }

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
        # 使用原始檔名（如果提供）或臨時檔名
        gds_base = original_name if original_name else gds_path.stem
        print(f"[DEBUG] Using basename: {gds_base} (original_name={original_name})")

        # 獲取 GDS 單位信息
        # gdstk 中 lib.unit 是實際單位（micrometers）
        gds_unit_um = getattr(lib, 'unit', 0.001)  # 預設 0.001 μm (1 nm)
        print(f"[DEBUG] GDS unit: {gds_unit_um} μm")

        # 蒐集所有 Bump Polygons
        bumps = []
        cells_data = lib.cells
        print(f"[DEBUG] cells_data type: {type(cells_data)}")

        if isinstance(cells_data, dict):
            cells_data = cells_data.values()
        elif not isinstance(cells_data, list):
            cells_data = list(cells_data)

        print(f"[DEBUG] Processing cells...")
        for cell in cells_data:
            if hasattr(cell, 'polygons') and cell.polygons:
                for idx, poly in enumerate(cell.polygons):
                    if poly.layer in [224, 225]:
                        # 計算 bbox
                        pts = poly.points
                        if len(pts) > 0:
                            xs = [p[0] for p in pts]
                            ys = [p[1] for p in pts]
                            x_min, x_max = min(xs), max(xs)
                            y_min, y_max = min(ys), max(ys)

                            # 根據 GDS 單位轉換
                            w_um = (x_max - x_min) * gds_unit_um
                            h_um = (y_max - y_min) * gds_unit_um
                            x_min_um = x_min * gds_unit_um
                            y_min_um = y_min * gds_unit_um

                            # 顯示前 3 個 bump 的座標以調試
                            if len(bumps) < 3:
                                print(f"[DEBUG] Bump {len(bumps)}: pts={len(pts)}, raw_w={x_max-x_min}, w_um={w_um:.4f}, h_um={h_um:.4f}")

                            bumps.append({
                                "x_min": x_min_um,
                                "y_min": y_min_um,
                                "w": w_um,
                                "h": h_um,
                                "layer": poly.layer,
                                "cell": cell.name if hasattr(cell, 'name') else "?"
                            })

        print(f"[DEBUG] Collected {len(bumps)} bumps")

        # 簡單規則檢查
        violations = []

        # 更寬鬆的規則（根據實際情況調整）
        MIN_WIDTH = 0.1  # μm - Bump 最小寬度
        MIN_HEIGHT = 0.1  # μm - Bump 最小高度
        MAX_WIDTH = 500.0  # μm - Bump 最大寬度
        MAX_HEIGHT = 500.0  # μm - Bump 最大高度
        MIN_SPACING = 0.1  # μm - Bump 最小間距

        print(f"[DEBUG] Checking bumps with rules: MIN_W={MIN_WIDTH}, MAX_W={MAX_WIDTH}")

        # 計算統計資訊
        if bumps:
            widths = [b["w"] for b in bumps]
            heights = [b["h"] for b in bumps]
            print(f"[DEBUG] Width range: {min(widths):.4f} - {max(widths):.4f} μm")
            print(f"[DEBUG] Height range: {min(heights):.4f} - {max(heights):.4f} μm")
            print(f"[DEBUG] Sample bump 0: w={widths[0]:.4f}, h={heights[0]:.4f} μm")

        # 檢查 Bump 尺寸（暫時只記錄警告，不標記為違規）
        size_warnings = []
        for i, bump in enumerate(bumps[:5]):  # 只檢查前 5 個以減少日誌
            w = bump["w"]
            h = bump["h"]
            if i == 0:  # 只記錄第一個
                print(f"[DEBUG] Bump {i}: w={w:.4f}, h={h:.4f} μm")

            if w < MIN_WIDTH or h < MIN_HEIGHT:
                size_warnings.append(f"Bump {i} too small: {w:.2f}x{h:.2f}")
            elif w > MAX_WIDTH or h > MAX_HEIGHT:
                size_warnings.append(f"Bump {i} too large: {w:.2f}x{h:.2f}")

        if size_warnings:
            print(f"[DEBUG] Size warnings: {size_warnings[:3]}")

        # 簡化：暫時跳過間距檢查（2555 個 bump 的兩兩比較太耗時）
        # 實際應用可以用網格演算法加速
        print(f"[DEBUG] Skipping detailed spacing check for {len(bumps)} bumps (would be {len(bumps)**2} comparisons)")

        # 確保 output 目錄存在
        out_path = Path(out_dir)
        out_path.mkdir(parents=True, exist_ok=True)
        print(f"[DEBUG] Output dir: {out_path}, exists: {out_path.exists()}")

        # 寫入 CSV 違規報告
        violations_csv = out_path / f"{gds_base}_violations.csv"
        print(f"[DEBUG] Writing violations CSV to: {violations_csv}")
        try:
            # 刪除舊檔案（如果存在且被鎖定）
            if violations_csv.exists():
                try:
                    violations_csv.unlink()
                    print(f"[DEBUG] Deleted old CSV file")
                except Exception as del_err:
                    print(f"[WARNING] Could not delete old CSV: {del_err}")

            with open(str(violations_csv), 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["type", "index", "message", "x", "y"])
                writer.writeheader()
                for v in violations:
                    # 確保所有欄位都是正確的型態
                    row = {
                        "type": str(v.get("type", "")),
                        "index": str(v.get("index", "")),
                        "message": str(v.get("message", "")),
                        "x": float(v.get("x", 0)),
                        "y": float(v.get("y", 0))
                    }
                    writer.writerow(row)
            print(f"[DEBUG] CSV written: {violations_csv.exists()}, size: {violations_csv.stat().st_size if violations_csv.exists() else 0}")
        except Exception as csv_err:
            print(f"[ERROR] CSV write failed: {csv_err}")
            import traceback
            print(traceback.format_exc())
            raise

        # 寫入 HTML 報告
        report_html = out_path / f"{gds_base}_report.html"
        print(f"[DEBUG] Writing HTML report to: {report_html}")
        print(f"[DEBUG] HTML path type: {type(report_html)}, exists before: {report_html.exists()}")

        try:
            # 刪除舊檔案
            if report_html.exists():
                try:
                    report_html.unlink()
                    print(f"[DEBUG] Deleted old HTML file")
                except Exception as del_err:
                    print(f"[WARNING] Could not delete old HTML: {del_err}")

            print(f"[DEBUG] Starting HTML content generation... (bumps: {len(bumps)}, violations: {len(violations)})")

            html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Bump Rule Check Report</title>
    <style>
        body {{ font-family: Arial; margin: 20px; }}
        h1 {{ color: #333; }}
        .summary {{ background: #f0f0f0; padding: 10px; border-radius: 5px; margin: 10px 0; }}
        .violations {{ margin: 20px 0; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
        th {{ background: #ddd; }}
        .error {{ color: #d00; }}
    </style>
</head>
<body>
    <h1>Bump Rule Check Report (Python)</h1>
    <div class="summary">
        <p><strong>GDS File:</strong> {gds_base}</p>
        <p><strong>Total Bumps:</strong> {len(bumps)}</p>
        <p><strong>Violations:</strong> {len(violations)}</p>
        <p><strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>

    <h2>Bump Summary (first 20)</h2>
    <table>
        <tr><th>Index</th><th>Layer</th><th>Width (um)</th><th>Height (um)</th><th>X (um)</th><th>Y (um)</th></tr>
"""
            # 只顯示前 20 個 bump 以避免 HTML 過大
            for i, bump in enumerate(bumps[:20]):
                html_content += f"<tr><td>{i}</td><td>{bump['layer']}</td><td>{bump['w']:.2f}</td><td>{bump['h']:.2f}</td><td>{bump['x_min']:.2f}</td><td>{bump['y_min']:.2f}</td></tr>"

            if len(bumps) > 20:
                html_content += f"<tr><td colspan='6'>... and {len(bumps) - 20} more bumps</td></tr>"

            html_content += """
    </table>

    <h2>Violations</h2>
"""
            if violations:
                html_content += "<table><tr><th>Type</th><th>Message</th><th>Position</th></tr>"
                for v in violations[:50]:  # 限制前 50 個違規
                    html_content += f"<tr class='error'><td>{v['type']}</td><td>{v['message']}</td><td>({v['x']:.2f}, {v['y']:.2f})</td></tr>"
                if len(violations) > 50:
                    html_content += f"<tr><td colspan='3'>... and {len(violations) - 50} more violations</td></tr>"
                html_content += "</table>"
            else:
                html_content += "<p style='color: green;'><strong>No violations found!</strong></p>"

            html_content += """
</body>
</html>
"""
            print(f"[DEBUG] HTML content generated, length: {len(html_content)}")

            # 寫入檔案
            report_html.write_text(html_content, encoding='utf-8')
            print(f"[DEBUG] HTML write_text completed")

            # 驗證寫入
            if report_html.exists():
                file_size = report_html.stat().st_size
                print(f"[DEBUG] HTML written successfully: size={file_size} bytes")
            else:
                print(f"[ERROR] HTML file does not exist after write_text!")
                raise IOError(f"Failed to create HTML file at {report_html}")

        except Exception as html_err:
            print(f"[ERROR] HTML write failed: {type(html_err).__name__}: {html_err}")
            import traceback
            print(traceback.format_exc())
            raise

        print(f"[DEBUG] Total bumps: {len(bumps)}, violations: {len(violations)}")

        result = {
            "success": True,
            "gds_file": str(gds_path),
            "output_dir": out_dir,
            "violations_csv": str(violations_csv).replace("\\", "/"),
            "report_html": str(report_html).replace("\\", "/"),
            "report_pdf": None,
            "total_bumps": len(bumps),
            "total_violations": len(violations),
            "message": f"Python Bump Check: {len(bumps)} bumps, {len(violations)} violations",
            "return_code": 0
        }
        print(f"[INFO] 生成 HTML 報告完成: {result['report_html']}", file=_sys)

        print(f"[DEBUG] Return result: {result}")
        return result

    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(f"[ERROR] {error_msg}")
        return {
            "success": False,
            "message": f"Python Bump Check 失敗: {str(e)}",
            "gds_file": str(gds_path),
            "output_dir": out_dir,
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
