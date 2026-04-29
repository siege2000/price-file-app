import pyodbc
from collections import Counter
import html as html_module

base = r"C:\Users\Colin\OneDrive - Healthsoft Ltd\Documents\github\price-file-app\tests"
new_path = base + r"\Suppliers - Copy.mdb"
legacy_path = base + r"\Suppliers.mdb"
password = "LOCKIE MONDAY"


def connect_mdb(path):
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={path};"
        f"PWD={password};"
    )
    return pyodbc.connect(conn_str)


def load_items(path):
    conn = connect_mdb(path)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM SpecialItems WHERE SpecialID = 1237 ORDER BY Barcode")
    cols = [d[0] for d in cursor.description]
    rows = {row[2]: dict(zip(cols, row)) for row in cursor.fetchall()}
    conn.close()
    return cols, rows


_, new_rows = load_items(new_path)
_, legacy_rows = load_items(legacy_path)

# Retail diffs
retail_diffs = []
for bc, r in new_rows.items():
    leg = legacy_rows.get(bc)
    if leg and str(r["SpecialRetail"]) != str(leg["SpecialRetail"]):
        retail_diffs.append({
            "barcode": bc, "desc": r["Description"],
            "deal_new": r["DealName"], "deal_leg": leg["DealName"],
            "retail_new": r["SpecialRetail"], "retail_leg": leg["SpecialRetail"],
        })

# GiftPharmacode
gp_new = Counter(r["GiftPharmacode"] for r in new_rows.values())
gp_leg = Counter(r["GiftPharmacode"] for r in legacy_rows.values())

# DealName diffs
deal_diffs = [
    (bc, legacy_rows[bc]["DealName"], r["DealName"])
    for bc, r in new_rows.items()
    if r["DealName"] != legacy_rows[bc]["DealName"]
]
deal_pairs = Counter((lv or "(empty)", nv) for _, lv, nv in deal_diffs)

# SpecialNote diffs
note_added = sum(
    1 for bc, r in new_rows.items()
    if r["SpecialNote"] and not legacy_rows[bc]["SpecialNote"]
)
note_changed = sum(
    1 for bc, r in new_rows.items()
    if r["SpecialNote"] and legacy_rows[bc]["SpecialNote"]
    and r["SpecialNote"] != legacy_rows[bc]["SpecialNote"]
)

# NULL retail by deal
null_samples = [(bc, r) for bc, r in new_rows.items() if r["SpecialRetail"] is None]
null_by_deal = Counter(r["DealName"] for _, r in null_samples)

h = html_module.escape

rows_table_null = ""
for r in retail_diffs[:10]:
    bc = r["barcode"] or "(empty)"
    rows_table_null += (
        f"      <tr><td>{h(str(bc))}</td><td>{h(r['desc'])}</td>"
        f"<td>{h(r['deal_new'])}</td>"
        f"<td>{r['retail_leg']}</td>"
        f"<td><span class='badge badge-red'>NULL</span></td></tr>\n"
    )

rows_null_by_deal = ""
for deal, cnt in null_by_deal.most_common():
    rows_null_by_deal += f"      <tr><td>{h(deal)}</td><td>{cnt}</td><td>POS message</td></tr>\n"

rows_deal_pairs = ""
for (leg, nw), cnt in deal_pairs.most_common(15):
    rows_deal_pairs += f"      <tr><td>{h(nw)}</td><td>{cnt}</td><td><em>{h(leg)}</em></td></tr>\n"

report = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>SpecialItems Comparison - SpecialID 1237</title>
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background: #f4f6f9; color: #1a1a2e; }}
  .container {{ max-width: 1100px; margin: 0 auto; padding: 32px 24px; }}
  h1 {{ font-size: 1.6rem; color: #1a1a2e; border-bottom: 3px solid #2563eb; padding-bottom: 10px; margin-bottom: 6px; }}
  .meta {{ color: #64748b; font-size: 0.85rem; margin-bottom: 32px; }}
  h2 {{ font-size: 1.15rem; color: #1e3a8a; margin-top: 0; margin-bottom: 8px; border-left: 4px solid #2563eb; padding-left: 10px; }}
  .summary-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 28px; }}
  .card {{ background: white; border-radius: 8px; padding: 18px 20px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  .card .num {{ font-size: 2rem; font-weight: 700; color: #2563eb; }}
  .card .num.warn {{ color: #dc2626; }}
  .card .num.ok {{ color: #16a34a; }}
  .card .lbl {{ font-size: 0.82rem; color: #64748b; margin-top: 2px; }}
  .card .sub {{ font-size: 0.78rem; color: #94a3b8; margin-top: 4px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.82rem; margin-top: 10px; }}
  th {{ background: #1e3a8a; color: white; padding: 8px 10px; text-align: left; }}
  td {{ padding: 7px 10px; border-bottom: 1px solid #e2e8f0; }}
  tr:nth-child(even) td {{ background: #f8fafc; }}
  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 0.72rem; font-weight: 600; }}
  .badge-red {{ background: #fee2e2; color: #b91c1c; }}
  .badge-blue {{ background: #dbeafe; color: #1d4ed8; }}
  .badge-green {{ background: #dcfce7; color: #15803d; }}
  .badge-amber {{ background: #fef9c3; color: #92400e; }}
  .bug-box {{ background: #fff7ed; border: 1px solid #fb923c; border-radius: 8px; padding: 16px 20px; margin-top: 12px; }}
  .bug-box h3 {{ margin: 0 0 8px 0; color: #c2410c; font-size: 0.95rem; }}
  .bug-box code {{ background: #fef3c7; padding: 1px 5px; border-radius: 3px; font-size: 0.82rem; }}
  .fix-box {{ background: #f0fdf4; border: 1px solid #4ade80; border-radius: 8px; padding: 16px 20px; margin-top: 8px; }}
  .fix-box h3 {{ margin: 0 0 8px 0; color: #166534; font-size: 0.95rem; }}
  .fix-box code {{ background: #dcfce7; padding: 1px 5px; border-radius: 3px; font-size: 0.82rem; }}
  .info-box {{ background: #eff6ff; border: 1px solid #93c5fd; border-radius: 8px; padding: 14px 18px; margin-top: 12px; font-size: 0.85rem; }}
  pre {{ font-family: Consolas, monospace; font-size: 0.8rem; background: #1e293b; color: #e2e8f0; padding: 12px 16px; border-radius: 6px; overflow-x: auto; margin: 8px 0; }}
  .section {{ background: white; border-radius: 8px; padding: 20px 24px; box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-top: 20px; }}
</style>
</head>
<body>
<div class="container">
  <h1>SpecialItems Comparison &mdash; SpecialID 1237</h1>
  <div class="meta">
    Source: DM02 2026 as at 20042026 - Final (RxOne).xlsx &nbsp;|&nbsp;
    New app: Suppliers - Copy.mdb &nbsp;|&nbsp; Legacy app: Suppliers.mdb &nbsp;|&nbsp;
    Compared: 28 April 2026
  </div>

  <div class="summary-grid">
    <div class="card">
      <div class="num">4,987</div>
      <div class="lbl">Rows in both databases</div>
      <div class="sub">Same barcodes, no missing or extra rows</div>
    </div>
    <div class="card">
      <div class="num warn">1,032</div>
      <div class="lbl">SpecialRetail: NULL vs legacy value</div>
      <div class="sub">New app stores NULL; legacy stored -1 or 0</div>
    </div>
    <div class="card">
      <div class="num warn">3,402</div>
      <div class="lbl">GiftPharmacode: 'False' vs ''</div>
      <div class="sub">New app writes string 'False' instead of empty</div>
    </div>
    <div class="card">
      <div class="num ok">1,586</div>
      <div class="lbl">DealName: blank to populated</div>
      <div class="sub">New feature; legacy never wrote DealName</div>
    </div>
    <div class="card">
      <div class="num ok">2,658</div>
      <div class="lbl">SpecialNote: blank to populated</div>
      <div class="sub">New feature; legacy left notes blank</div>
    </div>
    <div class="card">
      <div class="num warn">1</div>
      <div class="lbl">Empty-barcode row collision</div>
      <div class="sub">Different item assigned to barcode ''</div>
    </div>
  </div>

  <div class="section">
    <h2>Bug 1 &mdash; SpecialRetail stores NULL for no-price items</h2>
    <div class="bug-box">
      <h3>&#x26A0; Problem</h3>
      <p>For POS-message deal types (GWP / prize draws) the source xlsx has no <em>Product price</em>.
      The new app's <code>_parse_special_price()</code> returns <code>(None, 0.0)</code> for an empty price,
      writing NULL to <code>SpecialRetail</code>. The legacy app stored <code>-1</code> as a sentinel
      meaning "no fixed price".</p>
    </div>
    <div class="fix-box">
      <h3>&#x2714; Fix &mdash; specials/specials_helpers.py, _parse_special_price()</h3>
      <pre>if price_raw == "":
    return -1, 0.0  # was: return None, 0.0</pre>
      <p>This aligns the new app with the legacy sentinel value of <code>-1</code> for price-free specials.</p>
    </div>

    <h2 style="margin-top:20px;">Breakdown by deal</h2>
    <table>
      <tr><th>Deal Name</th><th>Rows affected</th><th>Deal Type</th></tr>
{rows_null_by_deal}    </table>

    <h2 style="margin-top:20px;">Sample affected rows (first 10)</h2>
    <table>
      <tr><th>Barcode</th><th>Description</th><th>Deal (New)</th><th>Retail (Legacy)</th><th>Retail (New)</th></tr>
{rows_table_null}    </table>
  </div>

  <div class="section">
    <h2>Bug 2 &mdash; GiftPharmacode stores 'False' instead of empty string</h2>
    <div class="bug-box">
      <h3>&#x26A0; Problem</h3>
      <p>Line 181 of <code>specials/specials_helpers.py</code>:</p>
      <pre>gift_pharmacode = str(r.get("Gift Pharmacode", "") or "").strip() or "False"</pre>
      <p>When the xlsx has no <em>Gift Pharmacode</em> column (which is the case for DM02),
      <code>r.get("Gift Pharmacode", "")</code> returns <code>""</code>.
      The expression <code>"" or "False"</code> then evaluates to the string <code>"False"</code>
      rather than an empty string. This affects every single row.</p>
    </div>
    <div class="fix-box">
      <h3>&#x2714; Fix &mdash; remove the fallback default</h3>
      <pre>gift_pharmacode = str(r.get("Gift Pharmacode", "") or "").strip()</pre>
      <p>Remove the trailing <code>or "False"</code>. An empty string is the correct value when no pharmacode is present.</p>
    </div>
    <table style="margin-top:14px;">
      <tr><th>GiftPharmacode value</th><th>New app rows</th><th>Legacy rows</th></tr>
      <tr><td>'' (empty string)</td><td>0</td><td>3,402</td></tr>
      <tr><td>'False' (string)</td><td><span class='badge badge-red'>4,987 (all rows)</span></td><td>1,585</td></tr>
    </table>
  </div>

  <div class="section">
    <h2>Feature improvement &mdash; DealName now populated (1,586 rows)</h2>
    <div class="info-box">
      The legacy app never wrote to <code>DealName</code> &mdash; all rows were blank.
      The new app correctly populates it from the xlsx <em>Deal name</em> column.
      This is not a bug; it is new functionality.
    </div>
    <table style="margin-top:12px;">
      <tr><th>New DealName</th><th>Rows</th><th>Legacy DealName</th></tr>
{rows_deal_pairs}    </table>
  </div>

  <div class="section">
    <h2>Feature improvement &mdash; SpecialNote now populated (2,658 rows)</h2>
    <div class="info-box">
      The legacy app left <code>SpecialNote</code> blank for all rows in this special.
      The new app writes it from the xlsx <em>Promotion POS note</em> column.
      This is correct and expected new behaviour.
    </div>
    <p style="margin-top:12px;font-size:0.85rem;">
      {note_added} rows gained a note &nbsp;|&nbsp; {note_changed} rows have a different note from legacy.
    </p>
  </div>

  <div class="section">
    <h2>Data issue &mdash; Empty-barcode row collision</h2>
    <div class="bug-box">
      <h3>&#x26A0; Problem</h3>
      <p>Both databases contain exactly one row with an empty <code>Barcode</code> field, but they
      hold completely different items. Multiple xlsx rows have no barcode value; only the last one
      processed survives because barcode is used as the key.</p>
    </div>
    <table style="margin-top:10px;">
      <tr><th>Database</th><th>Description</th><th>Deal Name</th><th>SpecialRetail</th><th>SpecialPercent</th></tr>
      <tr>
        <td><span class='badge badge-blue'>Legacy</span></td>
        <td>AVEENO Daily Moist Lot Recess. Pk</td>
        <td>DM02-26 Aveeno BOGOHP</td>
        <td>0</td><td>50%</td>
      </tr>
      <tr>
        <td><span class='badge badge-amber'>New App</span></td>
        <td>W&amp;L Candle Christmas 303g</td>
        <td>DM02-26 Wavertree &amp; London Prize Draw</td>
        <td>NULL</td><td>0%</td>
      </tr>
    </table>
    <div class="info-box" style="margin-top:10px;">
      The xlsx source has multiple items with no barcode (blank <em>Barcodes</em> cell) in different
      deals. Both apps store only the last one processed &mdash; but they process rows in different
      order, so a different item lands in the empty-barcode slot. The NULL SpecialRetail on this row
      is also covered by Bug 1 above.
    </div>
  </div>

  <div class="section">
    <h2>Minor &mdash; Trailing space in one Description</h2>
    <table>
      <tr><th>Barcode</th><th>Legacy Description</th><th>New App Description</th></tr>
      <tr>
        <td>9417034120993</td>
        <td>DUPLICATE TO MERGE - Antipodes HEAVENLY</td>
        <td>DUPLICATE TO MERGE - Antipodes HEAVENLY&nbsp;<span class='badge badge-red'>trailing space</span></td>
      </tr>
    </table>
    <div class="info-box" style="margin-top:8px;">
      The new app truncates descriptions over 40 characters by looking up the Details table, which
      has a trailing space for this product. Impact is minimal; fix by stripping the Details lookup result.
    </div>
  </div>

  <div class="section" style="background:#f8fafc;">
    <h2>Summary &mdash; Action Required</h2>
    <table>
      <tr><th>Issue</th><th>Rows affected</th><th>Severity</th><th>Fix location</th></tr>
      <tr>
        <td>SpecialRetail = NULL instead of -1</td><td>1,032</td>
        <td><span class='badge badge-red'>Bug</span></td>
        <td>specials/specials_helpers.py &mdash; _parse_special_price()</td>
      </tr>
      <tr>
        <td>GiftPharmacode = 'False' instead of ''</td><td>3,402</td>
        <td><span class='badge badge-red'>Bug</span></td>
        <td>specials/specials_helpers.py &mdash; line 181</td>
      </tr>
      <tr>
        <td>Empty-barcode row holds wrong item</td><td>1</td>
        <td><span class='badge badge-amber'>Data / ordering</span></td>
        <td>Source xlsx has blank barcodes &mdash; deduplicate or skip on import</td>
      </tr>
      <tr>
        <td>Trailing space in Description (one row)</td><td>1</td>
        <td><span class='badge badge-amber'>Minor</span></td>
        <td>Strip Details table lookup result</td>
      </tr>
      <tr>
        <td>DealName populated (was blank in legacy)</td><td>1,586</td>
        <td><span class='badge badge-green'>Improvement</span></td>
        <td>N/A</td>
      </tr>
      <tr>
        <td>SpecialNote populated (was blank in legacy)</td><td>2,658</td>
        <td><span class='badge badge-green'>Improvement</span></td>
        <td>N/A</td>
      </tr>
    </table>
  </div>

</div>
</body>
</html>"""

output_path = base + r"\comparison_report.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(report)
print(f"Report written to: {output_path}")
