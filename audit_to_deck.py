#!/usr/bin/env python3
"""
audit_to_deck.py - Turn any checklist-style audit spreadsheet into an
executive summary (console + JSON) and a PowerPoint deck.

Works on .xlsx/.xlsm/.csv files where one sheet holds a list of checks with a
status column (Pass / Fail / Compliant / Non-compliant / Partial / N/A ...).
The sheet, header row and column roles are detected automatically; every
detection can be overridden from the command line.

Install:
    pip install openpyxl python-pptx

Examples:
    python audit_to_deck.py report.xlsx
    python audit_to_deck.py report.xlsx --title "Q3 Security Review" --client "Acme Corp"
    python audit_to_deck.py report.csv --status-col "Result" --section-col "Domain"
    python audit_to_deck.py report.xlsx --status-map "Waived=N/A,Exception=PARTIAL"
    python audit_to_deck.py report.xlsx --no-deck --json summary.json

Outputs (next to the input unless --out-dir is given):
    <name>_Summary.pptx   the deck
    <name>_Summary.json   the numbers behind it (with --json or always when deck is built)
"""
import argparse
import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict

# --------------------------------------------------------------------------
# 1. Configuration: column synonyms and status vocabulary
# --------------------------------------------------------------------------
ROLE_SYNONYMS = {
    "status": ["status", "result", "outcome", "compliance", "compliance status", "state",
               "assessment", "rating", "check result", "audit result", "pass fail"],
    "item": ["checkpoint", "rpo checkpoint", "check", "check name", "control", "control name",
             "requirement", "test", "test name", "item", "rule", "title", "question",
             "criteria", "criterion", "best practice"],
    "section": ["section", "category", "area", "domain", "control family", "family", "group",
                "pillar", "component", "module", "topic"],
    "action": ["recommended action", "recommendation", "remediation", "action", "fix",
               "next step", "mitigation", "recommended fix"],
    "evidence": ["config value found", "evidence", "observation", "actual", "current value",
                 "current state", "finding", "details", "value found", "notes", "comment",
                 "comments", "observed"],
    "expected": ["expected standard", "expected", "standard", "baseline", "target",
                 "expected value", "best practice value"],
    "severity": ["severity", "priority", "risk", "risk level", "impact", "criticality"],
    "entity": ["property name", "property", "asset", "system", "application", "app", "host",
               "hostname", "server", "account", "site", "environment", "resource"],
}
ROLE_ORDER = ["status", "item", "section", "action", "evidence", "expected", "severity", "entity"]

BUCKETS = ["PASS", "PARTIAL", "FAIL", "NOT CONFIGURED", "NEEDS EVIDENCE", "N/A", "UNMAPPED"]
SCORED = ["PASS", "PARTIAL", "FAIL", "NOT CONFIGURED"]  # count toward the score
BUCKET_LABEL = {"PASS": "Pass", "PARTIAL": "Partial", "FAIL": "Fail",
                "NOT CONFIGURED": "Not configured", "NEEDS EVIDENCE": "Needs evidence",
                "N/A": "Not applicable", "UNMAPPED": "Unrecognized status"}

# Order matters: negatives are tested before positives ("not configured" before "configured").
STATUS_RULES = [
    ("N/A", r"\bn/?a\b|not applicable|out of scope|excluded|waived"),
    ("PARTIAL", r"partial|partly|in progress|some"),
    ("NOT CONFIGURED", r"not configured|not implemented|not enabled|not set|absent|not present|disabled"),
    ("NEEDS EVIDENCE", r"external|manual|unknown|not tested|not assessed|untested|review|pending|"
                       r"tbd|unverified|validat|needs|inconclusive|not evaluated"),
    ("FAIL", r"\bfail|non[- ]?compliant|not compliant|not met|\bno\b|deficien|\bgap\b|violation|"
             r"\bmissing\b|\bred\b|✗|✘|\bx\b|\bfalse\b|critical"),
    ("PASS", r"\bpass|compliant|\bok\b|\byes\b|\bmet\b|success|implemented|configured|enabled|"
             r"✓|✔|\btrue\b|\bgood\b|\bgreen\b|satisf"),
]

SEVERITY_RANK = [(r"critical|p0|sev ?1", 4), (r"high|p1|sev ?2", 3), (r"medium|med|moderate|p2|sev ?3", 2),
                 (r"low|p3|sev ?4|info", 1)]


def norm(s):
    return re.sub(r"[^a-z0-9/ ]+", " ", str(s or "").lower()).strip()


def normalize_status(raw, custom):
    if raw is None or str(raw).strip() == "":
        return None
    key = str(raw).strip()
    for k, v in custom.items():
        if key.lower() == k.lower():
            return v
    n = key.lower()
    for bucket, pattern in STATUS_RULES:
        if re.search(pattern, n):
            return bucket
    return "UNMAPPED"


SMALL_WORDS = {"and", "of", "the", "for", "to", "in", "on", "or", "a"}


def pretty(name):
    """'PROPERTY MANAGER SETTINGS' -> 'Property Manager Settings'; keeps short acronyms (DNS, TLS, API)."""
    if not name or not name.isupper():
        return name
    out = []
    for i, w in enumerate(name.split()):
        lw = w.lower()
        if i and lw in SMALL_WORDS:
            out.append(lw)
        elif len(w) <= 3 and w.isalpha():
            out.append(w)
        else:
            out.append(w.capitalize())
    return " ".join(out)


def severity_rank(v):
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return float(v)
    n = str(v).lower()
    for pat, rank in SEVERITY_RANK:
        if re.search(pat, n):
            return rank
    return 0


# --------------------------------------------------------------------------
# 2. Loading: find the checklist sheet, header row and columns
# --------------------------------------------------------------------------
def read_sheets(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".tsv", ".txt"):
        with open(path, newline="", encoding="utf-8-sig") as f:
            sample = f.read(4096)
            f.seek(0)
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|") if sample else csv.excel
            return {os.path.basename(path): [tuple(r) for r in csv.reader(f, dialect)]}
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    return {ws.title: [tuple(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


def match_role(header, role):
    h = norm(header)
    if not h:
        return 0
    best = 0
    for syn in ROLE_SYNONYMS[role]:
        if h == syn:
            best = max(best, 3)
        elif h.startswith(syn + " ") or h.endswith(" " + syn):
            best = max(best, 2)
        elif re.search(r"\b" + re.escape(syn) + r"\b", h):
            best = max(best, 1)
    return best


def map_columns(header, overrides):
    cols = {}
    used = set()
    for role in ROLE_ORDER:
        if overrides.get(role):
            want = norm(overrides[role])
            for i, h in enumerate(header):
                if norm(h) == want:
                    cols[role] = i
                    used.add(i)
                    break
            else:
                sys.exit(f"Column '{overrides[role]}' (for {role}) not found. Headers: {header}")
            continue
        scored = [(match_role(h, role), -i, i) for i, h in enumerate(header) if i not in used]
        scored = [s for s in scored if s[0] > 0]
        if scored:
            _, _, idx = max(scored)
            cols[role] = idx
            used.add(idx)
    return cols


def find_checklist(sheets, overrides, custom_status, forced_sheet=None, scan_rows=25):
    best = None
    for name, rows in sheets.items():
        if forced_sheet and name != forced_sheet:
            continue
        for hr in range(min(scan_rows, len(rows))):
            header = [str(c).strip() if c is not None else "" for c in rows[hr]]
            if sum(1 for h in header if h) < 2:
                continue
            cols = map_columns(header, overrides)
            if "status" not in cols:
                continue
            si = cols["status"]
            data = [r for r in rows[hr + 1:] if si < len(r) and normalize_status(r[si], custom_status)]
            recognized = sum(1 for r in data if normalize_status(r[si], custom_status) != "UNMAPPED")
            score = (recognized, len(cols))
            if recognized and (best is None or score > best[0]):
                best = (score, name, hr, header, cols, data)
    if not best:
        sys.exit("Could not find a sheet with a recognizable status column. "
                 "Use --sheet and --status-col to point at it.")
    _, name, hr, header, cols, data = best
    return name, hr, header, cols, data


def find_key_values(sheets, skip_sheet):
    """Collect 2-cell 'label | value' rows from other sheets (summary tabs)."""
    kv = {}
    for name, rows in sheets.items():
        if name == skip_sheet:
            continue
        for r in rows:
            cells = [c for c in r if c is not None and str(c).strip() != ""]
            if len(cells) == 2 and isinstance(cells[0], str) and len(cells[0]) < 40:
                kv[cells[0].strip()] = cells[1]
    return kv


# --------------------------------------------------------------------------
# 3. Analysis
# --------------------------------------------------------------------------
def analyze(path, args):
    custom = {}
    for pair in filter(None, (args.status_map or "").split(",")):
        k, _, v = pair.partition("=")
        v = v.strip().upper()
        if v not in BUCKETS:
            sys.exit(f"--status-map target '{v}' must be one of {BUCKETS}")
        custom[k.strip()] = v
    overrides = {r: getattr(args, f"{r}_col") for r in ROLE_ORDER}

    sheets = read_sheets(path)
    sheet, hr, header, cols, data = find_checklist(sheets, overrides, custom, args.sheet)

    def get(r, role):
        i = cols.get(role)
        if i is None or i >= len(r):
            return None
        v = r[i]
        return str(v).strip() if v is not None else None

    rows = []
    for r in data:
        rows.append({
            "status_raw": get(r, "status"),
            "status": normalize_status(r[cols["status"]], custom),
            "item": get(r, "item") or "(unnamed check)",
            "section": pretty((get(r, "section") or "General").strip()),
            "action": get(r, "action"),
            "evidence": get(r, "evidence"),
            "expected": get(r, "expected"),
            "severity": get(r, "severity"),
            "entity": get(r, "entity"),
        })

    totals = Counter(r["status"] for r in rows)
    by_section = defaultdict(Counter)
    for r in rows:
        by_section[r["section"]][r["status"]] += 1
    entities = sorted({r["entity"] for r in rows if r["entity"]})
    if len(entities) > 1:  # several assets: tag each check so lists stay unambiguous
        for r in rows:
            if r["entity"]:
                r["item"] = f"{r['item']} ({r['entity']})"
    by_entity = defaultdict(Counter)
    if 1 < len(entities) <= 50:
        for r in rows:
            by_entity[r["entity"]][r["status"]] += 1

    verifiable = sum(totals[b] for b in SCORED)
    score = (totals["PASS"] + 0.5 * totals["PARTIAL"]) / verifiable if verifiable else 0.0

    # Workbook's own summary figures (if any), for reconciliation
    kv = find_key_values(sheets, sheet)
    reported_counts, reported_score, reported_risk = {}, None, None
    for k, v in kv.items():
        nk = k.lower()
        if "score" in nk:
            # prefer the headline score (final / overall / compliance) over sub-scores
            if reported_score is None or re.search(r"final|overall|compliance|total", nk):
                reported_score = v
        elif "risk" in nk and isinstance(v, str):
            reported_risk = v
        elif isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
            b = normalize_status(k, custom)
            if b and b != "UNMAPPED" and len(k.split()) <= 3 and not re.match(r"(total|number|#|count)", nk):
                reported_counts[b] = reported_counts.get(b, 0) + int(v)

    if reported_risk:
        risk, risk_source = str(reported_risk).strip().title(), "workbook"
    else:
        high_fails = sum(1 for r in rows if r["status"] == "FAIL" and severity_rank(r["severity"]) >= 3)
        risk = "High" if score < args.high_below or high_fails >= 3 else \
               "Medium" if score < args.medium_below else "Low"
        risk_source = "derived from score"

    # Evidence grouping: named exports/reports if the text mentions them, else section
    pat = re.compile(args.evidence_pattern) if args.evidence_pattern else None
    ev_groups = defaultdict(list)
    for r in rows:
        if r["status"] != "NEEDS EVIDENCE":
            continue
        key = None
        if pat:
            hits = pat.findall(" ".join(filter(None, [r["evidence"], r["action"]])))
            key = hits[-1] if hits else None
        ev_groups[key or r["section"]].append(r["item"])

    def sort_key(r):
        return (-severity_rank(r["severity"]), r["section"], r["item"])

    return {
        "file": os.path.basename(path), "sheet": sheet, "header_row": hr + 1,
        "columns": {role: header[i] for role, i in cols.items()},
        "rows": rows, "totals": {b: totals[b] for b in BUCKETS if totals[b]},
        "by_section": {s: dict(c) for s, c in by_section.items()},
        "by_entity": {e: dict(c) for e, c in by_entity.items()},
        "entities": entities, "verifiable": verifiable, "score": score,
        "risk": risk, "risk_source": risk_source,
        "reported": {"counts": reported_counts, "score": reported_score, "risk": reported_risk},
        "unmapped_statuses": sorted({r["status_raw"] for r in rows if r["status"] == "UNMAPPED"}),
        "fails": sorted([r for r in rows if r["status"] == "FAIL"], key=sort_key),
        "partials": sorted([r for r in rows if r["status"] in ("PARTIAL", "NOT CONFIGURED")], key=sort_key),
        "passes": [r for r in rows if r["status"] == "PASS"],
        "evidence_groups": dict(sorted(ev_groups.items(), key=lambda kv: -len(kv[1]))),
    }


def print_report(a):
    t = a["totals"]
    total = sum(t.values())
    print(f"File: {a['file']}  |  sheet: '{a['sheet']}'  |  header row: {a['header_row']}")
    print("Columns detected: " + ", ".join(f"{k}='{v}'" for k, v in a["columns"].items()))
    if a["unmapped_statuses"]:
        print(f"WARNING: unrecognized status values {a['unmapped_statuses']} "
              f"- map them with --status-map \"Value=PASS,...\"")

    print(f"\n=== STATUS TOTALS (n={total}) ===")
    for b in BUCKETS:
        if t.get(b):
            print(f"  {BUCKET_LABEL[b]:<22}{t[b]:>5}")

    shown = [b for b in BUCKETS if t.get(b)]
    print("\n=== BY SECTION ===")
    print(f"  {'Section':<30}" + "".join(f"{BUCKET_LABEL[b][:11]:>13}" for b in shown))
    for s, c in sorted(a["by_section"].items(), key=lambda kv: -kv[1].get("FAIL", 0)):
        print(f"  {s[:29]:<30}" + "".join(f"{c.get(b, 0):>13}" for b in shown))

    if a["by_entity"]:
        print("\n=== BY ENTITY ===")
        for e, c in sorted(a["by_entity"].items(), key=lambda kv: -kv[1].get("FAIL", 0)):
            print(f"  {e[:40]:<42} fails={c.get('FAIL', 0):<4} pass={c.get('PASS', 0)}")

    print("\n=== SCORE ===")
    print(f"  Verified checks passing: {t.get('PASS', 0)} of {a['verifiable']}")
    print(f"  Score (pass + half-credit partial, over verifiable checks): {a['score']:.1%}")
    print(f"  Overall risk: {a['risk']} ({a['risk_source']})")

    rep = a["reported"]
    if rep["counts"] or rep["score"] is not None:
        print("\n=== RECONCILIATION: summary figures in workbook vs line items ===")
        if rep["score"] is not None:
            print(f"  Reported score: {rep['score']}   |   recomputed: {a['score']:.2%}")
        clean = True
        for b, v in rep["counts"].items():
            ok = v == t.get(b, 0)
            clean &= ok
            print(f"  {BUCKET_LABEL[b]:<22} reported={v:<5} line items={t.get(b, 0):<5}{'' if ok else '  <-- MISMATCH'}")
        if clean and rep["counts"]:
            print("  Counts match.")

    print(f"\n=== FAILED CHECKS ({len(a['fails'])}) ===")
    for r in a["fails"]:
        sev = f" [{r['severity']}]" if r["severity"] else ""
        print(f"  [{r['section']}] {r['item']}{sev}")
        if r["evidence"]:
            print(f"      found: {r['evidence'][:110]}")
        if r["action"]:
            print(f"      fix:   {r['action'][:110]}")

    if a["evidence_groups"]:
        n = sum(map(len, a["evidence_groups"].values()))
        print(f"\n=== NEEDS EVIDENCE ({n}) ===")
        for g, items in a["evidence_groups"].items():
            print(f"  {g} ({len(items)})")
            for i in items:
                print(f"      - {i}")


# --------------------------------------------------------------------------
# 4. Deck
# --------------------------------------------------------------------------
def build_deck(a, out_path, title, subtitle, date_label):
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt

    INK, SLATE, MUTE, TINT, DARK, WHITE = "17212B", "5B6B7A", "8A97A3", "F3F6F8", "1F2D3A", "FFFFFF"
    ACCENT, SOFT = "F2A33A", "A9B6C2"
    COLOR = {"PASS": "2E9E6B", "PARTIAL": "E0A526", "FAIL": "D64545", "NOT CONFIGURED": "7C8B99",
             "NEEDS EVIDENCE": "3F7CC4", "N/A": "B8C2CC", "UNMAPPED": "9B59B6"}
    HF, BF = "Arial", "Calibri"
    rgb = RGBColor.from_string

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    blank = prs.slide_layouts[6]
    t = a["totals"]
    total = sum(t.values())
    page = [0]

    def clip(s, n):
        s = re.sub(r"\s+", " ", str(s or "")).strip()
        return s if len(s) <= n else s[: n - 1].rstrip() + "…"

    def new_slide(bg=WHITE):
        s = prs.slides.add_slide(blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = rgb(bg)
        page[0] += 1
        return s

    def text(s, x, y, w, h, paras, size=12, color=INK, bold=False, font=BF, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, italic=False, bullets=False, space_after=0):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = anchor
        if isinstance(paras, str):
            paras = [paras]
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if space_after:
                p.space_after = Pt(space_after)
            if bullets:
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", "171450")
                pPr.set("indent", "-171450")
                bu = pPr.makeelement(qn("a:buChar"), {"char": "•"})
                pPr.append(bu)
            runs = para if isinstance(para, list) else [(para, {})]
            for rt, ro in runs:
                r = p.add_run()
                r.text = rt
                f = r.font
                f.name = ro.get("font", font)
                f.size = Pt(ro.get("size", size))
                f.bold = ro.get("bold", bold)
                f.italic = ro.get("italic", italic)
                f.color.rgb = rgb(ro.get("color", color))
        return tb

    def box(s, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06):
        sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = radius
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
        sh.line.fill.background()
        sh.shadow.inherit = False
        return sh

    def header(s, title_, sub=None):
        text(s, 0.5, 0.3, 9, 0.6, clip(title_, 60), size=26, bold=True, font=HF)
        if sub:
            text(s, 0.5, 0.88, 9, 0.35, clip(sub, 110), size=14, color=SLATE)

    def footer(s):
        text(s, 0.5, 5.25, 9, 0.25, f"{clip(subtitle, 60)}  ·  {page[0]}", size=9, color=MUTE, align=PP_ALIGN.RIGHT)

    def plural(n, word):
        return f"{n} {word}{'' if n == 1 else 's'}"

    fails, sections = a["fails"], a["by_section"]
    npass, nfail = t.get("PASS", 0), t.get("FAIL", 0)
    nev = t.get("NEEDS EVIDENCE", 0)

    # ---- 1. Title
    s = new_slide(DARK)
    text(s, 0.6, 1.2, 8.8, 0.35, "AUDIT SUMMARY", size=13, bold=True, color=ACCENT)
    text(s, 0.6, 1.6, 8.8, 0.9, clip(title, 45), size=36, bold=True, color=WHITE, font=HF)
    text(s, 0.6, 2.55, 8.8, 0.4, clip(subtitle, 80), size=18, color="C9D3DC")
    stats = [(total, "checks audited"), (len(sections), "areas reviewed")]
    if len(a["entities"]) > 1:
        stats.append((len(a["entities"]), "assets in scope"))
    for i, (n, l) in enumerate(stats):
        text(s, 0.6 + i * 2.5, 3.55, 2.2, 0.6, str(n), size=30, bold=True, color=WHITE, font=HF)
        text(s, 0.6 + i * 2.5, 4.12, 2.2, 0.3, l, size=12, color=SOFT)
    text(s, 0.6, 4.95, 4, 0.3, date_label, size=11, color=SOFT)

    # ---- 2. At a glance
    s = new_slide()
    worst = sorted(sections.items(), key=lambda kv: -kv[1].get("FAIL", 0))
    worst_names = [k for k, v in worst[:2] if v.get("FAIL")]
    sub = f"Overall risk is {a['risk']}" + (f", driven mainly by {' and '.join(worst_names)}" if worst_names else "")
    header(s, "At a glance", sub)
    risk_c = {"High": "D64545", "Critical": "D64545", "Medium": "E0A526", "Low": "2E9E6B"}.get(a["risk"], SLATE)
    risk_bg = {"D64545": "FBEAEA", "E0A526": "FCF4E1", "2E9E6B": "E8F5EE"}.get(risk_c, TINT)
    box(s, 0.5, 1.5, 2.6, 1.55, risk_bg)
    text(s, 0.7, 1.65, 2.2, 0.3, "OVERALL RISK", size=11, bold=True, color=risk_c)
    text(s, 0.7, 1.95, 2.2, 0.7, a["risk"], size=40, bold=True, color=risk_c, font=HF)
    text(s, 0.7, 2.62, 2.3, 0.3, plural(nfail, "failed check"), size=12)
    box(s, 0.5, 3.25, 2.6, 1.55, "E8F5EE")
    text(s, 0.7, 3.4, 2.3, 0.3, "VERIFIED CHECKS PASSING", size=11, bold=True, color=COLOR["PASS"])
    text(s, 0.7, 3.7, 2.2, 0.7, [[(str(npass), {"size": 40, "bold": True}), (f" of {a['verifiable']}", {"size": 20})]],
         color=COLOR["PASS"], font=HF)
    text(s, 0.7, 4.37, 2.3, 0.3, f"{nev} more need evidence" if nev else f"Score {a['score']:.0%}", size=12)
    shown = [b for b in BUCKETS if t.get(b)]
    cd = CategoryChartData()
    cd.categories = [BUCKET_LABEL[b] for b in shown]
    cd.add_series("Status", [t[b] for b in shown])
    ch = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(3.3), Inches(1.4), Inches(3.1), Inches(3.5), cd).chart
    ch.has_legend = False
    ch.has_title = False
    plot = ch.plots[0]
    hole = plot._element.find(qn("c:holeSize"))
    if hole is None:
        hole = plot._element.makeelement(qn("c:holeSize"), {})
        plot._element.append(hole)
    hole.set("val", "58")
    for i, b in enumerate(shown):
        pt = plot.series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(COLOR[b])
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_value, dl.show_category_name, dl.show_percentage = True, False, False
    dl.number_format, dl.number_format_is_linked = "0;;;", False
    dl.font.size, dl.font.bold, dl.font.color.rgb = Pt(11), True, rgb(WHITE)
    text(s, 4.25, 2.75, 1.2, 0.8, [[(str(total), {"size": 26, "bold": True})], [("checks", {"size": 11})]],
         font=HF, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    step = min(0.58, 3.0 / max(len(shown), 1))
    for i, b in enumerate(shown):
        y = 1.75 + i * step
        box(s, 6.7, y + 0.07, 0.2, 0.2, COLOR[b], shape=MSO_SHAPE.OVAL)
        text(s, 7.05, y, 2.0, 0.34, BUCKET_LABEL[b], size=13, anchor=MSO_ANCHOR.MIDDLE)
        text(s, 8.9, y, 0.6, 0.34, str(t[b]), size=14, bold=True, font=HF, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    s.notes_slide.notes_text_frame.text = (
        f"Score = (pass + half of partial) / verifiable checks = {a['score']:.1%}. "
        f"Risk level: {a['risk']} ({a['risk_source']}).")
    footer(s)

    # ---- 3. By section
    s = new_slide()
    top_fail = sum(v.get("FAIL", 0) for _, v in worst[:2])
    sub = (f"{top_fail} of {nfail} failures sit in {' and '.join(worst_names)}" if len(worst_names) == 2 and nfail
           else "Status of every check, grouped by area")
    header(s, "Where the gaps sit", sub)
    order = sorted(sections.items(), key=lambda kv: (-kv[1].get("FAIL", 0), -sum(kv[1].values())))
    if len(order) > 12:
        other = Counter()
        for _, c in order[11:]:
            other.update(c)
        order = order[:11] + [("Other areas", dict(other))]
    order = list(reversed(order))
    cd = CategoryChartData()
    cd.categories = [clip(k, 28) for k, _ in order]
    for b in shown:
        cd.add_series(BUCKET_LABEL[b], [c.get(b, 0) for _, c in order])
    ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(0.4), Inches(1.35), Inches(6.2), Inches(3.85), cd).chart
    plot = ch.plots[0]
    plot.gap_width, plot.overlap = 45, 100
    for i, b in enumerate(shown):
        ser = plot.series[i]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(COLOR[b])
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_value, dl.number_format, dl.number_format_is_linked = True, "0;;;", False
    dl.position = XL_LABEL_POSITION.CENTER
    dl.font.size, dl.font.bold, dl.font.color.rgb = Pt(9), True, rgb(WHITE)
    ch.has_title = False
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.tick_labels.font.size = Pt(10 if len(order) > 9 else 11)
    ch.category_axis.tick_labels.font.color.rgb = rgb(INK)
    ch.category_axis.format.line.fill.background()
    ch.has_legend = True
    ch.legend.position, ch.legend.include_in_layout = XL_LEGEND_POSITION.BOTTOM, False
    ch.legend.font.size, ch.legend.font.color.rgb = Pt(10), rgb(SLATE)

    def rate(c, b):
        n = sum(c.values())
        return c.get(b, 0) / n if n else 0
    callouts = []
    if nfail:
        k, c = worst[0]
        callouts.append(("Needs attention", f"{k}: {plural(c['FAIL'], 'failed check')} of {sum(c.values())}.", COLOR["FAIL"]))
    strong = sorted([kv for kv in sections.items() if kv[1].get("PASS")], key=lambda kv: (-rate(kv[1], "PASS"), -kv[1]["PASS"]))
    if strong:
        k, c = strong[0]
        callouts.append(("Strongest area", f"{k}: {c['PASS']} of {sum(c.values())} checks pass.", COLOR["PASS"]))
    unv = sorted([kv for kv in sections.items() if kv[1].get("NEEDS EVIDENCE")], key=lambda kv: -kv[1]["NEEDS EVIDENCE"])
    if unv:
        k, c = unv[0]
        callouts.append(("Least verified", f"{k}: {plural(c['NEEDS EVIDENCE'], 'check')} need{'s' if c['NEEDS EVIDENCE'] == 1 else ''} outside evidence.", COLOR["NEEDS EVIDENCE"]))
    for i, (h, b, c) in enumerate(callouts):
        y = 1.45 + i * 1.22
        box(s, 6.85, y, 2.65, 1.05, TINT)
        text(s, 7.0, y + 0.1, 2.4, 0.28, h, size=12, bold=True, color=c, font=HF)
        text(s, 7.0, y + 0.38, 2.4, 0.62, clip(b, 90), size=11)
    footer(s)

    # ---- 4. What's working
    if a["passes"]:
        s = new_slide()
        header(s, "What's already working well", f"{npass} checks pass across {plural(sum(1 for c in sections.values() if c.get('PASS')), 'area')}")
        pass_by_sec = defaultdict(list)
        for r in a["passes"]:
            pass_by_sec[r["section"]].append(r["item"])
        cards = sorted(pass_by_sec.items(), key=lambda kv: -len(kv[1]))[:4]
        for i, (sec, items) in enumerate(cards):
            x, y = 0.5 + (i % 2) * 4.6, 1.5 + (i // 2) * 1.8
            box(s, x, y, 4.4, 1.6, TINT)
            box(s, x + 0.25, y + 0.3, 0.6, 0.6, COLOR["PASS"], shape=MSO_SHAPE.OVAL)
            text(s, x + 0.25, y + 0.3, 0.6, 0.6, str(len(items)), size=16, bold=True, color=WHITE, font=HF,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            text(s, x + 1.05, y + 0.22, 3.2, 0.35, clip(sec, 32), size=14, bold=True, font=HF)
            text(s, x + 1.05, y + 0.58, 3.2, 0.9, clip("; ".join(items), 150), size=11.5, color=SLATE)
        footer(s)

    # ---- 5. Failures by area
    if fails:
        s = new_slide()
        fail_by_sec = defaultdict(list)
        for r in fails:
            fail_by_sec[r["section"]].append(r)
        groups = sorted(fail_by_sec.items(), key=lambda kv: -len(kv[1]))
        header(s, f"{plural(nfail, 'failure')}, {plural(len(groups), 'area')}",
               "Grouped by area; highest-severity items listed first")
        shown_g = groups[:4]
        w = (9 - 0.2 * (len(shown_g) - 1)) / len(shown_g)
        for i, (sec, rs) in enumerate(shown_g):
            x = 0.5 + i * (w + 0.2)
            box(s, x, 1.45, w, 3.55, TINT)
            text(s, x + 0.18, 1.55, w - 0.36, 0.62, str(len(rs)), size=30, bold=True, color=COLOR["FAIL"], font=HF)
            text(s, x + 0.18, 2.2, w - 0.36, 0.5, clip(sec, 40), size=13, bold=True, font=HF)
            items = [clip(r["item"], 60) for r in rs[:4]] + ([f"+ {len(rs) - 4} more (see appendix)"] if len(rs) > 4 else [])
            text(s, x + 0.18, 2.9, w - 0.3, 2.0, items, size=10.5, color=SLATE, bullets=True, space_after=4)
        if len(groups) > 4:
            rest = sum(len(v) for _, v in groups[4:])
            text(s, 0.5, 5.02, 6, 0.22, f"+ {plural(rest, 'more failure')} in {plural(len(groups) - 4, 'other area')} (see appendix)",
                 size=10, italic=True, color=SLATE)
        footer(s)

    # ---- 6. Action plan
    s = new_slide()
    header(s, "Recommended action plan", "Fix failures first, then close partial gaps, then collect missing evidence")

    def act(r):
        return clip(r["action"] or r["item"], 70)
    phases = [
        ("NOW", "0–30 days", "Fix failed checks", COLOR["FAIL"], fails, act),
        ("NEXT", "30–90 days", "Close partial gaps", COLOR["PARTIAL"], a["partials"], act),
    ]
    later_items = [f"{clip(g, 40)} ({len(v)})" for g, v in a["evidence_groups"].items()]
    for i, (tag, when, h, c, src, fn) in enumerate(phases + [("LATER", "90+ days", "Confirm unverified checks", COLOR["NEEDS EVIDENCE"], later_items, None)]):
        x, w = 0.5 + i * 3.07, 2.87
        box(s, x, 1.45, w, 3.65, TINT)
        box(s, x + 0.2, 1.65, 0.85, 0.32, c, radius=0.5)
        text(s, x + 0.2, 1.65, 0.85, 0.32, tag, size=11, bold=True, color=WHITE, font=HF, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 1.15, 1.65, 1.5, 0.32, when, size=11, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 0.2, 2.1, w - 0.4, 0.35, h, size=14, bold=True, font=HF)
        entries = [fn(r) for r in src] if fn else list(src)
        # de-duplicate identical recommendations
        seen, uniq = set(), []
        for e in entries:
            if e.lower() not in seen:
                seen.add(e.lower())
                uniq.append(e)
        if not uniq:
            uniq = ["Nothing outstanding"]
        body = uniq[:5] + ([f"+ {len(uniq) - 5} more"] if len(uniq) > 5 else [])
        text(s, x + 0.2, 2.5, w - 0.35, 2.5, body, size=10.5, bullets=True, space_after=4)
    footer(s)

    # ---- 7. Evidence needed
    if a["evidence_groups"]:
        s = new_slide()
        groups = list(a["evidence_groups"].items())
        header(s, "To finish the picture", f"{plural(nev, 'check')} can’t be judged from the data provided")
        showg = groups[:6]
        rowh = min(0.73, 3.6 / len(showg))
        for i, (g, items) in enumerate(showg):
            y = 1.42 + i * rowh
            box(s, 0.5, y + 0.05, 0.5, 0.5, COLOR["NEEDS EVIDENCE"], shape=MSO_SHAPE.OVAL)
            text(s, 0.5, y + 0.05, 0.5, 0.5, str(len(items)), size=13, bold=True, color=WHITE, font=HF,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            text(s, 1.2, y, 5.3, 0.3, clip(g, 50), size=13, bold=True, font=HF)
            text(s, 1.2, y + 0.3, 5.3, 0.28, clip(", ".join(items), 78), size=10.5, color=SLATE)
        if len(groups) > 6:
            text(s, 1.2, 1.42 + 6 * rowh, 5, 0.25, f"+ {plural(len(groups) - 6, 'more group')}", size=10, italic=True, color=SLATE)
        box(s, 7.0, 1.45, 2.5, 3.45, DARK)
        text(s, 7.2, 1.65, 2.1, 0.35, "Why it matters", size=14, bold=True, color=ACCENT, font=HF)
        text(s, 7.2, 2.05, 2.1, 2.7,
             f"These {nev} checks are {nev / total:.0%} of the audit. Once the evidence is in, the score and risk level "
             f"can be confirmed.", size=11.5, color="E4EAEF")
        footer(s)

    # ---- 8. Next steps
    s = new_slide(DARK)
    text(s, 0.6, 0.5, 8.8, 0.7, "Next steps", size=32, bold=True, color=WHITE, font=HF)
    steps = []
    if fails:
        steps.append((f"Assign owners for the {plural(nfail, 'failed check')}", "Target the Now list within 30 days"))
    if a["partials"]:
        steps.append((f"Plan fixes for {plural(len(a['partials']), 'partial or missing setting')}", "Schedule within 90 days"))
    if nev:
        steps.append((f"Provide evidence for {plural(nev, 'unverified check')}", "So the audit can be completed"))
    steps.append(("Re-audit and re-score", "Confirm the risk level after changes land"))
    for i, (h, b) in enumerate(steps[:4]):
        y = 1.45 + i * 0.95
        box(s, 0.6, y, 0.6, 0.6, ACCENT, shape=MSO_SHAPE.OVAL)
        text(s, 0.6, y, 0.6, 0.6, str(i + 1), size=20, bold=True, color=DARK, font=HF, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, 1.5, y - 0.02, 7.9, 0.36, h, size=17, bold=True, color=WHITE, font=HF)
        text(s, 1.5, y + 0.33, 7.9, 0.3, b, size=12.5, color=SOFT)
    text(s, 0.6, 5.0, 8.8, 0.3, f"Line-by-line detail for all {total} checks remains in the source workbook.",
         size=11, italic=True, color=SOFT)

    # ---- 9+. Appendix: failed checks table, paginated
    per = 10
    cols_ = [("Area", "section", 1.6), ("Check", "item", 2.3)]
    has_ev, has_act = any(r["evidence"] for r in fails), any(r["action"] for r in fails)
    if has_ev and has_act:
        cols_ += [("Found", "evidence", 2.55), ("Recommended fix", "action", 2.55)]
    elif has_ev or has_act:
        cols_ += [("Found" if has_ev else "Recommended fix", "evidence" if has_ev else "action", 5.1)]
    else:
        cols_[1] = ("Check", "item", 7.4)
    widths = [c[2] for c in cols_]
    scale = 9 / sum(widths)
    npages = max(1, -(-len(fails) // per))
    size = -(-len(fails) // npages) if fails else 0
    chunks = [fails[i:i + size] for i in range(0, len(fails), size)] if fails else []
    for ci, chunk in enumerate(chunks):
        s = new_slide()
        suffix = f" ({ci + 1}/{len(chunks)})" if len(chunks) > 1 else ""
        text(s, 0.5, 0.3, 9, 0.6, f"Appendix: failed checks{suffix}", size=24, bold=True, font=HF)
        tbl = s.shapes.add_table(len(chunk) + 1, len(cols_), Inches(0.5), Inches(1.0), Inches(9),
                                 Inches(0.33 * (len(chunk) + 1))).table
        for j, (lab, _, wdt) in enumerate(cols_):
            tbl.columns[j].width = Inches(wdt * scale)
        limits = {"section": 28, "item": 55, "evidence": 70, "action": 70}
        for i in range(len(chunk) + 1):
            tbl.rows[i].height = Inches(0.33)
            for j, (lab, key, _) in enumerate(cols_):
                cell = tbl.cell(i, j)
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                val = lab if i == 0 else clip(chunk[i - 1][key] or "—", 110 if len(cols_) == 2 else limits[key])
                cell.text = val
                p = cell.text_frame.paragraphs[0]
                f = p.runs[0].font
                f.name, f.size = BF, Pt(9 if len(cols_) == 4 else 10)
                f.bold = i == 0
                f.color.rgb = rgb(WHITE if i == 0 else INK)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(DARK if i == 0 else (TINT if i % 2 else WHITE))
        footer(s)

    prs.save(out_path)


# --------------------------------------------------------------------------
# 5. CLI
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", help=".xlsx, .xlsm or .csv file")
    ap.add_argument("--sheet", help="sheet name holding the checklist (auto-detected by default)")
    for role in ROLE_ORDER:
        ap.add_argument(f"--{role}-col", help=f"exact header of the {role} column (auto-detected by default)")
    ap.add_argument("--status-map", help='extra status mappings, e.g. "Waived=N/A,Exception=PARTIAL". '
                                         f"Targets: {', '.join(BUCKETS[:-1])}")
    ap.add_argument("--evidence-pattern", default=r"\b([A-Z][A-Z0-9]*(?:_[A-Z0-9]+)*_(?:EXPORT|REPORT|EVIDENCE))\b",
                    help="regex used to group needs-evidence checks by the artifact they mention "
                         "(default finds tokens like DNS_EXPORT); '' groups by section instead")
    ap.add_argument("--high-below", type=float, default=0.60, help="score below this = High risk when the workbook gives none")
    ap.add_argument("--medium-below", type=float, default=0.80, help="score below this = Medium risk")
    ap.add_argument("--title", default="Findings & Recommended Plan")
    ap.add_argument("--client", help="subtitle line (defaults to the audited asset or file name)")
    ap.add_argument("--date", help="date label on the title slide (defaults to current month)")
    ap.add_argument("--out-dir", help="where to write outputs (defaults to the input's folder)")
    ap.add_argument("--json", help="also write the analysis to this JSON path")
    ap.add_argument("--no-deck", action="store_true", help="analysis only, skip the PowerPoint")
    args = ap.parse_args()

    a = analyze(args.input, args)
    print_report(a)

    stem = os.path.splitext(os.path.basename(args.input))[0]
    out_dir = args.out_dir or os.path.dirname(os.path.abspath(args.input))
    os.makedirs(out_dir, exist_ok=True)

    if args.json:
        slim = {k: v for k, v in a.items() if k not in ("rows", "passes")}
        with open(args.json, "w") as f:
            json.dump(slim, f, indent=2, default=str)
        print(f"\nWrote {args.json}")

    if not args.no_deck:
        import datetime
        ents = a["entities"]
        subtitle = args.client or (", ".join(ents) if 0 < len(ents) <= 2 else stem.replace("_", " "))
        date_label = args.date or datetime.date.today().strftime("%B %Y")
        out = os.path.join(out_dir, f"{stem}_Summary.pptx")
        build_deck(a, out, args.title, subtitle, date_label)
        print(f"Wrote {out}")


if __name__ == "__main__":
    main()