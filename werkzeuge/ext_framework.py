"""Framework: zusätzliche Aufgabenblätter (Aufgabe 2 bis 4) in die Übungsdateien einfügen und Modultexte erzeugen."""
import re, sys, math
from datetime import date, time, datetime, timedelta
sys.argv = sys.argv[:1] + ["/tmp/_a", "/tmp/_l"]
from gen_stufe1 import (EURO, PROZ, ZAHL, F10, F10B, F12B, F10I, GRAU, RAHMEN, WRAP, ZENTR, zeilenhoehe)
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.styles import Alignment

DATUM = "DD.MM.YYYY"
UHR = "hh:mm"
HR = 13           # Kopfzeile der Tabelle
R0 = 14           # erste Datenzeile
BASIS_HINT = "Verwenden Sie in Ihren Formeln Zellbezüge und keine eingetippten Ergebnisse. Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."

FN = [("ROUNDUP", "AUFRUNDEN"), ("ROUNDDOWN", "ABRUNDEN"), ("ROUND", "RUNDEN"), ("COUNTA", "ANZAHL2"), ("COUNTIF", "ZÄHLENWENN"),
      ("COUNT", "ANZAHL"), ("SUMIF", "SUMMEWENN"), ("SUM", "SUMME"), ("AVERAGE", "MITTELWERT"), ("VLOOKUP", "SVERWEIS"),
      ("RANK", "RANG"), ("TODAY", "HEUTE"), ("YEAR", "JAHR"), ("MONTH", "MONAT"), ("DAYS360", "TAGE360"), ("DAY", "TAG"),
      ("HOUR", "STUNDE"), ("AND", "UND"), ("OR", "ODER"), ("IF", "WENN")]


def de(f):
    out, i, q = [], 0, False
    seg = ""
    parts = re.split(r'("[^"]*")', f)
    res = []
    for p in parts:
        if p.startswith('"'):
            res.append(p); continue
        for en, d in FN:
            p = re.sub(r'\b' + en + r'\(', d + '(', p)
        p = re.sub(r'\bTRUE\b', 'WAHR', p); p = re.sub(r'\bFALSE\b', 'FALSCH', p)
        p = p.replace(",", ";")
        p = re.sub(r'(?<=\d)\.(?=\d)', ',', p)
        res.append(p)
    return "".join(res)


def num_de(v, nd):
    s = f"{v:,.{nd}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmtval(v, fmt):
    if v is None: return "leer"
    if isinstance(v, datetime): return v.strftime("%d.%m.%Y")
    if isinstance(v, date): return v.strftime("%d.%m.%Y")
    if isinstance(v, time): return v.strftime("%H:%M")
    if isinstance(v, timedelta):
        m = round(v.total_seconds() / 60); return f"{m // 60}:{m % 60:02d}"
    if isinstance(v, str): return f"„{v}“" if v != "" else "leer (Text leer)"
    if isinstance(v, bool): return "WAHR" if v else "FALSCH"
    if fmt == EURO: return num_de(v, 2) + " €"
    if fmt and '.0000' in fmt: return num_de(v, 4) + " €"
    if fmt == PROZ: return num_de(v * 100, 2) + " %"
    if fmt == "0.00": return num_de(v, 2)
    if fmt in (ZAHL, "0"): return num_de(v, 0) if fmt == ZAHL else str(int(round(v)))
    if abs(v - round(v)) < 1e-9: return num_de(round(v), 0)
    return num_de(v, 4).rstrip("0").rstrip(",")


def build_extra(wb, es, index):
    if es["name"] in wb.sheetnames: del wb[es["name"]]
    ws = wb.create_sheet(es["name"], index)
    n = len(es["headers"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    c = ws.cell(row=1, column=1, value=es["titel"]); c.font = F12B; c.alignment = ZENTR
    for i, h in enumerate(es["headers"]):
        x = ws.cell(row=3, column=i + 1, value=h); x.font = F10B; x.fill = GRAU; x.border = RAHMEN; x.alignment = ZENTR
    for k, row in enumerate(es["rows"]):
        for i, v in enumerate(row):
            x = ws.cell(row=4 + k, column=i + 1, value=v); x.font = F10; x.border = RAHMEN
            f = es["fmts"][i] if es.get("fmts") else None
            if f: x.number_format = f
    for i, w in enumerate(es["widths"]): ws.column_dimensions[L(i + 1)].width = w
    return ws


def build_sheet(wb, spec, loesung, index):
    ws = wb.create_sheet(spec["name"], index)
    for j, es in enumerate(spec.get("extra_sheets", [])): build_extra(wb, es, index + 1 + j)
    cols = spec["cols"]
    n = len(cols)
    widths = {L(i + 1): c[1] for i, c in enumerate(cols)}
    widths.update(spec.get("widths", {}))
    # Merge-Breite bestimmen
    cum, last = 0, 0
    while cum < 105 or last < n:
        last += 1
        cum += widths.get(L(last), 10)
    for col, w in widths.items(): ws.column_dimensions[col].width = w
    def merged(r, text, font, height=True):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last)
        c = ws.cell(row=r, column=1, value=text); c.font = font
        c.alignment = ZENTR if r == 1 else WRAP
        if height: ws.row_dimensions[r].height = zeilenhoehe(text, int(cum * 0.95))
    merged(1, spec["titel"], F12B)
    merged(2, spec["sit"], F10)
    merged(3, "Ihre Aufträge", F10B)
    for i, a in enumerate(spec["auftr"]): merged(4 + i, f"{i + 1}. {a}", F10)
    merged(11, spec.get("hinweis", BASIS_HINT), F10I)
    # Tabellenkopf
    for i, c in enumerate(cols):
        h = ws.cell(row=HR, column=i + 1, value=c[0]); h.font = F10B; h.fill = GRAU; h.border = RAHMEN; h.alignment = ZENTR
    ws.row_dimensions[HR].height = 30
    rows = spec["rows"]
    for k, row in enumerate(rows):
        r = R0 + k
        it = iter(row)
        for i, c in enumerate(cols):
            cell = ws.cell(row=r, column=i + 1); cell.font = F10; cell.border = RAHMEN
            if c[3] == "d":
                cell.value = next(it)
                if c[2]: cell.number_format = c[2]
            elif c[3] == "x":
                cell.value = c[4].format(r=r) if loesung else c[5](r)
                cell.number_format = c[2] or "General"
            elif loesung:
                cell.value = c[4].format(r=r); cell.number_format = c[2] or "General"
    for e in spec.get("extras", []):
        kind = e[0]
        if kind == "label":
            c = ws[e[1]]; c.value = e[2]; c.font = F10B if (len(e) > 3 and e[3]) else F10; c.border = RAHMEN
            c.alignment = Alignment(wrap_text=True, vertical="center")
        elif kind == "input":
            c = ws[e[1]]; c.value = e[2]; c.font = F10; c.border = RAHMEN
            if e[3]: c.number_format = e[3]
        elif kind == "target":
            c = ws[e[1]]; c.border = RAHMEN; c.font = F10B if (len(e) > 4 and e[4]) else F10
            if loesung:
                c.value = e[2]
                if e[3]: c.number_format = e[3]
        elif kind == "answer":
            c = ws[e[1]]; c.border = RAHMEN; c.font = F10; c.alignment = WRAP
            if loesung: c.value = e[2]
            if len(e) > 4 and e[4]:
                ws.merge_cells(e[4])
        elif kind == "table":
            _, r0, c0, heads, trows, fmts = e
            for i, h in enumerate(heads):
                x = ws.cell(row=r0, column=c0 + i, value=h); x.font = F10B; x.fill = GRAU; x.border = RAHMEN; x.alignment = ZENTR
            for k, row in enumerate(trows):
                for i, v in enumerate(row):
                    x = ws.cell(row=r0 + 1 + k, column=c0 + i, value=v); x.font = F10; x.border = RAHMEN
                    if fmts and i < len(fmts) and fmts[i]: x.number_format = fmts[i]
        elif kind == "ftable":   # Tabelle mit Formelspalten: (Zielspalten bleiben leer)
            pass
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    if loesung:
        ws.oddFooter.right.text = "Vorname Nachname, Klasse"; ws.oddFooter.right.size = 10; ws.oddFooter.right.font = "Arial"
    else:
        pass
    return ws


def add_overview(path, specs, nr_first):
    wb = load_workbook(path)
    ws = wb["Aufgabenstellung"]
    r = ws.max_row + 2
    names = "„" + "“, „".join(s["name"] for s in specs) + "“"
    for text, font in (("Weitere Aufgaben", F10B),
                       (f"Zu diesem Modul gehören die zusätzlichen Tabellenblätter {names}. Jedes Blatt enthält eine eigene Ausgangssituation, Aufträge und eine Tabelle. "
                        "Die Schwierigkeit steigt von Blatt zu Blatt.", F10)):
        c = ws.cell(row=r, column=1, value=text); c.font = font; c.alignment = WRAP
        ws.row_dimensions[r].height = zeilenhoehe(text); r += 1
    return wb


def md_block(spec, nr, vals, loesung_wb):
    """Baut den Markdown-Block für eine Aufgabe. vals: Zellwerte aus dem neu berechneten Lösungsblatt."""
    ws = vals
    lines = [f":::aufgabe {spec['niveau']} {spec['min']} {nr} · {spec['kurz']}",
             f"Öffnen Sie das Blatt „{spec['name']}“ und lösen Sie die Aufträge:", ""]
    for a in spec["auftr"]: lines.append(f"- {a}")
    lines += ["", ":::checkliste Selbstcheck"]
    fcols = [(i, c) for i, c in enumerate(spec["cols"]) if c[3] in ("f", "x")]
    for i, c in fcols:
        if c[3] == "x":
            lines.append(f"- Spalte {L(i + 1)} („{c[0]}“): Der Fehler ist gefunden, behoben und in allen Zeilen korrigiert.")
        else:
            lines.append(f"- Spalte {L(i + 1)} („{c[0]}“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.")
    for t in spec.get("extras", []):
        if t[0] == "target" and len(t) > 5 and t[5]:
            pass
    absol = any("$" in c[4] for _, c in fcols if c[3] == "f")
    lines.append("- Alle Ergebnisse sind Formeln mit Zellbezügen." + (" Bezüge, die beim Kopieren nicht wandern sollen, sind mit Dollarzeichen festgesetzt." if absol else ""))
    for chk in spec.get("checks", []): lines.append(f"- {chk}")
    lines += [":::", ""]
    if spec.get("hilfe"):
        for k, h in enumerate(spec["hilfe"], 1):
            title = f"Hilfe {k}" + (": " + h[0] if h[0] else "")
            lines += [f":::loesung {title}", h[1], ":::", ""]
    # Kontrollwerte
    ko = []
    nrows = len(spec["rows"])
    for i, c in fcols:
        vals = [fmtval(ws[f"{L(i + 1)}{R0 + k}"].value, c[2]) for k in range(nrows)]
        ko.append(f"- Spalte {L(i + 1)} ({L(i + 1)}{R0} bis {L(i + 1)}{R0 + nrows - 1}): " + "; ".join(vals))
    tg = [e for e in spec.get("extras", []) if e[0] in ("target", "answer")]
    for e in tg:
        if e[0] == "target":
            lab = e[5] if len(e) > 5 and e[5] else None
            txt = fmtval(ws[e[1]].value, e[3])
        else:
            lab = e[3] if len(e) > 3 and e[3] else "Antwort"
            txt = e[2].rstrip(".")
        ko.append(f"- {lab} ({e[1]}): {txt}" if lab else f"- {e[1]}: {txt}")
    lines += [":::loesung Kontrolle: Erwartete Ergebnisse", "\n".join(ko), ":::", ""]
    lo = []
    for i, c in fcols:
        if c[3] == "x":
            lo.append(f"Spalte {L(i + 1)}: Fehler: {c[6]} Richtig (ab {L(i + 1)}{R0}): `{de(c[4].format(r=R0))}`")
        else:
            lo.append(f"Spalte {L(i + 1)} (ab {L(i + 1)}{R0}): `{de(c[4].format(r=R0))}`")
    for e in tg:
        if e[0] == "target": lo.append(f"{e[1]}: `{de(e[2])}`")
    lines += [":::loesung Lösung", ". ".join(lo) + ".", ":::", ":::", ""]
    return "\n".join(lines)


def block_ranges(lines):
    res = []
    i = 0
    while i < len(lines):
        if lines[i].startswith(":::aufgabe"):
            d = 0; j = i
            while j < len(lines):
                ln = lines[j].rstrip()
                if ln == ":::": d -= 1
                elif ln.startswith(":::"): d += 1
                if d == 0: break
                j += 1
            res.append((i, j)); i = j + 1
        else: i += 1
    return res


def patch_md(path, new_blocks_text, extra_min, nr_kreativ, names_text, rename=True):
    t = open(path, encoding="utf-8").read()
    head, rest = t.split("\n---\n", 1)
    m = re.search(r"^minuten: (\d+)", head, re.M)
    head = head.replace(m.group(0), f"minuten: {int(m.group(1)) + extra_min}")
    lines = rest.split("\n")
    rng = block_ranges(lines)
    # Plus-Titel anpassen und Kreativ nummerieren
    kre_start = None
    for a, b in rng:
        ln = lines[a]
        mm = re.match(r"^:::aufgabe plus (\S+) \d+ · (.*)$", ln)
        if mm and rename: lines[a] = f":::aufgabe plus {mm.group(1)} Zusatz zu Aufgabe 1 · {mm.group(2)}"
        mk = re.match(r"^:::aufgabe kreativ (\S+) \d+ · (.*)$", ln)
        if mk: lines[a] = f":::aufgabe kreativ {mk.group(1)} {nr_kreativ} · {mk.group(2)}"; kre_start = a
    assert kre_start is not None, path
    lines[kre_start:kre_start] = new_blocks_text.split("\n")
    # Material-Hinweis
    for i, ln in enumerate(lines):
        if ln.startswith(":::material"):
            j = i + 1
            while lines[j].strip() != ":::": j += 1
            lines[j - 1] = lines[j - 1].rstrip() + " " + names_text
            break
    open(path, "w", encoding="utf-8").write(head + "\n---\n" + "\n".join(lines))
