import sys, os, subprocess, importlib, re
ARGS = sys.argv[1:]
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ext_framework import *
import ext_framework as F
KURS = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
ORIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'md_orig')
STEM = {"r1":"r1-umsatz","r2":"r2-preisliste","r3":"r3-bestellungen","r4":"r4-rabatt","r5":"r5-kunden","r6":"r6-bonus","rm":"rm-kalkulation",
        "z1":"z1-inventur","z2":"z2-bestellungen","z3":"z3-verkaeufe","zm":"zm-auswertung",
        "n1":"n1-angebot","n2":"n2-kundendaten","n3":"n3-runden","n4":"n4-rang","n5":"n5-datum-zeit","nm":"nm-rechnung"}


def overview(wb, specs):
    ws = wb["Aufgabenstellung"]
    for row in ws.iter_rows(min_col=1, max_col=1):
        if row[0].value == "Weitere Aufgaben": return
    r = ws.max_row + 2
    names = "„" + "“, „".join(s["name"] for s in specs) + "“"
    for text, font in (("Weitere Aufgaben", F10B),
                       (f"Zu diesem Modul gehören die zusätzlichen Tabellenblätter {names}. Jedes Blatt enthält eine eigene Ausgangssituation, Aufträge und eine Tabelle. "
                        "Die Schwierigkeit steigt von Blatt zu Blatt.", F10)):
        c = ws.cell(row=r, column=1, value=text); c.font = font; c.alignment = WRAP
        ws.row_dimensions[r].height = zeilenhoehe(text); r += 1


def patch_milestone(key, block, names):
    t = open(f"{ORIG}/{key}.md", encoding="utf-8").read()
    m = re.search(r"^## \d+\. Abgabe", t, re.M)
    sect = ("## Zusatz für Schnelle (freiwillig)\n\nWer vor dem Ende der Zeit fertig ist, kann die folgende Erweiterung bearbeiten. Sie wird nicht für die Bewertung des Meilensteins herangezogen. " + names + "\n\n" + block + "\n")
    t = t[:m.start()] + sect + "\n" + t[m.start():]
    open(f"{KURS}/module/{key}.md", "w", encoding="utf-8").write(t)


def run(key, specs, patch=True):
    stem = STEM[key]
    pa = f"{KURS}/material/{stem}.xlsx"; pl = f"{KURS}/lehrkraft/loesungen/{stem}-loesung.xlsx"
    for loes, p in ((False, pa), (True, pl)):
        wb = load_workbook(p)
        for s in specs:
            if s["name"] in wb.sheetnames: del wb[s["name"]]
        if "Weitere Aufgaben" in str(wb["Aufgabenstellung"].cell(row=wb["Aufgabenstellung"].max_row - 1, column=1).value or ""):
            pass
        for s in specs: build_sheet(wb, s, loes, len(wb.sheetnames))
        overview(wb, specs)
        wb.save(p)
    if os.path.exists("/mnt/skills/public/xlsx/scripts/recalc.py"):
        out = subprocess.run(["python3", "/mnt/skills/public/xlsx/scripts/recalc.py", pl, "60"], capture_output=True, text=True)
        print(key, out.stdout.strip()[:300], out.stderr.strip()[:200])
    else:  # ohne die Skill-Umgebung: LibreOffice rechnet neu
        import tempfile, shutil
        tmp = tempfile.mkdtemp()
        subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, pl], check=True, capture_output=True, timeout=180)
        shutil.copy(os.path.join(tmp, os.path.basename(pl)), pl)
    vals = load_workbook(pl, data_only=True)
    blocks = []
    for i, s in enumerate(specs, 2):
        blocks.append(md_block(s, s.get('nr', i), vals[s["name"]], None))
    if patch:
        extra = sum(s["min"] for s in specs if s["niveau"] == "basis")
        if len(specs) == 1: names = f"Die Erweiterung finden Sie im Blatt „{specs[0]['name']}“ derselben Datei."
        else: names = "Zusätzlich enthält die Datei die Blätter " + ", ".join(f"„{s['name']}“" for s in specs) + ". Die Schwierigkeit steigt von Blatt zu Blatt."
        shutil_copy = open(f"{ORIG}/{key}.md", encoding="utf-8").read()
        open(f"{KURS}/module/{key}.md", "w", encoding="utf-8").write(shutil_copy)
        if key in ("rm","zm","nm"):
            patch_milestone(key, blocks[0], names); return vals
        patch_md(f"{KURS}/module/{key}.md", "\n".join(blocks), extra, max(s.get('nr', 0) for s in specs) + 1 if any('nr' in s for s in specs) else len(specs) + 2, names, rename=not any('nr' in s for s in specs))
    # Fehlerprüfung
    bad = []
    for ws in vals:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#"): bad.append((ws.title, c.coordinate, c.value))
    print(key, "Fehlerzellen:", bad)
    return vals

if __name__ == "__main__":
    mod = importlib.import_module(ARGS[0])
    keys = ARGS[1:] or list(mod.SPECS)
    for k in keys: run(k, mod.SPECS[k])
