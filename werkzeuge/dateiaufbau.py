"""Gemeinsamer Dateiaufbau für alle Übungsdateien: Startblatt (wie die erste Seite einer Prüfung) und Blatt
„Corporate Design und Abgabe“. Texte und Zahlen kommen aus modellunternehmen/stammdaten.yaml (werkzeuge/unternehmen.py)."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import unternehmen as U
from ext_framework import zeilenhoehe
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

NAVY, MITTE, HELL, BLASS, GELB = "1F3864", "2F5597", "DCE6F1", "F2F2F2", "FFF2CC"
LINKS = Alignment(horizontal="left", vertical="top", wrap_text=True)
RAND = Border(left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"),
              top=Side(style="thin", color="BFBFBF"), bottom=Side(style="thin", color="BFBFBF"))


def f(size=10, bold=False, italic=False, color="000000"):
    return Font(name="Arial", size=size, bold=bold, italic=italic, color=color)


def fill(c):
    return PatternFill("solid", fgColor=c)


def euro_text(euro):
    return "Eurobeträge mit zwei Dezimalstellen" if euro == "zwei" else "Eurobeträge mit 1000er-Trennzeichen ohne Dezimalstellen"


def cd_zeilen(euro="zwei"):
    """Corporate Design (Tabellenkalkulation) als Textzeilen, wie auf der Prüfungsseite."""
    c = U.CD
    return [f"Schriftart {c['schrift']}, Schriftgröße {c['schriftgroesse']}",
            f"Überschriften {c['ueberschrift']['ausrichtung']}, Schriftgröße {c['ueberschrift']['groesse']}, Fettschrift",
            f"Diagramm mit {c['diagramm']['schrift']}; Diagrammtitel in Schriftgröße {c['diagramm']['titel_groesse']}",
            "Format Zahl:",
            f"•  Prozentsätze mit {'zwei' if c['prozent_nachkommastellen'] == 2 else c['prozent_nachkommastellen']} Dezimalstellen",
            f"•  {euro_text(euro)}",
            "•  Andere Zahlen ohne Dezimalstellen"]


def start_blatt(wb, titel, thema, dateiname, specs, loesung, euro="zwei"):
    ws = wb.create_sheet("Start", 0)
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 2, "B": 36, "C": 12, "D": 11, "E": 62}.items():
        ws.column_dimensions[col].width = w

    def zeile(r, text, font, fuell=None, hoehe=None, von="B", bis="E", rand=False):
        ws.merge_cells(f"{von}{r}:{bis}{r}")
        c = ws[f"{von}{r}"]
        c.value = text
        c.font = font
        c.alignment = LINKS
        if fuell:
            for col in "BCDE":
                ws[f"{col}{r}"].fill = fill(fuell)
        if rand:
            for col in "BCDE":
                ws[f"{col}{r}"].border = RAND
        ws.row_dimensions[r].height = hoehe or zeilenhoehe(text, 112)

    zeile(1, f"Übungsdatei: {titel}", f(18, True, color="FFFFFF"), NAVY, 40)
    ws["B1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    zeile(2, f"{U.FIRMA}  ·  {U.D['firma']['beruf']}  ·  Lernpfad Excel-Funktionen", f(10, italic=True, color="FFFFFF"), MITTE, 20)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    zeile(4, "Ihre Situation", f(11, True, color=NAVY), HELL, 20)
    zeile(5, f"Sie arbeiten in der Abteilung {U.ABTEILUNG} der {U.ADRESSE.split(', ', 1)[0]}, {U.ADRESSE.split(', ', 1)[1]}. {thema}", f(10), rand=True)

    zeile(7, "Sie benötigen folgende Dateien", f(11, True, color=NAVY), HELL, 20)
    for col, h in zip("BCE", ["Ordner", "Datei", "Tabellenblätter"]):
        c = ws[f"{col}8"]; c.value = h; c.font = f(10, True); c.fill = fill(BLASS)
    ws.merge_cells("C8:D8")
    for col in "BCDE": ws[f"{col}8"].border = RAND
    ws.merge_cells("C9:D9")
    for col, v in zip("BCE", ["Übungsdateien", dateiname, ", ".join(s["name"] for s in specs)]):
        c = ws[f"{col}9"]; c.value = v; c.font = f(10); c.alignment = LINKS
    for col in "BCDE": ws[f"{col}9"].border = RAND
    ws.row_dimensions[9].height = zeilenhoehe(", ".join(s["name"] for s in specs), 62) + 4

    zeile(11, "Corporate Design (Tabellenkalkulation)", f(11, True, color=NAVY), HELL, 20)
    r = 12
    for t in cd_zeilen(euro):
        zeile(r, t, f(10, t == "Format Zahl:"), rand=True)
        r += 1
    zeile(r, "Die Zellen, die Sie ausfüllen, liegen im Format „Standard“ vor. Ausführliche Hinweise zum Formatieren stehen im Blatt „Corporate Design und Abgabe“.", f(9, italic=True, color="595959"))
    r += 2

    zeile(r, "So arbeiten Sie", f(11, True, color=NAVY), HELL, 20)
    r += 1
    for i, s in enumerate(["Öffnen Sie nacheinander die Tabellenblätter in der Liste unten. Auf jedem Blatt steht oben der Auftrag mit der Situation.",
                           "Tragen Sie Ihre Formeln in die leeren Ergebniszellen ein. Verwenden Sie Zellbezüge und keine eingetippten Ergebnisse.",
                           "Prüfen Sie Ihre Ergebnisse mit den Selbstchecks im Modul und formatieren Sie nach dem Corporate Design.",
                           "Zum Schluss lesen Sie das Blatt „Corporate Design und Abgabe“."]):
        zeile(r, f"{i + 1}.  {s}", f(10))
        r += 1
    r += 1
    zeile(r, "Ihre Tabellenblätter", f(11, True, color=NAVY), HELL, 20)
    r += 1
    for col, h in zip("BCDE", ["Blatt (anklicken)", "Niveau", "Zeit", "Thema"]):
        c = ws[f"{col}{r}"]; c.value = h; c.font = f(10, True); c.fill = fill(BLASS); c.border = RAND
    r += 1
    eintraege = [(s["name"], {"basis": "Basis", "plus": "Plus"}.get(s["niveau"], s["niveau"]), f"{s['min']} Min.", s["kurz"]) for s in specs]
    eintraege.append(("Corporate Design und Abgabe", "Pflicht", "5 Min.", "Gestaltung und Abgabe der Datei"))
    for name, lab, zeit, kurz in eintraege:
        for col, v in zip("BCDE", [name, lab, zeit, kurz]):
            c = ws[f"{col}{r}"]; c.value = v; c.font = f(10); c.border = RAND; c.alignment = LINKS
        ws[f"B{r}"].hyperlink = Hyperlink(ref=f"B{r}", location=f"'{name}'!A1", display=name)
        ws[f"B{r}"].font = Font(name="Arial", size=10, underline="single", color="0563C1")
        r += 1
    r += 1
    if loesung:
        zeile(r, "NUR FÜR LEHRKRÄFTE: Musterlösung mit Formeln und Kontrollwerten. Nicht an Lernende ausgeben.", f(10, True, color="C00000"), GELB, 22)
    ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def design_abgabe_blatt(wb, loesung, euro="zwei"):
    ws = wb.create_sheet("Corporate Design und Abgabe")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 2, "B": 30, "C": 58, "D": 22}.items():
        ws.column_dimensions[col].width = w

    def kopf(r, text):
        ws.merge_cells(f"B{r}:D{r}")
        c = ws[f"B{r}"]; c.value = text; c.font = f(11, True, color=NAVY); c.alignment = Alignment(vertical="center", indent=1)
        for col in "BCD": ws[f"{col}{r}"].fill = fill(HELL)
        ws.row_dimensions[r].height = 22

    def text(r, t):
        ws.merge_cells(f"B{r}:D{r}")
        c = ws[f"B{r}"]; c.value = t; c.font = f(10); c.alignment = LINKS
        ws.row_dimensions[r].height = zeilenhoehe(t, 108)

    ws.merge_cells("B1:D1")
    ws["B1"].value = "Corporate Design und Abgabe"; ws["B1"].font = f(18, True, color="FFFFFF"); ws["B1"].alignment = Alignment(vertical="center", indent=1)
    for col in "BCD": ws[f"{col}1"].fill = fill(NAVY)
    ws.row_dimensions[1].height = 40
    kopf(3, f"Corporate Design der {U.FIRMA}")
    text(4, "Formatieren Sie Ihre Ergebnisse nach diesen Vorgaben. Welches Format ein Auftrag verlangt, steht im Auftrag. Das Beispiel zeigt, wie das Ergebnis aussehen soll.")
    for col, h in zip("BCD", ["Element", "Vorgabe", "Beispiel"]):
        c = ws[f"{col}5"]; c.value = h; c.font = f(10, True); c.fill = fill(BLASS); c.border = RAND
    cd = U.CD
    regeln = [("Schrift", f"Schriftart {cd['schrift']}, Schriftgröße {cd['schriftgroesse']}", "Arial 10", None),
              ("Überschriften", f"{cd['ueberschrift']['ausrichtung']}, Schriftgröße {cd['ueberschrift']['groesse']}, Fettschrift", "Überschrift", "ueb"),
              ("Diagramme", f"voreingestellte Schriftart; Diagrammtitel in Schriftgröße {cd['diagramm']['titel_groesse']}", "", None),
              ("Prozentsätze", "mit zwei Dezimalstellen", 0.1234, "0.00%"),
              ("Eurobeträge", euro_text(euro), 1234.5, cd["euro"][euro]),
              ("Andere Zahlen", "ohne Dezimalstellen", 12345, "#,##0"),
              ("Datum", f"im Format {cd['datum']}", 46150, "DD.MM.YYYY"),
              ("Uhrzeit", f"im Format {cd['uhrzeit']}", 0.3541666667, "hh:mm")]
    r = 6
    for el, vor, bsp, fmt in regeln:
        a, b, c = ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"]
        a.value, b.value = el, vor
        a.font = f(10, True); b.font = f(10)
        c.value = bsp
        if fmt == "ueb":
            c.font = f(12, True); c.alignment = Alignment(horizontal="center")
        else:
            c.font = f(10)
            if fmt: c.number_format = fmt
        for x in (a, b, c): x.border = RAND
        r += 1
    r += 1
    kopf(r, "So formatieren Sie (Microsoft Excel)"); r += 1
    for i, s in enumerate(["Markieren Sie die Zellen, die Sie formatieren möchten.",
                           "Öffnen Sie mit Strg+1 das Fenster „Zellen formatieren“ oder klicken Sie auf Start › Zahl.",
                           "Wählen Sie das Zahlenformat (Prozent, Zahl, Währung, Datum, Uhrzeit) und die Nachkommastellen.",
                           "Schrift und Ausrichtung stellen Sie auf der Registerkarte Start ein."], 1):
        text(r, f"{i}.  {s}"); r += 1
    r += 1
    kopf(r, "Abgabe"); r += 1
    for i, s in enumerate(["Ergänzen Sie in der Fußzeile rechtsbündig Ihren Namen und Ihre Klassenbezeichnung (Seitenlayout › Kopf- und Fußzeile).",
                           "Speichern Sie die Datei als Excel-Arbeitsmappe (.xlsx) unter dem Namen „<Dateiname>_Nachname“.",
                           "Laden Sie die Datei in Moodle hoch, wenn eine Abgabe verlangt ist."], 1):
        text(r, f"{i}.  {s}"); r += 1
    r += 1
    kopf(r, "Checkliste vor der Abgabe"); r += 1
    for s in ["Alle Ergebnisse sind Formeln mit Zellbezügen, keine eingetippten Zahlen.",
              "Zahlenformate, Schrift und Überschriften entsprechen dem Corporate Design.",
              "Die Fußzeile enthält Name und Klasse (rechtsbündig).",
              "Die Datei ist als .xlsx gespeichert und trägt Ihren Nachnamen."]:
        text(r, "☐  " + s); r += 1
    ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
