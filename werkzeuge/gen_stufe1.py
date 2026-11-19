#!/usr/bin/env python3
"""Erzeugt die Übungsdateien (Aufgabe + Lösung) für Stufe 1 des Lernpfads Excel-Funktionen.
Formeln mit englischen Namen und Komma (openpyxl); Excel/OnlyOffice zeigen sie je nach Oberflächensprache deutsch an.
Corporate Design BüroComfort GmbH: Arial 10, Titel zentriert 12 fett, Prozent 2 Nachkommastellen,
Zahlen ohne Nachkommastellen, Euro mit Tausenderpunkt und 2 Nachkommastellen.
Zellen, die Lernende selbst ausfüllen, bleiben in der Aufgabe unformatiert (Standard)."""
import math
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT_AUFGABE = sys.argv[1]
OUT_LOESUNG = sys.argv[2]
os.makedirs(OUT_AUFGABE, exist_ok=True)
os.makedirs(OUT_LOESUNG, exist_ok=True)

EURO = '#,##0.00 "€"'
PROZ = '0.00%'
ZAHL = '#,##0'
F10 = Font(name="Arial", size=10)
F10B = Font(name="Arial", size=10, bold=True)
F12B = Font(name="Arial", size=12, bold=True)
F10I = Font(name="Arial", size=10, italic=True)
GRAU = PatternFill("solid", fgColor="E7E6E6")
DUENN = Side(style="thin", color="808080")
RAHMEN = Border(left=DUENN, right=DUENN, top=DUENN, bottom=DUENN)
WRAP = Alignment(wrap_text=True, vertical="top")
ZENTR = Alignment(horizontal="center", vertical="center", wrap_text=True)


def zeilenhoehe(text, breite=110):
    zeilen = 0
    for teil in str(text).split("\n"):
        zeilen += max(1, math.ceil(len(teil) / breite))
    return 13.5 * zeilen + 3


def aufgabenblatt(wb, titel, situation, auftraege, hinweis, loesung_hinweis=None, loesung=False):
    ws = wb.active
    ws.title = "Aufgabenstellung"
    ws.column_dimensions["A"].width = 110
    zeilen = [
        (titel, F12B, ZENTR),
        ("Ausbildungsberuf: Kaufmann/Kauffrau für Büromanagement  |  Lernpfad: Excel-Funktionen  |  Sozialform: Einzelarbeit", F10I, ZENTR),
        ("", F10, WRAP),
        ("Ausgangssituation", F10B, WRAP),
        (situation, F10, WRAP),
        ("", F10, WRAP),
        ("Ihre Aufträge", F10B, WRAP),
    ]
    for i, a in enumerate(auftraege, 1):
        zeilen.append((f"{i}. {a}", F10, WRAP))
    zeilen += [
        ("", F10, WRAP),
        ("Hinweis", F10B, WRAP),
        (hinweis, F10, WRAP),
        ("", F10, WRAP),
        ("Corporate Design der BüroComfort GmbH", F10B, WRAP),
        ("Die Zellen, die Sie ausfüllen, liegen im Format „Standard“ vor. Formatieren Sie Ihre Ergebnisse selbst: Schriftart Arial, "
         "Schriftgröße 10; Überschriften zentriert, Schriftgröße 12, fett; Prozentsätze mit zwei Nachkommastellen; "
         "andere Zahlen ohne Nachkommastellen; Eurobeträge mit Tausenderpunkt und zwei Nachkommastellen.", F10, WRAP),
        ("", F10, WRAP),
        ("Abgabe", F10B, WRAP),
        ("Ergänzen Sie in der Fußzeile rechtsbündig Ihren Namen und Ihre Klassenbezeichnung. "
         "Speichern Sie die Datei unter dem Namen „<Dateiname>_Nachname“ und laden Sie sie in Moodle hoch, wenn eine Abgabe verlangt ist.", F10, WRAP),
    ]
    if loesung and loesung_hinweis:
        zeilen += [("", F10, WRAP), ("NUR FÜR LEHRKRÄFTE – nicht an Lernende ausgeben", F10B, WRAP), (loesung_hinweis, F10, WRAP)]
    for r, (text, font, al) in enumerate(zeilen, 1):
        c = ws.cell(row=r, column=1, value=text)
        c.font = font
        c.alignment = al
        ws.row_dimensions[r].height = zeilenhoehe(text)
    return ws


def tabelle_kopf(ws, zeile, ueberschriften, start_col=1):
    for i, h in enumerate(ueberschriften):
        c = ws.cell(row=zeile, column=start_col + i, value=h)
        c.font = F10B
        c.fill = GRAU
        c.border = RAHMEN
        c.alignment = ZENTR


def daten_zellen(ws, start_zeile, zeilen, formate, start_col=1):
    for r, row in enumerate(zeilen, start_zeile):
        for i, wert in enumerate(row):
            c = ws.cell(row=r, column=start_col + i, value=wert)
            c.font = F10
            c.border = RAHMEN
            f = formate[i] if i < len(formate) else None
            if f:
                c.number_format = f


def titel_zeile(ws, text, bis_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=bis_col)
    c = ws.cell(row=1, column=1, value=text)
    c.font = F12B
    c.alignment = ZENTR


def ziel(ws, adresse, formel, fmt=None, loesung=False, fett=False):
    """Zielzelle: in der Aufgabe unberührt, in der Lösung Formel + Corporate Design."""
    if not loesung:
        return
    c = ws[adresse]
    c.value = formel
    c.font = F10B if fett else F10
    c.border = RAHMEN
    if fmt:
        c.number_format = fmt


def beschriftung(ws, adresse, text, fett=False):
    c = ws[adresse]
    c.value = text
    c.font = F10B if fett else F10
    c.border = RAHMEN


def breiten(ws, werte):
    for spalte, w in werte.items():
        ws.column_dimensions[spalte].width = w


def fuss(ws, loesung):
    if loesung:
        ws.oddFooter.right.text = "Vorname Nachname, Klasse"
        ws.oddFooter.right.size = 10
        ws.oddFooter.right.font = "Arial"


def speichern(wb, name, loesung):
    ordner = OUT_LOESUNG if loesung else OUT_AUFGABE
    suffix = "-loesung" if loesung else ""
    pfad = os.path.join(ordner, f"{name}{suffix}.xlsx")
    wb.save(pfad)
    return pfad


# ------------------------------------------------------------------ r1
def r1(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Tagesumsätze auswerten in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Geschäftsleitung benötigt für die Wochenbesprechung eine Übersicht über die Umsätze der vier Filialen in der Kalenderwoche 12. "
        "Die Tagesumsätze liegen im Tabellenblatt „Umsätze KW 12“ vor.",
        ["Ermitteln Sie für jeden Tag den Gesamtumsatz aller Filialen (Spalte F).",
         "Ermitteln Sie für jede Filiale den Wochenumsatz (Zeile 10) und den Gesamtumsatz der Woche (Zelle F10).",
         "Ermitteln Sie für jede Filiale und für den Gesamtumsatz den durchschnittlichen Tagesumsatz (Zeile 11).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Für Summen und Durchschnitte gibt es eigene Funktionen. Prüfen Sie am Ende, ob die Ergebnisse zu den Einzelwerten passen.",
        "Lösung: F4:F9 =SUMME(B4:E4) usw.; B10:E10 =SUMME(B4:B9); F10 =SUMME(F4:F9); B11:F11 =MITTELWERT(B4:B9) usw. "
        "Gesamtumsatz der Woche: 107.130,00 €.", loesung)
    ws = wb.create_sheet("Umsätze KW 12")
    titel_zeile(ws, "BüroComfort GmbH – Tagesumsätze der Filialen in KW 12", 6)
    tabelle_kopf(ws, 3, ["Tag", "Rendsburg", "Kiel", "Flensburg", "Neumünster", "Gesamt"])
    daten = [("Montag", 4120, 5380, 3290, 2760), ("Dienstag", 3980, 5120, 3410, 2890),
             ("Mittwoch", 4410, 5760, 3620, 3050), ("Donnerstag", 4290, 5490, 3550, 2980),
             ("Freitag", 5230, 6840, 4120, 3470), ("Samstag", 6180, 7920, 5060, 4210)]
    daten_zellen(ws, 4, daten, [None, EURO, EURO, EURO, EURO])
    beschriftung(ws, "A10", "Wochenumsatz", True)
    beschriftung(ws, "A11", "Ø Tagesumsatz", True)
    for r in range(4, 10):
        ziel(ws, f"F{r}", f"=SUM(B{r}:E{r})", EURO, loesung)
    for col in "BCDE":
        ziel(ws, f"{col}10", f"=SUM({col}4:{col}9)", EURO, loesung)
        ziel(ws, f"{col}11", f"=AVERAGE({col}4:{col}9)", EURO, loesung)
    ziel(ws, "F10", "=SUM(F4:F9)", EURO, loesung, True)
    ziel(ws, "F11", "=AVERAGE(F4:F9)", EURO, loesung, True)
    breiten(ws, {"A": 16, "B": 14, "C": 14, "D": 14, "E": 14, "F": 16})
    fuss(ws, loesung)
    return speichern(wb, "r1-umsatz", loesung)


# ------------------------------------------------------------------ r2
def r2(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Preisliste prüfen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Der Einkauf möchte wissen, in welcher Preisspanne sich das Sortiment bewegt. Die aktuelle Preisliste liegt im Tabellenblatt „Preisliste“ vor.",
        ["Ermitteln Sie den niedrigsten Preis (Zelle C16) und den höchsten Preis (Zelle C17) der Preisliste.",
         "Tragen Sie in die Zellen B16 und B17 ein, um welche Artikel es sich handelt.",
         "Berechnen Sie die Preisspanne, also den Unterschied zwischen höchstem und niedrigstem Preis (Zelle C18).",
         "Berechnen Sie den Durchschnittspreis aller Artikel (Zelle C19).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Beantworten Sie die Fragen mit Funktionen und nicht durch Hinsehen. Wenn sich ein Preis ändert, soll sich Ihr Ergebnis automatisch anpassen.",
        "Lösung: C16 =MIN(C4:C15) = 0,89 €; C17 =MAX(C4:C15) = 489,90 €; B16 Kugelschreiber; B17 Schreibtisch höhenverstellbar; "
        "C18 =C17-C16 = 489,01 €; C19 =MITTELWERT(C4:C15) ≈ 92,84 €.", loesung)
    ws = wb.create_sheet("Preisliste")
    titel_zeile(ws, "BüroComfort GmbH – Preisliste (Verkaufspreise)", 3)
    tabelle_kopf(ws, 3, ["Art.-Nr.", "Artikel", "Preis"])
    daten = [("A-1001", "Schreibtischstuhl Basic", 129.00), ("A-1002", "Schreibtisch höhenverstellbar", 489.90),
             ("A-1003", "Aktenschrank 2 Ordnerhöhen", 219.00), ("A-1004", "Ordner A4 breit", 2.49),
             ("A-1005", "Kugelschreiber", 0.89), ("A-1006", "Kopierpapier 500 Blatt", 4.95),
             ("A-1007", "Locher", 8.90), ("A-1008", "Tacker", 6.50), ("A-1009", "Monitorständer", 34.90),
             ("A-1010", "Schreibtischlampe LED", 42.50), ("A-1011", "Rollcontainer", 119.00),
             ("A-1012", "Whiteboard 90 x 60 cm", 56.00)]
    daten_zellen(ws, 4, daten, [None, None, EURO])
    beschriftung(ws, "A16", "Niedrigster Preis", True)
    beschriftung(ws, "A17", "Höchster Preis", True)
    beschriftung(ws, "A18", "Preisspanne", True)
    beschriftung(ws, "A19", "Durchschnittspreis", True)
    ziel(ws, "C16", "=MIN(C4:C15)", EURO, loesung)
    ziel(ws, "C17", "=MAX(C4:C15)", EURO, loesung)
    ziel(ws, "C18", "=C17-C16", EURO, loesung)
    ziel(ws, "C19", "=AVERAGE(C4:C15)", EURO, loesung)
    if loesung:
        ws["B16"].value = "Kugelschreiber"
        ws["B17"].value = "Schreibtisch höhenverstellbar"
        for a in ("B16", "B17"):
            ws[a].font = F10
            ws[a].border = RAHMEN
    breiten(ws, {"A": 20, "B": 32, "C": 16})
    fuss(ws, loesung)
    return speichern(wb, "r2-preisliste", loesung)


# ------------------------------------------------------------------ r3
def r3(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Versandkosten prüfen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Der Vertrieb möchte in der Bestellliste sofort sehen, welche Bestellungen versandkostenfrei sind. "
        "Versandkostenfrei sind Bestellungen mit einem Bestellwert über 50,00 €. Alle anderen Bestellungen werden mit Versandkosten berechnet.",
        ["Zeigen Sie in Spalte D für jede Bestellung den Text „Versandkostenfrei“ oder „Versandkosten“ an.",
         "Prüfen Sie die Bestellung B-2004 mit genau 50,00 € besonders genau. Begründen Sie in Zelle F4 in einem Satz, warum Ihr Ergebnis stimmt.",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Die Anzeige soll sich automatisch ändern, wenn sich ein Bestellwert ändert.",
        "Lösung: D4 =WENN(C4>50;\"Versandkostenfrei\";\"Versandkosten\"). B-2004 (50,00 €) ergibt „Versandkosten“, weil 50 nicht größer als 50 ist. "
        "Plus (Spalte E): =WENN(C4>50;0;4,95).", loesung)
    ws = wb.create_sheet("Bestellungen")
    titel_zeile(ws, "BüroComfort GmbH – Bestellungen mit Versandprüfung", 5)
    tabelle_kopf(ws, 3, ["Bestell-Nr.", "Kunden-Nr.", "Bestellwert", "Versand", "Versandkosten (Plus)"])
    daten = [("B-2001", "K-2301", 24.90), ("B-2002", "K-2302", 67.80), ("B-2003", "K-2303", 49.99),
             ("B-2004", "K-2304", 50.00), ("B-2005", "K-2305", 128.50), ("B-2006", "K-2306", 12.40),
             ("B-2007", "K-2307", 305.00), ("B-2008", "K-2308", 50.01), ("B-2009", "K-2309", 89.90),
             ("B-2010", "K-2310", 19.95)]
    daten_zellen(ws, 4, daten, [None, None, EURO])
    for r in range(4, 14):
        ziel(ws, f"D{r}", f'=IF(C{r}>50,"Versandkostenfrei","Versandkosten")', None, loesung)
        ziel(ws, f"E{r}", f"=IF(C{r}>50,0,4.95)", EURO, loesung)
    if loesung:
        ws["F4"].value = "50 ist nicht größer als 50, deshalb fallen bei genau 50,00 € Versandkosten an."
        ws["F4"].font = F10
    breiten(ws, {"A": 14, "B": 14, "C": 16, "D": 20, "E": 22, "F": 40})
    fuss(ws, loesung)
    return speichern(wb, "r3-bestellungen", loesung)


# ------------------------------------------------------------------ r4
def r4(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Mengenrabatte berechnen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Geschäftsleitung hat einen gestaffelten Rabatt auf den Bestellwert beschlossen. Die Staffel finden Sie rechts neben der Tabelle.",
        ["Ermitteln Sie in Spalte C für jede Bestellung den Rabattsatz.",
         "Berechnen Sie in Spalte D den Rabatt in Euro und in Spalte E den Rechnungsbetrag.",
         "Prüfen Sie die Grenzfälle (99,99 €, 100,00 €, 499,99 € und 500,00 €) und korrigieren Sie Ihre Formel, falls ein Ergebnis nicht zur Staffel passt.",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Zwei WENN-Funktionen lassen sich ineinander schachteln und liefern so drei Ergebnisse. Überlegen Sie, mit welcher Bedingung Sie beginnen.",
        "Lösung: C4 =WENN(B4>=500;0,1;WENN(B4>=100;0,05;0)); D4 =B4*C4; E4 =B4-D4. "
        "Beispiele: 99,99 € → 0 %, 100,00 € → 5 %, 499,99 € → 5 %, 500,00 € → 10 %.", loesung)
    ws = wb.create_sheet("Rabatt")
    titel_zeile(ws, "BüroComfort GmbH – Rabattberechnung", 5)
    tabelle_kopf(ws, 3, ["Bestell-Nr.", "Bestellwert", "Rabattsatz", "Rabatt in €", "Rechnungsbetrag"])
    werte = [45.00, 99.99, 100.00, 250.00, 499.99, 500.00, 780.50, 999.99, 1000.00, 1450.00]
    daten_zellen(ws, 4, [(f"B-3{str(i + 1).zfill(3)}", w) for i, w in enumerate(werte)], [None, EURO])
    for r in range(4, 14):
        ziel(ws, f"C{r}", f"=IF(B{r}>=500,0.1,IF(B{r}>=100,0.05,0))", PROZ, loesung)
        ziel(ws, f"D{r}", f"=B{r}*C{r}", EURO, loesung)
        ziel(ws, f"E{r}", f"=B{r}-D{r}", EURO, loesung)
    tabelle_kopf(ws, 3, ["Rabattstaffel: Bestellwert ab", "Rabattsatz"], 7)
    daten_zellen(ws, 4, [(0, 0.0), (100, 0.05), (500, 0.10)], [EURO, PROZ], 7)
    breiten(ws, {"A": 14, "B": 16, "C": 14, "D": 16, "E": 18, "F": 4, "G": 28, "H": 14})
    fuss(ws, loesung)
    return speichern(wb, "r4-rabatt", loesung)


KUNDEN5 = [("K-2301", "Ja", 250.00, "DE"), ("K-2302", "Nein", 320.00, "DE"), ("K-2303", "Ja", 150.00, "AT"),
           ("K-2304", "Ja", 200.00, "CH"), ("K-2305", "Nein", 80.00, "AT"), ("K-2306", "Ja", 480.00, "DK"),
           ("K-2307", "Ja", 199.99, "DE"), ("K-2308", "Nein", 1200.00, "DK"), ("K-2309", "Ja", 205.00, "AT"),
           ("K-2310", "Nein", 95.00, "DE")]


# ------------------------------------------------------------------ r5
def r5(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Bedingungen prüfen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Der Vertrieb vergibt einen Bonus nur an Bonuskunden mit einem Bestellwert ab 200,00 €. "
        "Der Standardversand gilt für Bestellungen aus Deutschland oder Österreich.",
        ["Prüfen Sie in Spalte E, ob ein Kunde Bonuskunde ist und zugleich einen Bestellwert ab 200,00 € hat. Das Ergebnis soll WAHR oder FALSCH lauten.",
         "Prüfen Sie in Spalte F, ob die Bestellung aus Deutschland (DE) oder aus Österreich (AT) kommt. Das Ergebnis soll WAHR oder FALSCH lauten.",
         "Prüfen Sie die Grenzfälle (200,00 € und 199,99 €) und erklären Sie in einem Satz in Zelle H4, wie Ihre Bedingung diese Fälle behandelt."],
        "Überlegen Sie, ob alle Bedingungen gleichzeitig erfüllt sein müssen oder ob eine reicht.",
        "Lösung: E4 =UND(B4=\"Ja\";C4>=200); F4 =ODER(D4=\"DE\";D4=\"AT\"). K-2304 (200,00 €) → WAHR, K-2307 (199,99 €) → FALSCH.", loesung)
    ws = wb.create_sheet("Kunden")
    titel_zeile(ws, "BüroComfort GmbH – Kunden- und Bestelldaten", 6)
    tabelle_kopf(ws, 3, ["Kunden-Nr.", "Bonuskunde", "Bestellwert", "Land", "Bonus möglich?", "Standardversand?"])
    daten_zellen(ws, 4, KUNDEN5, [None, None, EURO, None])
    for r in range(4, 14):
        ziel(ws, f"E{r}", f'=AND(B{r}="Ja",C{r}>=200)', None, loesung)
        ziel(ws, f"F{r}", f'=OR(D{r}="DE",D{r}="AT")', None, loesung)
    if loesung:
        ws["H4"].value = "Ab 200,00 € heißt: 200,00 € zählt mit (>=), 199,99 € nicht."
        ws["H4"].font = F10
    breiten(ws, {"A": 14, "B": 14, "C": 16, "D": 8, "E": 18, "F": 20, "G": 4, "H": 50})
    fuss(ws, loesung)
    return speichern(wb, "r5-kunden", loesung)


# ------------------------------------------------------------------ r6
def r6(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Bonus und Versandart festlegen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Bonuskunden mit einem Bestellwert ab 200,00 € erhalten einen Bonus von 5 % des Bestellwerts, alle anderen Kunden erhalten keinen Bonus. "
        "Bestellungen aus Deutschland oder Österreich werden mit „Standardversand“ geliefert, alle anderen mit „Auslandsversand“.",
        ["Berechnen Sie in Spalte E den Bonus in Euro für jede Bestellung.",
         "Zeigen Sie in Spalte F die Versandart („Standardversand“ oder „Auslandsversand“) an.",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Verbinden Sie die Funktionen, die Sie schon kennen, in einer Formel: Die Bedingung der WENN-Funktion kann eine andere Funktion sein.",
        "Lösung: E4 =WENN(UND(B4=\"Ja\";C4>=200);C4*0,05;0); F4 =WENN(ODER(D4=\"DE\";D4=\"AT\");\"Standardversand\";\"Auslandsversand\").", loesung)
    ws = wb.create_sheet("Bonus")
    titel_zeile(ws, "BüroComfort GmbH – Bonus und Versandart", 6)
    tabelle_kopf(ws, 3, ["Kunden-Nr.", "Bonuskunde", "Bestellwert", "Land", "Bonus in €", "Versandart"])
    daten = [("K-3101", "Ja", 260.00, "DE"), ("K-3102", "Nein", 410.00, "AT"), ("K-3103", "Ja", 180.00, "DE"),
             ("K-3104", "Ja", 200.00, "DK"), ("K-3105", "Nein", 90.00, "CH"), ("K-3106", "Ja", 520.00, "AT"),
             ("K-3107", "Ja", 199.00, "DE"), ("K-3108", "Nein", 75.00, "DE"), ("K-3109", "Ja", 305.00, "CH"),
             ("K-3110", "Nein", 640.00, "DK")]
    daten_zellen(ws, 4, daten, [None, None, EURO, None])
    for r in range(4, 14):
        ziel(ws, f"E{r}", f'=IF(AND(B{r}="Ja",C{r}>=200),C{r}*0.05,0)', EURO, loesung)
        ziel(ws, f"F{r}", f'=IF(OR(D{r}="DE",D{r}="AT"),"Standardversand","Auslandsversand")', None, loesung)
    breiten(ws, {"A": 14, "B": 14, "C": 16, "D": 8, "E": 16, "F": 22})
    fuss(ws, loesung)
    return speichern(wb, "r6-bonus", loesung)


# ------------------------------------------------------------------ rm
def rm(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Meilenstein: Bestellungen der KW 12 kalkulieren in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Buchhaltung braucht für die Kalenderwoche 12 eine vollständige Kalkulation aller Bestellungen. "
        "Es gelten: Rabattstaffel auf den Bestellwert: 5 % ab 100,00 €, 10 % ab 500,00 €. "
        "Zusatzbonus: Bonuskunden mit einem Bestellwert ab 300,00 € erhalten 2 % des Bestellwerts als Zusatzbonus. "
        "Versandart: Deutschland oder Österreich = „Standard“, alle anderen Länder = „Ausland“.",
        ["Ermitteln Sie in Spalte F den Rabattsatz und in Spalte G den Rabatt in Euro.",
         "Berechnen Sie in Spalte H den Zusatzbonus in Euro.",
         "Zeigen Sie in Spalte I die Versandart an.",
         "Berechnen Sie in Spalte J den Zahlbetrag (Bestellwert minus Rabatt minus Zusatzbonus).",
         "Ermitteln Sie in der Kennzahlen-Tabelle ab Zeile 21: Summe der Bestellwerte, Durchschnitt der Bestellwerte, größten und kleinsten Bestellwert, Summe der Rabatte und Summe der Zusatzboni.",
         "Formatieren Sie die Tabelle nach dem Corporate Design der BüroComfort GmbH."],
        "Verwenden Sie in Ihren Formeln Zellbezüge und keine eingetippten Ergebnisse. Prüfen Sie die Grenzfälle (z. B. 100,00 € und 500,00 €) selbst.",
        "Lösung: F =WENN(E>=500;0,1;WENN(E>=100;0,05;0)); G =E*F; H =WENN(UND(C=\"Ja\";E>=300);E*0,02;0); "
        "I =WENN(ODER(D=\"DE\";D=\"AT\");\"Standard\";\"Ausland\"); J =E-G-H. Kennzahlen mit SUMME, MITTELWERT, MAX, MIN. "
        "Erwartete Summe der Bestellwerte: 7.581,15 €.", loesung)
    ws = wb.create_sheet("Bestellungen KW 12")
    titel_zeile(ws, "BüroComfort GmbH – Kalkulation der Bestellungen KW 12", 10)
    tabelle_kopf(ws, 3, ["Bestell-Nr.", "Kunden-Nr.", "Bonuskunde", "Land", "Bestellwert", "Rabattsatz",
                         "Rabatt in €", "Zusatzbonus in €", "Versandart", "Zahlbetrag"])
    daten = [("B-5001", "K-4101", "Ja", "DE", 85.00), ("B-5002", "K-4102", "Nein", "AT", 230.50),
             ("B-5003", "K-4103", "Ja", "DE", 640.00), ("B-5004", "K-4104", "Ja", "DK", 1180.00),
             ("B-5005", "K-4105", "Nein", "DE", 99.90), ("B-5006", "K-4106", "Ja", "AT", 310.00),
             ("B-5007", "K-4107", "Nein", "CH", 520.75), ("B-5008", "K-4108", "Ja", "DE", 1000.00),
             ("B-5009", "K-4109", "Nein", "DE", 45.20), ("B-5010", "K-4110", "Ja", "CH", 799.00),
             ("B-5011", "K-4111", "Ja", "DE", 300.00), ("B-5012", "K-4112", "Nein", "AT", 1495.80),
             ("B-5013", "K-4113", "Ja", "DK", 100.00), ("B-5014", "K-4114", "Nein", "DE", 275.00),
             ("B-5015", "K-4115", "Ja", "AT", 500.00)]
    daten_zellen(ws, 4, daten, [None, None, None, None, EURO])
    for r in range(4, 19):
        ziel(ws, f"F{r}", f"=IF(E{r}>=500,0.1,IF(E{r}>=100,0.05,0))", PROZ, loesung)
        ziel(ws, f"G{r}", f"=E{r}*F{r}", EURO, loesung)
        ziel(ws, f"H{r}", f'=IF(AND(C{r}="Ja",E{r}>=300),E{r}*0.02,0)', EURO, loesung)
        ziel(ws, f"I{r}", f'=IF(OR(D{r}="DE",D{r}="AT"),"Standard","Ausland")', None, loesung)
        ziel(ws, f"J{r}", f"=E{r}-G{r}-H{r}", EURO, loesung)
    tabelle_kopf(ws, 21, ["Kennzahl", "", "", "", "Wert"])
    kennzahlen = [("Summe der Bestellwerte", "=SUM(E4:E18)", EURO), ("Durchschnitt der Bestellwerte", "=AVERAGE(E4:E18)", EURO),
                  ("Größter Bestellwert", "=MAX(E4:E18)", EURO), ("Kleinster Bestellwert", "=MIN(E4:E18)", EURO),
                  ("Summe der Rabatte", "=SUM(G4:G18)", EURO), ("Summe der Zusatzboni", "=SUM(H4:H18)", EURO)]
    for i, (text, formel, fmt) in enumerate(kennzahlen):
        r = 22 + i
        beschriftung(ws, f"A{r}", text)
        ziel(ws, f"E{r}", formel, fmt, loesung)
    breiten(ws, {"A": 14, "B": 14, "C": 14, "D": 8, "E": 16, "F": 12, "G": 14, "H": 18, "I": 14, "J": 16})
    fuss(ws, loesung)
    return speichern(wb, "rm-kalkulation", loesung)


if __name__ == "__main__":
    for fn in (r1, r2, r3, r4, r5, r6, rm):
        for l in (False, True):
            print(fn(l))
