#!/usr/bin/env python3
"""Stufe 1 (R1 bis R6, RM) im neuen Aufbau erzeugen.

Jede Übungsdatei besteht aus:
  Start  |  Arbeitsblätter (Auftrag oben auf dem Blatt, Tabelle ab Zeile 14)  |  Corporate Design und Abgabe
Erzeugt Übungsdateien (material/), Musterlösungen (lehrkraft/loesungen/) und schreibt module/*.md neu
(aus werkzeuge/md_orig/*.md; die Aufgabenblöcke werden aus den Blattbeschreibungen erzeugt).

Aufruf aus dem Kursordner:   python3 werkzeuge/gen_stufe1_neu.py [r1 r2 ...]
Benötigt openpyxl und LibreOffice (soffice) zum Neuberechnen der Lösungen.
"""
import os, re, shutil, subprocess, sys, tempfile
AUSWAHL = sys.argv[1:]
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ext_framework import *
import gen_stufe1 as G1
import ext_s1, ext_s4
import ext_run as ER
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

KURS = ER.KURS
ORIG = ER.ORIG

NAVY, MITTE, HELL, BLASS, GELB = "1F3864", "2F5597", "DCE6F1", "F2F2F2", "FFF2CC"
ARIAL = "Arial"
LINKS = Alignment(horizontal="left", vertical="top", wrap_text=True)
RAND = Border(left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"),
              top=Side(style="thin", color="BFBFBF"), bottom=Side(style="thin", color="BFBFBF"))


def f(size=10, bold=False, italic=False, color="000000"):
    return Font(name=ARIAL, size=size, bold=bold, italic=italic, color=color)


def fill(c):
    return PatternFill("solid", fgColor=c)


# ------------------------------------------------------------------ Basisblätter (Aufgabe 1)
KUNDEN5 = G1.KUNDEN5
BASIS = {}

BASIS["r1"] = dict(name="Umsätze KW 12", niveau="basis", min=15, kurz="Wochenumsätze auswerten", nr=1,
 titel="Lernsituation: Tagesumsätze auswerten",
 sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. Die Geschäftsleitung benötigt für die Wochenbesprechung eine Übersicht über die Umsätze der vier Filialen in der Kalenderwoche 12.",
 auftr=["Ermitteln Sie für jeden Tag den Gesamtumsatz aller Filialen (Spalte F).",
        "Ermitteln Sie für jede Filiale den Wochenumsatz (Zeile 20) und den Gesamtumsatz der Woche (Zelle F20).",
        "Ermitteln Sie für jede Filiale und für den Gesamtumsatz den durchschnittlichen Tagesumsatz (Zeile 21).",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design (siehe Blatt „Corporate Design und Abgabe“)."],
 hinweis="Für Summen und Durchschnitte gibt es eigene Funktionen. Prüfen Sie am Ende, ob die Ergebnisse zu den Einzelwerten passen.",
 cols=[("Tag", 16, None, "d", ""), ("Rendsburg", 14, EURO, "d", ""), ("Kiel", 14, EURO, "d", ""), ("Flensburg", 14, EURO, "d", ""),
       ("Neumünster", 14, EURO, "d", ""), ("Gesamt", 16, EURO, "f", "=SUM(B{r}:E{r})")],
 rows=[("Montag", 4120, 5380, 3290, 2760), ("Dienstag", 3980, 5120, 3410, 2890), ("Mittwoch", 4410, 5760, 3620, 3050),
       ("Donnerstag", 4290, 5490, 3550, 2980), ("Freitag", 5230, 6840, 4120, 3470), ("Samstag", 6180, 7920, 5060, 4210)],
 extras=[("label", "A20", "Wochenumsatz", True), ("label", "A21", "Ø Tagesumsatz", True)] +
        [("target", f"{c}20", f"=SUM({c}14:{c}19)", EURO, c == "F", f"Wochenumsatz {n}") for c, n in zip("BCDEF", ["Rendsburg", "Kiel", "Flensburg", "Neumünster", "Gesamt"])] +
        [("target", f"{c}21", f"=AVERAGE({c}14:{c}19)", EURO, c == "F", f"Ø Tagesumsatz {n}") for c, n in zip("BCDEF", ["Rendsburg", "Kiel", "Flensburg", "Neumünster", "Gesamt"])],

 hilfe=[("Wo beginne ich?", "Beginnen Sie mit F14. Dort soll die Summe der Zellen B14 bis E14 stehen. Danach kopieren Sie die Formel mit dem Ausfüllkästchen bis F19."),
        ("Welche Formeln brauche ich?", "Summen: `=SUMME(B14:E14)` für einen Tag, `=SUMME(B14:B19)` für eine Filiale. Durchschnitt: `=MITTELWERT(B14:B19)`.")])

BASIS["r2"] = dict(name="Preisliste", niveau="basis", min=12, kurz="Preisspanne ermitteln", nr=1,
 titel="Lernsituation: Preisliste prüfen",
 sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. Der Einkauf möchte wissen, in welcher Preisspanne sich das Sortiment bewegt. Die aktuelle Preisliste steht in der Tabelle.",
 auftr=["Ermitteln Sie den niedrigsten Preis (Zelle C28) und den höchsten Preis (Zelle C29) der Preisliste.",
        "Tragen Sie in die Zellen B28 und B29 ein, um welche Artikel es sich handelt.",
        "Berechnen Sie die Preisspanne, also den Unterschied zwischen höchstem und niedrigstem Preis (Zelle C30).",
        "Berechnen Sie den Durchschnittspreis aller Artikel (Zelle C31).",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design (siehe Blatt „Corporate Design und Abgabe“)."],
 hinweis="Beantworten Sie die Fragen mit Funktionen und nicht durch Hinsehen. Wenn sich ein Preis ändert, soll sich Ihr Ergebnis automatisch anpassen.",
 cols=[("Art.-Nr.", 14, None, "d", ""), ("Artikel", 32, None, "d", ""), ("Preis", 14, EURO, "d", "")],
 rows=[("A-1001", "Schreibtischstuhl Basic", 129.00), ("A-1002", "Schreibtisch höhenverstellbar", 489.90), ("A-1003", "Aktenschrank 2 Ordnerhöhen", 219.00),
       ("A-1004", "Ordner A4 breit", 2.49), ("A-1005", "Kugelschreiber", 0.89), ("A-1006", "Kopierpapier 500 Blatt", 4.95), ("A-1007", "Locher", 8.90),
       ("A-1008", "Tacker", 6.50), ("A-1009", "Monitorständer", 34.90), ("A-1010", "Schreibtischlampe LED", 42.50), ("A-1011", "Rollcontainer", 119.00),
       ("A-1012", "Whiteboard 90 x 60 cm", 56.00)],
 extras=[("label", "A28", "Niedrigster Preis", True), ("label", "A29", "Höchster Preis", True), ("label", "A30", "Preisspanne", True), ("label", "A31", "Durchschnittspreis", True),
         ("answer", "B28", "Kugelschreiber", "Artikel zum niedrigsten Preis", None), ("answer", "B29", "Schreibtisch höhenverstellbar", "Artikel zum höchsten Preis", None),
         ("target", "C28", "=MIN(C14:C25)", EURO, False, "Niedrigster Preis"), ("target", "C29", "=MAX(C14:C25)", EURO, False, "Höchster Preis"),
         ("target", "C30", "=C29-C28", EURO, False, "Preisspanne"), ("target", "C31", "=AVERAGE(C14:C25)", EURO, False, "Durchschnittspreis")],
 checks=["C28 und C29 enthalten Funktionen und keine eingetippten Preise.", "Die Artikelnamen in B28 und B29 passen zu den Preisen."],
 hilfe=[("Welche Funktionen?", "Für den kleinsten Wert gibt es MIN, für den größten MAX. Den Durchschnitt kennen Sie aus Modul R1.")],
 widths={"A": 20})

BASIS["r3"] = dict(name="Bestellungen", niveau="basis", min=20, kurz="Versandkosten anzeigen", nr=1,
 titel="Lernsituation: Versandkosten prüfen",
 sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. Der Vertrieb möchte in der Bestellliste sofort sehen, welche Bestellungen versandkostenfrei sind. Versandkostenfrei sind Bestellungen mit einem Bestellwert über 50,00 €. Alle anderen Bestellungen werden mit Versandkosten berechnet.",
 auftr=["Zeigen Sie in Spalte D für jede Bestellung den Text „Versandkostenfrei“ oder „Versandkosten“ an.",
        "Prüfen Sie die Bestellung B-2004 mit genau 50,00 € besonders genau. Begründen Sie in Zelle B24 in einem Satz, warum Ihr Ergebnis stimmt.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design (siehe Blatt „Corporate Design und Abgabe“)."],
 hinweis="Die Anzeige soll sich automatisch ändern, wenn sich ein Bestellwert ändert.",
 cols=[("Bestell-Nr.", 14, None, "d", ""), ("Kunden-Nr.", 14, None, "d", ""), ("Bestellwert", 16, EURO, "d", ""),
       ("Versand", 20, None, "f", '=IF(C{r}>50,"Versandkostenfrei","Versandkosten")'),
       ("Versandkosten (Plus)", 22, EURO, "f", "=IF(C{r}>50,0,4.95)")],
 rows=[("B-2001", "K-2301", 24.90), ("B-2002", "K-2302", 67.80), ("B-2003", "K-2303", 49.99), ("B-2004", "K-2304", 50.00), ("B-2005", "K-2305", 128.50),
       ("B-2006", "K-2306", 12.40), ("B-2007", "K-2307", 305.00), ("B-2008", "K-2308", 50.01), ("B-2009", "K-2309", 89.90), ("B-2010", "K-2310", 19.95)],
 extras=[("label", "A24", "Begründung B-2004", True),
         ("answer", "B24", "50 ist nicht größer als 50, deshalb fallen bei genau 50,00 € Versandkosten an.", None, "B24:E24")],
 checks=["Die Bestellungen B-2002, B-2005, B-2007, B-2008 und B-2009 sind versandkostenfrei.", "B-2004 (50,00 €) ist **nicht** versandkostenfrei."],
 hilfe=[("Welche Prüfung?", "Überlegen Sie: Wann soll „Versandkostenfrei“ erscheinen? Wenn der Bestellwert in Spalte C größer als 50 ist."),
        ("Formelaufbau", "`=WENN(C14>50; …; …)`. Setzen Sie für die Punkte die beiden Texte in Anführungszeichen ein.")])

BASIS["r4"] = dict(name="Rabatt", niveau="basis", min=20, kurz="Rabatt berechnen", nr=1,
 titel="Lernsituation: Mengenrabatte berechnen",
 sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. Die Geschäftsleitung hat einen gestaffelten Rabatt auf den Bestellwert beschlossen. Die Staffel finden Sie rechts neben der Tabelle.",
 auftr=["Ermitteln Sie in Spalte C für jede Bestellung den Rabattsatz.",
        "Berechnen Sie in Spalte D den Rabatt in Euro und in Spalte E den Rechnungsbetrag.",
        "Prüfen Sie die Grenzfälle (99,99 €, 100,00 €, 499,99 € und 500,00 €) und korrigieren Sie Ihre Formel, falls ein Ergebnis nicht zur Staffel passt.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design (siehe Blatt „Corporate Design und Abgabe“)."],
 hinweis="Zwei WENN-Funktionen lassen sich ineinander schachteln und liefern so drei Ergebnisse. Überlegen Sie, mit welcher Bedingung Sie beginnen.",
 cols=[("Bestell-Nr.", 14, None, "d", ""), ("Bestellwert", 16, EURO, "d", ""), ("Rabattsatz", 14, PROZ, "f", "=IF(B{r}>=500,0.1,IF(B{r}>=100,0.05,0))"),
       ("Rabatt in €", 16, EURO, "f", "=B{r}*C{r}"), ("Rechnungsbetrag", 18, EURO, "f", "=B{r}-D{r}")],
 rows=[(f"B-3{str(i + 1).zfill(3)}", w) for i, w in enumerate([45.00, 99.99, 100.00, 250.00, 499.99, 500.00, 780.50, 999.99, 1000.00, 1450.00])],
 extras=[("table", 13, 7, ["Rabattstaffel: Bestellwert ab", "Rabattsatz"], [(0, 0.0), (100, 0.05), (500, 0.10)], [EURO, PROZ])],
 widths={"G": 28, "H": 14},
 checks=["99,99 € erhält 0 %, 100,00 € erhält 5 %.", "499,99 € erhält 5 %, 500,00 € erhält 10 %."],
 hilfe=[("Reihenfolge", "Beginnen Sie mit der höchsten Grenze (500). Dann folgt 100. Ganz am Ende steht 0."),
        ("Formelaufbau", "`=WENN(B14>=500; … ;WENN(B14>=100; … ; …))`. Setzen Sie für die Punkte die Rabattsätze ein.")])

BASIS["r5"] = dict(name="Kunden", niveau="basis", min=12, kurz="Bedingungen prüfen", nr=1,
 titel="Lernsituation: Bedingungen prüfen",
 sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. Der Vertrieb vergibt einen Bonus nur an Bonuskunden mit einem Bestellwert ab 200,00 €. Der Standardversand gilt für Bestellungen aus Deutschland oder Österreich.",
 auftr=["Prüfen Sie in Spalte E, ob ein Kunde Bonuskunde ist und zugleich einen Bestellwert ab 200,00 € hat. Das Ergebnis soll WAHR oder FALSCH lauten.",
        "Prüfen Sie in Spalte F, ob die Bestellung aus Deutschland (DE) oder aus Österreich (AT) kommt. Das Ergebnis soll WAHR oder FALSCH lauten.",
        "Prüfen Sie die Grenzfälle (200,00 € und 199,99 €) und erklären Sie in einem Satz in Zelle B24, wie Ihre Bedingung diese Fälle behandelt."],
 hinweis="Überlegen Sie, ob alle Bedingungen gleichzeitig erfüllt sein müssen oder ob eine reicht.",
 cols=[("Kunden-Nr.", 14, None, "d", ""), ("Bonuskunde", 14, None, "d", ""), ("Bestellwert", 16, EURO, "d", ""), ("Land", 8, None, "d", ""),
       ("Bonus möglich?", 18, None, "f", '=AND(B{r}="Ja",C{r}>=200)'), ("Standardversand?", 20, None, "f", '=OR(D{r}="DE",D{r}="AT")')],
 rows=KUNDEN5,
 extras=[("label", "A24", "Erklärung Grenzfälle", True),
         ("answer", "B24", "Ab 200,00 € heißt: 200,00 € zählt mit (>=), 199,99 € nicht.", None, "B24:F24")],
 checks=["K-2304 (Bonuskunde, 200,00 €) ist in Spalte E WAHR.", "K-2307 (Bonuskunde, 199,99 €) ist in Spalte E FALSCH."],
 hilfe=[("Was muss gelten?", "In Spalte E müssen zwei Dinge gleichzeitig gelten. In Spalte F reicht eine von zwei Möglichkeiten.")])

R6_ROWS = [("K-3101", "Ja", 260.00, "DE"), ("K-3102", "Nein", 410.00, "AT"), ("K-3103", "Ja", 180.00, "DE"), ("K-3104", "Ja", 200.00, "DK"),
           ("K-3105", "Nein", 90.00, "CH"), ("K-3106", "Ja", 520.00, "AT"), ("K-3107", "Ja", 199.00, "DE"), ("K-3108", "Nein", 75.00, "DE"),
           ("K-3109", "Ja", 305.00, "CH"), ("K-3110", "Nein", 640.00, "DK")]
BASIS["r6"] = dict(name="Bonus", niveau="basis", min=20, kurz="Bonus und Versandart festlegen", nr=1,
 titel="Lernsituation: Bonus und Versandart festlegen",
 sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. Bonuskunden mit einem Bestellwert ab 200,00 € erhalten einen Bonus von 5 % des Bestellwerts, alle anderen Kunden erhalten keinen Bonus. Bestellungen aus Deutschland oder Österreich werden mit „Standardversand“ geliefert, alle anderen mit „Auslandsversand“.",
 auftr=["Berechnen Sie in Spalte E den Bonus in Euro für jede Bestellung.",
        "Zeigen Sie in Spalte F die Versandart („Standardversand“ oder „Auslandsversand“) an.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design (siehe Blatt „Corporate Design und Abgabe“)."],
 hinweis="Verbinden Sie die Funktionen, die Sie schon kennen, in einer Formel: Die Bedingung der WENN-Funktion kann eine andere Funktion sein.",
 cols=[("Kunden-Nr.", 14, None, "d", ""), ("Bonuskunde", 14, None, "d", ""), ("Bestellwert", 16, EURO, "d", ""), ("Land", 8, None, "d", ""),
       ("Bonus in €", 16, EURO, "f", '=IF(AND(B{r}="Ja",C{r}>=200),C{r}*0.05,0)'),
       ("Versandart", 22, None, "f", '=IF(OR(D{r}="DE",D{r}="AT"),"Standardversand","Auslandsversand")')],
 rows=R6_ROWS,
 checks=["K-3104 (Bonuskunde, 200,00 €) erhält 10,00 € Bonus.", "K-3107 (Bonuskunde, 199,00 €) erhält keinen Bonus."],
 hilfe=[("Regel als Satz", "Spalte E: Wenn Bonuskunde **und** Bestellwert ≥ 200, dann 5 % vom Bestellwert, sonst 0. Spalte F: Wenn Land DE **oder** AT, dann „Standardversand“, sonst „Auslandsversand“.")])

RM_ROWS = [("B-5001", "K-4101", "Ja", "DE", 85.00), ("B-5002", "K-4102", "Nein", "AT", 230.50), ("B-5003", "K-4103", "Ja", "DE", 640.00),
           ("B-5004", "K-4104", "Ja", "DK", 1180.00), ("B-5005", "K-4105", "Nein", "DE", 99.90), ("B-5006", "K-4106", "Ja", "AT", 310.00),
           ("B-5007", "K-4107", "Nein", "CH", 520.75), ("B-5008", "K-4108", "Ja", "DE", 1000.00), ("B-5009", "K-4109", "Nein", "DE", 45.20),
           ("B-5010", "K-4110", "Ja", "CH", 799.00), ("B-5011", "K-4111", "Ja", "DE", 300.00), ("B-5012", "K-4112", "Nein", "AT", 1495.80),
           ("B-5013", "K-4113", "Ja", "DK", 100.00), ("B-5014", "K-4114", "Nein", "DE", 275.00), ("B-5015", "K-4115", "Ja", "AT", 500.00)]
KENN = [("Summe der Bestellwerte", "=SUM(E14:E28)"), ("Durchschnitt der Bestellwerte", "=AVERAGE(E14:E28)"), ("Größter Bestellwert", "=MAX(E14:E28)"),
        ("Kleinster Bestellwert", "=MIN(E14:E28)"), ("Summe der Rabatte", "=SUM(G14:G28)"), ("Summe der Zusatzboni", "=SUM(H14:H28)")]
BASIS["rm"] = dict(name="Bestellungen KW 12", niveau="basis", min=60, kurz="Bestellungen der KW 12 kalkulieren", nr=1,
 titel="Meilenstein: Bestellungen der KW 12 kalkulieren",
 sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. Die Buchhaltung braucht für die Kalenderwoche 12 eine vollständige Kalkulation aller Bestellungen. Es gelten: Rabattstaffel auf den Bestellwert: 5 % ab 100,00 €, 10 % ab 500,00 €. Zusatzbonus: Bonuskunden mit einem Bestellwert ab 300,00 € erhalten 2 % des Bestellwerts als Zusatzbonus. Versandart: Deutschland oder Österreich = „Standard“, alle anderen Länder = „Ausland“.",
 auftr=["Ermitteln Sie in Spalte F den Rabattsatz und in Spalte G den Rabatt in Euro.", "Berechnen Sie in Spalte H den Zusatzbonus in Euro.", "Zeigen Sie in Spalte I die Versandart an.",
        "Berechnen Sie in Spalte J den Zahlbetrag (Bestellwert minus Rabatt minus Zusatzbonus).",
        "Ermitteln Sie in der Kennzahlen-Tabelle ab Zeile 31: Summe der Bestellwerte, Durchschnitt, größten und kleinsten Bestellwert, Summe der Rabatte und Summe der Zusatzboni.",
        "Formatieren Sie die Tabelle nach dem Corporate Design (siehe Blatt „Corporate Design und Abgabe“)."],
 hinweis="Verwenden Sie in Ihren Formeln Zellbezüge und keine eingetippten Ergebnisse. Prüfen Sie die Grenzfälle (z. B. 100,00 € und 500,00 €) selbst.",
 cols=[("Bestell-Nr.", 12, None, "d", ""), ("Kunden-Nr.", 12, None, "d", ""), ("Bonuskunde", 12, None, "d", ""), ("Land", 8, None, "d", ""), ("Bestellwert", 14, EURO, "d", ""),
       ("Rabattsatz", 12, PROZ, "f", "=IF(E{r}>=500,0.1,IF(E{r}>=100,0.05,0))"), ("Rabatt in €", 14, EURO, "f", "=E{r}*F{r}"),
       ("Zusatzbonus in €", 18, EURO, "f", '=IF(AND(C{r}="Ja",E{r}>=300),E{r}*0.02,0)'), ("Versandart", 14, None, "f", '=IF(OR(D{r}="DE",D{r}="AT"),"Standard","Ausland")'),
       ("Zahlbetrag", 16, EURO, "f", "=E{r}-G{r}-H{r}")],
 rows=RM_ROWS,
 extras=[("label", "A31", "Kennzahl", True), ("label", "E31", "Wert", True)] + [("label", f"A{32 + i}", n) for i, (n, _) in enumerate(KENN)] +
        [("target", f"E{32 + i}", fo, EURO, False, n) for i, (n, fo) in enumerate(KENN)],
 widths={"A": 14})

THEMA = {
 "r1": "Sie addieren Werte mit SUMME und berechnen Durchschnitte mit MITTELWERT.",
 "r2": "Sie finden mit MIN und MAX den kleinsten und größten Wert einer Liste.",
 "r3": "Sie lassen Excel mit der WENN-Funktion eine Entscheidung treffen.",
 "r4": "Sie bilden mit zwei verschachtelten WENN-Funktionen drei Stufen ab, zum Beispiel einen Rabatt.",
 "r5": "Sie verknüpfen Bedingungen mit UND und ODER.",
 "r6": "Sie setzen UND und ODER als Prüfung in die WENN-Funktion ein.",
 "rm": "Meilenstein: Sie kalkulieren eine Bestellliste mit Rabatt, Zusatzbonus, Versandart und Kennzahlen.",
}
TITEL = {"r1": "SUMME und MITTELWERT", "r2": "MIN und MAX", "r3": "WENN", "r4": "WENN verschachtelt", "r5": "UND und ODER",
         "r6": "UND und ODER mit WENN", "rm": "Meilenstein Bestellungen KW 12"}
STEM = ER.STEM

# ------------------------------------------------------------------ Start, Corporate Design und Abgabe
def start_blatt(wb, key, specs, loesung):
    ws = wb.create_sheet("Start", 0)
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 2, "B": 36, "C": 12, "D": 11, "E": 62}.items():
        ws.column_dimensions[col].width = w

    def zeile(r, text, font, fuell=None, hoehe=None, bis="E", von="B"):
        ws.merge_cells(f"{von}{r}:{bis}{r}")
        c = ws[f"{von}{r}"]
        c.value = text
        c.font = font
        c.alignment = LINKS if fuell is None or r != 1 else Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        if fuell:
            for col in "BCDE":
                ws[f"{col}{r}"].fill = fill(fuell)
        if hoehe:
            ws.row_dimensions[r].height = hoehe
        elif not isinstance(text, (int, float)):
            ws.row_dimensions[r].height = zeilenhoehe(text, 112)

    zeile(1, f"Übungsdatei: {TITEL[key]}", f(18, True, color="FFFFFF"), NAVY, 40)
    ws["B1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    zeile(2, "BüroComfort GmbH  ·  Kaufmann/Kauffrau für Büromanagement  ·  Lernpfad Excel-Funktionen", f(10, italic=True, color="FFFFFF"), MITTE, 20)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    zeile(4, "Worum geht es?", f(11, True, color=NAVY), HELL, 20)
    zeile(5, THEMA[key] + " Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg.", f(10))
    zeile(7, "So arbeiten Sie", f(11, True, color=NAVY), HELL, 20)
    schritte = ["Öffnen Sie nacheinander die Tabellenblätter unten in der Liste. Auf jedem Blatt steht oben der Auftrag mit der Situation.",
                "Tragen Sie Ihre Formeln in die leeren Ergebniszellen ein. Verwenden Sie Zellbezüge und keine eingetippten Ergebnisse.",
                "Prüfen Sie Ihre Ergebnisse mit den Selbstchecks im Modul und formatieren Sie nach dem Corporate Design.",
                "Zum Schluss lesen Sie das Blatt „Corporate Design und Abgabe“."]
    for i, s in enumerate(schritte):
        zeile(8 + i, f"{i + 1}.  {s}", f(10))
    zeile(13, "Ihre Tabellenblätter", f(11, True, color=NAVY), HELL, 20)
    for col, h in zip("BCDE", ["Blatt (anklicken)", "Niveau", "Zeit", "Thema"]):
        c = ws[f"{col}14"]
        c.value = h; c.font = f(10, True); c.fill = fill(BLASS); c.border = RAND
    r = 15
    for s in specs:
        lab = {"basis": "Basis", "plus": "Plus"}.get(s["niveau"], s["niveau"])
        vals = [s["name"], lab, f"{s['min']} Min.", s["kurz"]]
        for col, v in zip("BCDE", vals):
            c = ws[f"{col}{r}"]; c.value = v; c.font = f(10); c.border = RAND; c.alignment = LINKS
        ws[f"B{r}"].hyperlink = Hyperlink(ref=f"B{r}", location=f"'{s['name']}'!A1", display=s["name"])
        ws[f"B{r}"].font = Font(name=ARIAL, size=10, underline="single", color="0563C1")
        r += 1
    ws[f"B{r}"].value = "Corporate Design und Abgabe"; ws[f"B{r}"].hyperlink = Hyperlink(ref=f"B{r}", location="'Corporate Design und Abgabe'!A1", display="Corporate Design und Abgabe")
    ws[f"B{r}"].font = Font(name=ARIAL, size=10, underline="single", color="0563C1"); ws[f"B{r}"].border = RAND
    for col, v in zip("CDE", ["Pflicht", "5 Min.", "Gestaltung und Abgabe der Datei"]):
        c = ws[f"{col}{r}"]; c.value = v; c.font = f(10); c.border = RAND
    r += 2
    if loesung:
        zeile(r, "NUR FÜR LEHRKRÄFTE: Musterlösung mit Formeln und Kontrollwerten. Nicht an Lernende ausgeben.", f(10, True, color="C00000"), GELB, 22)
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def design_abgabe_blatt(wb, loesung):
    ws = wb.create_sheet("Corporate Design und Abgabe")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 2, "B": 30, "C": 58, "D": 22}.items():
        ws.column_dimensions[col].width = w

    def kopf(r, text, farbe=HELL, schrift=None):
        ws.merge_cells(f"B{r}:D{r}")
        c = ws[f"B{r}"]; c.value = text; c.font = schrift or f(11, True, color=NAVY); c.alignment = Alignment(vertical="center", indent=1)
        for col in "BCD": ws[f"{col}{r}"].fill = fill(farbe)
        ws.row_dimensions[r].height = 22

    def text(r, t, bold=False):
        ws.merge_cells(f"B{r}:D{r}")
        c = ws[f"B{r}"]; c.value = t; c.font = f(10, bold); c.alignment = LINKS
        ws.row_dimensions[r].height = zeilenhoehe(t, 108)

    ws.merge_cells("B1:D1")
    ws["B1"].value = "Corporate Design und Abgabe"; ws["B1"].font = f(18, True, color="FFFFFF"); ws["B1"].alignment = Alignment(vertical="center", indent=1)
    for col in "BCD": ws[f"{col}1"].fill = fill(NAVY)
    ws.row_dimensions[1].height = 40
    kopf(3, "Corporate Design der BüroComfort GmbH")
    text(4, "Die Zellen, die Sie ausfüllen, liegen im Format „Standard“ vor. Formatieren Sie Ihre Ergebnisse selbst nach diesen Vorgaben:")
    for col, h in zip("BCD", ["Element", "Vorgabe", "Beispiel"]):
        c = ws[f"{col}5"]; c.value = h; c.font = f(10, True); c.fill = fill(BLASS); c.border = RAND
    regeln = [("Schrift", "Schriftart Arial, Schriftgröße 10", "Arial 10", None),
              ("Überschriften", "zentriert, Schriftgröße 12, fett", "Überschrift", "ueb"),
              ("Prozentsätze", "mit zwei Nachkommastellen", 0.1234, PROZ),
              ("Andere Zahlen", "ohne Nachkommastellen, mit Tausenderpunkt", 12345, "#,##0"),
              ("Eurobeträge", "mit Tausenderpunkt und zwei Nachkommastellen", 1234.5, EURO),
              ("Datum", "im Format TT.MM.JJJJ", 46150, "DD.MM.YYYY"),
              ("Uhrzeit", "im Format hh:mm", 0.3541666667, "hh:mm")]
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
                           "Öffnen Sie mit ++Strg+1++ das Fenster „Zellen formatieren“ oder klicken Sie auf Start › Zahl.",
                           "Wählen Sie das Zahlenformat (Prozent, Zahl, Währung, Datum, Uhrzeit) und die Nachkommastellen.",
                           "Schrift und Ausrichtung stellen Sie auf der Registerkarte Start ein."], 1):
        text(r, f"{i}.  {s.replace('++', '')}"); r += 1
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
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


# ------------------------------------------------------------------ Aufbau
def recalc(pfad):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, pfad],
                   check=True, capture_output=True, timeout=240)
    shutil.copy(os.path.join(tmp, os.path.basename(pfad)), pfad)


def extra_specs(key):
    if key == "rm":
        return ext_s4.SPECS["rm"]
    return ext_s1.SPECS.get(key, [])


def bauen(key):
    stem = STEM[key]
    basis = BASIS[key]
    extras = extra_specs(key)
    specs = [basis] + extras
    for loes in (False, True):
        wb = Workbook()
        wb.remove(wb.active)
        for s in specs:
            build_sheet(wb, s, loes, len(wb.sheetnames))
        design_abgabe_blatt(wb, loes)
        start_blatt(wb, key, specs, loes)
        pfad = os.path.join(KURS, "lehrkraft/loesungen" if loes else "material", f"{stem}{'-loesung' if loes else ''}.xlsx")
        wb.save(pfad)
        if loes:
            recalc(pfad)
    vals = load_workbook(os.path.join(KURS, "lehrkraft/loesungen", f"{stem}-loesung.xlsx"), data_only=True)
    bad = [(ws.title, c.coordinate, c.value) for ws in vals for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("#")]
    print(key, "Fehlerzellen:", bad)
    md_schreiben(key, specs, vals)


def md_spec(s):
    """Spalten, die nur als Plus-Aufgabe gedacht sind (Titel endet auf „(Plus)“), nicht in den Selbstcheck der Basisaufgabe."""
    return dict(s, cols=[c for c in s["cols"] if not c[0].endswith("(Plus)")])


def md_schreiben(key, specs, vals):
    t = open(f"{ORIG}/{key}.md", encoding="utf-8").read()
    t = t.replace("Das Blatt „Aufgabenstellung“ enthält den Auftrag.", "Das Blatt „Start“ gibt einen Überblick. Der Auftrag steht oben auf jedem Tabellenblatt.")
    basis = specs[0]
    if key == "rm":
        t = t.replace("Das Blatt „Aufgabenstellung“ enthält alle Regeln und Aufträge.", "Der Auftrag mit allen Regeln steht oben im Blatt. Das Blatt „Start“ gibt einen Überblick.")
        t = t.replace("Lesen Sie zuerst das Blatt „Aufgabenstellung“. Es nennt die Regeln:", "Lesen Sie zuerst den Auftrag oben im Blatt „Bestellungen KW 12“. Er nennt die Regeln:")
        t = t.replace("Kennzahlen-Tabelle ab Zeile 21", "Kennzahlen-Tabelle ab Zeile 31")
        open(f"{KURS}/module/{key}.md", "w", encoding="utf-8").write(t)
        names = f"Die Erweiterung finden Sie im Blatt „{specs[1]['name']}“ derselben Datei."
        block = md_block(specs[1], specs[1].get("nr", 2), vals[specs[1]["name"]], None)
        ER.patch_milestone(key, block, names)
        # patch_milestone liest wieder md_orig: Ersetzungen erneut anwenden
        p = f"{KURS}/module/{key}.md"
        u = open(p, encoding="utf-8").read()
        for a, b in [("Das Blatt „Aufgabenstellung“ enthält alle Regeln und Aufträge.", "Der Auftrag mit allen Regeln steht oben im Blatt. Das Blatt „Start“ gibt einen Überblick."),
                     ("Lesen Sie zuerst das Blatt „Aufgabenstellung“. Es nennt die Regeln:", "Lesen Sie zuerst den Auftrag oben im Blatt „Bestellungen KW 12“. Er nennt die Regeln:"),
                     ("Kennzahlen-Tabelle ab Zeile 21", "Kennzahlen-Tabelle ab Zeile 31")]:
            u = u.replace(a, b)
        open(p, "w", encoding="utf-8").write(u)
        return
    # erste Aufgabe (handgeschrieben) durch erzeugten Block ersetzen
    lines = t.split("\n")
    a, b = block_ranges(lines)[0]
    neu = md_block(md_spec(basis), 1, vals[basis["name"]], None).rstrip("\n").split("\n")
    lines[a:b + 1] = neu
    t = "\n".join(lines)
    open(f"{KURS}/module/{key}.md", "w", encoding="utf-8").write(t)
    # Zusatzblätter einfügen (wie ext_run)
    blocks = [md_block(s, s.get("nr", i), vals[s["name"]], None) for i, s in enumerate(specs[1:], 2)]
    extra_min = sum(s["min"] for s in specs[1:] if s["niveau"] == "basis")
    names = "Zusätzlich enthält die Datei die Blätter " + ", ".join(f"„{s['name']}“" for s in specs[1:]) + ". Die Schwierigkeit steigt von Blatt zu Blatt."
    hat_nr = any("nr" in s for s in specs[1:])
    nr_kreativ = max(s["nr"] for s in specs[1:]) + 1 if hat_nr else len(specs[1:]) + 2
    ER.patch_md(f"{KURS}/module/{key}.md", "\n".join(blocks), extra_min, nr_kreativ, names, rename=not hat_nr)


if __name__ == "__main__":
    for k in (AUSWAHL or list(BASIS)):
        bauen(k)
