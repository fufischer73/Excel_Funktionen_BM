#!/usr/bin/env python3
"""Übungsdateien Stufe 4 (p1, p2, p3, pm). Aufruf: python3 gen_stufe4.py <out_aufgabe> <out_loesung>"""
import random
from datetime import date, time
from gen_stufe1 import *
from gen_stufe3 import STAMM, KUNDEN, ZIELE, stamm_block, kunden_block, ziele_block, staffel_block, DATUM

BASIS = "Verwenden Sie in Ihren Formeln Zellbezüge und keine eingetippten Ergebnisse."
MSTAFFEL = [(1, 0.0), (10, 0.05), (50, 0.10), (100, 0.15)]
RSTAFFEL = [(0, 0.0), (100, 0.05), (500, 0.10), (1000, 0.15)]


# ------------------------------------------------------------------ p1
P1_POS = [("A-1005", 50), ("A-1004", 12), ("A-1006", 6), ("A-1009", 20), ("A-1007", 2), ("A-1012", 4)]


def p1(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Prüfungsnahe Aufgabe: Angebot für einen Kunden erstellen",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Der Kunde mit der Kunden-Nr. K-4103 hat eine Anfrage geschickt. Sie sollen ein Angebot erstellen. "
        "Gegeben sind die Kunden-Nr., das Angebotsdatum, die Artikelnummern und die Mengen. Alles andere entnehmen Sie dem Blatt „Stammdaten“. "
        "Mengenrabatt je Position nach der Mengenstaffel im Blatt „Stammdaten“. Versandkosten: Bei einem Nettobetrag unter dem Betrag in Zelle B8 fallen die Versandkosten aus Zelle B9 an, sonst nicht. "
        "Die Umsatzsteuer beträgt 19 %. Beträge sind kaufmännisch auf zwei Nachkommastellen zu runden. Das Angebot ist 14 Tage gültig.",
        ["Ermitteln Sie Kundenname und Ort des Kunden (B4, B5) und das Datum, bis zu dem das Angebot gültig ist (B7). (6 Punkte)",
         "Ermitteln Sie für jede Position Bezeichnung, Einzelpreis und Mengenrabattsatz (Spalten C, E, F). (9 Punkte)",
         "Berechnen Sie den Gesamtpreis jeder Position nach Abzug des Mengenrabatts, gerundet auf zwei Nachkommastellen (Spalte G). (6 Punkte)",
         "Berechnen Sie die Summe netto (G19), die Versandkosten (G20), den Nettobetrag (G21), die Umsatzsteuer (G22) und den Gesamtbetrag (G23). (12 Punkte)",
         "Formatieren Sie das Angebot nach dem Corporate Design der BüroComfort GmbH und ergänzen Sie die Fußzeile. (7 Punkte)"],
        "Insgesamt 40 Punkte. Lesen Sie zuerst alle Aufträge und die Regeln. Prüfen Sie am Ende Ihr Ergebnis mit dem Taschenrechner. " + BASIS,
        "Lösung: B4 =SVERWEIS(B3;Stammdaten!$F$4:$I$11;2;FALSCH); B5 Index 3; B7 =B6+14; C12 =SVERWEIS(B12;Stammdaten!$A$4:$D$15;2;FALSCH); E12 Index 3; "
        "F12 =SVERWEIS(D12;Stammdaten!$K$4:$L$7;2;WAHR); G12 =RUNDEN(D12*E12*(1-F12);2); G19 =SUMME(G12:G17); G20 =WENN(G19>=B8;0;B9); G21 =G19+G20; "
        "G22 =RUNDEN(G21*F22;2); G23 =G21+G22. Erwartet: Summe netto 154,59 €, Versandkosten 7,90 €, Nettobetrag 162,49 €, USt 30,87 €, Gesamtbetrag 193,36 €. "
        "Verteilung der Punkte: Kundendaten 6, Positionen 9, Gesamtpreise 6, Summen 12, Gestaltung 7.", loesung)
    ws = wb.create_sheet("Angebot")
    titel_zeile(ws, "BüroComfort GmbH – Angebot", 7)
    for r, t in ((3, "Kunden-Nr."), (4, "Kunde"), (5, "Ort"), (6, "Angebotsdatum"), (7, "Gültig bis"), (8, "Versandkostenfrei ab"), (9, "Versandkosten")):
        beschriftung(ws, f"A{r}", t, True)
    for adr, v, f in (("B3", "K-4103", None), ("B6", date(2026, 5, 4), DATUM), ("B8", 200.0, EURO), ("B9", 7.90, EURO)):
        c = ws[adr]; c.value = v; c.font = F10; c.border = RAHMEN
        if f: c.number_format = f
    ziel(ws, "B4", "=VLOOKUP(B3,Stammdaten!$F$4:$I$11,2,FALSE)", None, loesung)
    ziel(ws, "B5", "=VLOOKUP(B3,Stammdaten!$F$4:$I$11,3,FALSE)", None, loesung)
    ziel(ws, "B7", "=B6+14", DATUM, loesung)
    tabelle_kopf(ws, 11, ["Pos.", "Art.-Nr.", "Bezeichnung", "Menge", "Einzelpreis netto", "Mengenrabatt", "Gesamtpreis netto"])
    for i, (a, m) in enumerate(P1_POS):
        r = 12 + i
        for col, v in ((1, i + 1), (2, a), (4, m)):
            c = ws.cell(row=r, column=col, value=v); c.font = F10; c.border = RAHMEN
        ziel(ws, f"C{r}", f"=VLOOKUP(B{r},Stammdaten!$A$4:$D$15,2,FALSE)", None, loesung)
        ziel(ws, f"E{r}", f"=VLOOKUP(B{r},Stammdaten!$A$4:$D$15,3,FALSE)", EURO, loesung)
        ziel(ws, f"F{r}", f"=VLOOKUP(D{r},Stammdaten!$K$4:$L$7,2,TRUE)", PROZ, loesung)
        ziel(ws, f"G{r}", f"=ROUND(D{r}*E{r}*(1-F{r}),2)", EURO, loesung)
    for r, t in ((19, "Summe netto"), (20, "Versandkosten"), (21, "Nettobetrag"), (22, "Umsatzsteuer"), (23, "Gesamtbetrag (brutto)")):
        beschriftung(ws, f"E{r}", t, r == 23)
    c = ws["F22"]; c.value = 0.19; c.font = F10; c.border = RAHMEN; c.number_format = PROZ
    ziel(ws, "G19", "=SUM(G12:G17)", EURO, loesung)
    ziel(ws, "G20", "=IF(G19>=B8,0,B9)", EURO, loesung)
    ziel(ws, "G21", "=G19+G20", EURO, loesung)
    ziel(ws, "G22", "=ROUND(G21*F22,2)", EURO, loesung)
    ziel(ws, "G23", "=G21+G22", EURO, loesung, True)
    breiten(ws, {"A": 22, "B": 14, "C": 30, "D": 10, "E": 24, "F": 14, "G": 18})
    fuss(ws, loesung)
    w = wb.create_sheet("Stammdaten")
    titel_zeile(w, "BüroComfort GmbH – Stammdaten", 12)
    stamm_block(w, 3, 1); kunden_block(w, 3, 6)
    tabelle_kopf(w, 3, ["Menge ab", "Mengenrabatt"], 11)
    daten_zellen(w, 4, MSTAFFEL, [ZAHL, PROZ], 11)
    breiten(w, {"A": 10, "B": 30, "C": 14, "D": 9, "E": 3, "F": 11, "G": 28, "H": 13, "I": 14, "J": 3, "K": 12, "L": 14})
    return speichern(wb, "p1-angebot", loesung)


# ------------------------------------------------------------------ p2
def _p2_daten():
    rng = random.Random(2026)
    fil = ["Rendsburg", "Kiel", "Flensburg", "Neumünster"]
    wg = ["Büromöbel", "Papier", "Schreibwaren", "Technik"]
    gew = {"Rendsburg": 1.0, "Kiel": 1.4, "Flensburg": 0.9, "Neumünster": 0.7}
    zeilen = []
    for i in range(32):
        f = rng.choice(fil)
        w = rng.choice(wg)
        basis = {"Büromöbel": 450, "Papier": 90, "Schreibwaren": 45, "Technik": 300}[w]
        u = round(basis * gew[f] * rng.uniform(0.5, 1.6), 2)
        zeilen.append((f"V-{9001 + i}", f, w, u))
    return zeilen


P2_DATEN = _p2_daten()
P2_FIL = ["Rendsburg", "Kiel", "Flensburg", "Neumünster"]
P2_WG = ["Büromöbel", "Papier", "Schreibwaren", "Technik"]


def p2(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Prüfungsnahe Aufgabe: Umsatzauswertung der Filialen",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Geschäftsleitung benötigt für den Monatsbericht eine Auswertung der 32 Verkäufe aus dem Blatt „Verkäufe“. "
        "Die Auswertungstabellen stehen rechts neben der Liste. Das Umsatzziel je Filiale beträgt 2.200,00 € (Zelle H20).",
        ["Erstellen Sie die Auswertung je Filiale (G4 bis M7): Anzahl der Verkäufe, Umsatz, durchschnittlicher Umsatz je Verkauf, Anteil am Gesamtumsatz, Rang nach Umsatz (höchster Umsatz = Rang 1) und die Angabe „ja“ oder „nein“, ob das Umsatzziel erreicht wurde. Ergänzen Sie die Summenzeile (G8 bis I8 sowie K8). (18 Punkte)",
         "Ermitteln Sie die Kennzahlen (H11 bis H14): Gesamtumsatz laut Liste, durchschnittlicher Umsatz je Verkauf, größter und kleinster Einzelumsatz. (8 Punkte)",
         "Kontrollieren Sie Ihr Ergebnis (H15, H16): Berechnen Sie die Differenz zwischen dem Gesamtumsatz laut Liste und der Summe der Filialumsätze (auf Cent gerundet). Zeigen Sie in H16 den Text „stimmt“, wenn die Differenz 0 beträgt, sonst „Abweichung“. (6 Punkte)",
         "Erstellen Sie die Auswertung je Warengruppe (G24 bis J28): Anzahl, Umsatz und Anteil am Gesamtumsatz. (8 Punkte)",
         "Beantworten Sie in Zelle G31 in ein bis zwei Sätzen: Welche Filiale hat den höchsten Umsatz, und wie viele Filialen haben ihr Umsatzziel erreicht? (4 Punkte)",
         "Formatieren Sie die Auswertung nach dem Corporate Design der BüroComfort GmbH und ergänzen Sie die Fußzeile. (6 Punkte)"],
        "Insgesamt 50 Punkte. " + BASIS + " Bei der Auswertung hilft Ihnen die Zählfunktion mit Bedingung und die Summenfunktion mit Bedingung. Der Anteil ist Teil durch Ganzes.",
        "Lösung (Werte nach Recalc im Modul angegeben): H4 =ZÄHLENWENN($B$4:$B$35;G4); I4 =SUMMEWENN($B$4:$B$35;G4;$D$4:$D$35); J4 =I4/H4; K4 =I4/$I$8; L4 =RANG(I4;$I$4:$I$7;0); "
        "M4 =WENN(I4>=$H$20;\"ja\";\"nein\"); H8 =SUMME(H4:H7); I8 =SUMME(I4:I7); K8 =SUMME(K4:K7); H11 =SUMME(D4:D35); H12 =MITTELWERT(D4:D35); H13 =MAX(D4:D35); H14 =MIN(D4:D35); "
        "H15 =RUNDEN(H11-I8;2); H16 =WENN(H15=0;\"stimmt\";\"Abweichung\"). Warengruppen: H25 =ZÄHLENWENN($C$4:$C$35;G25); I25 =SUMMEWENN($C$4:$C$35;G25;$D$4:$D$35); J25 =I25/$I$29. "
        "Punkte: Filiale 18, Kennzahlen 8, Kontrolle 6, Warengruppen 8, Antwortsatz 4, Gestaltung 6.", loesung)
    ws = wb.create_sheet("Verkäufe")
    titel_zeile(ws, "BüroComfort GmbH – Verkäufe im Monat", 13)
    tabelle_kopf(ws, 3, ["Beleg-Nr.", "Filiale", "Warengruppe", "Umsatz"])
    daten_zellen(ws, 4, P2_DATEN, [None, None, None, EURO])
    tabelle_kopf(ws, 3, ["Filiale", "Anzahl", "Umsatz", "Ø je Verkauf", "Anteil", "Rang", "Ziel erreicht"], 7)
    # Spalten G..M: G Filiale, H Anzahl, I Umsatz, J Ø, K Anteil, L Rang, M Ziel
    for i, f in enumerate(P2_FIL):
        r = 4 + i
        beschriftung(ws, f"G{r}", f)
        ziel(ws, f"H{r}", f"=COUNTIF($B$4:$B$35,G{r})", ZAHL, loesung)
        ziel(ws, f"I{r}", f"=SUMIF($B$4:$B$35,G{r},$D$4:$D$35)", EURO, loesung)
        ziel(ws, f"J{r}", f"=I{r}/H{r}", EURO, loesung)
        ziel(ws, f"K{r}", f"=I{r}/$I$8", PROZ, loesung)
        ziel(ws, f"L{r}", f"=RANK(I{r},$I$4:$I$7,0)", ZAHL, loesung)
        ziel(ws, f"M{r}", f'=IF(I{r}>=$H$20,"ja","nein")', None, loesung)
    beschriftung(ws, "G8", "Gesamt", True)
    ziel(ws, "H8", "=SUM(H4:H7)", ZAHL, loesung, True)
    ziel(ws, "I8", "=SUM(I4:I7)", EURO, loesung, True)
    ziel(ws, "K8", "=SUM(K4:K7)", PROZ, loesung, True)
    tabelle_kopf(ws, 10, ["Kennzahl", "Wert"], 7)
    for r, t in ((11, "Gesamtumsatz laut Liste"), (12, "Ø Umsatz je Verkauf"), (13, "Größter Einzelumsatz"), (14, "Kleinster Einzelumsatz"),
                 (15, "Differenz zur Filialsumme"), (16, "Kontrolle")):
        beschriftung(ws, f"G{r}", t)
    ziel(ws, "H11", "=SUM(D4:D35)", EURO, loesung)
    ziel(ws, "H12", "=AVERAGE(D4:D35)", EURO, loesung)
    ziel(ws, "H13", "=MAX(D4:D35)", EURO, loesung)
    ziel(ws, "H14", "=MIN(D4:D35)", EURO, loesung)
    ziel(ws, "H15", "=ROUND(H11-I8,2)", EURO, loesung)
    ziel(ws, "H16", '=IF(H15=0,"stimmt","Abweichung")', None, loesung)
    tabelle_kopf(ws, 19, ["Vorgabe", "Wert"], 7)
    beschriftung(ws, "G20", "Umsatzziel je Filiale")
    c = ws["H20"]; c.value = 2200.0; c.font = F10; c.border = RAHMEN; c.number_format = EURO
    tabelle_kopf(ws, 24, ["Warengruppe", "Anzahl", "Umsatz", "Anteil"], 7)
    for i, wg in enumerate(P2_WG):
        beschriftung(ws, f"G{25 + i}", wg)
    for i, wg in enumerate(P2_WG):
        r = 25 + i
        ziel(ws, f"H{r}", f"=COUNTIF($C$4:$C$35,G{r})", ZAHL, loesung)
        ziel(ws, f"I{r}", f"=SUMIF($C$4:$C$35,G{r},$D$4:$D$35)", EURO, loesung)
        ziel(ws, f"J{r}", f"=I{r}/$I$29", PROZ, loesung)
    beschriftung(ws, "G29", "Gesamt", True)
    ziel(ws, "H29", "=SUM(H25:H28)", ZAHL, loesung, True)
    ziel(ws, "I29", "=SUM(I25:I28)", EURO, loesung, True)
    ziel(ws, "J29", "=SUM(J25:J28)", PROZ, loesung, True)
    beschriftung(ws, "G31", "Antwort")
    breiten(ws, {"A": 11, "B": 13, "C": 15, "D": 12, "E": 3, "F": 3, "G": 28, "H": 14, "I": 14, "J": 14, "K": 10, "L": 8, "M": 14})
    fuss(ws, loesung)
    return speichern(wb, "p2-auswertung", loesung)


# ------------------------------------------------------------------ p3
P3_STICHTAG = date(2026, 5, 15)
P3_RECH = [("R-2026-0401", "K-4101", date(2026, 2, 16), 30, 1250.00), ("R-2026-0402", "K-4102", date(2026, 4, 1), 14, 480.50),
           ("R-2026-0403", "K-4103", date(2026, 4, 16), 14, 2310.00), ("R-2026-0404", "K-4104", date(2026, 3, 31), 14, 735.80),
           ("R-2026-0405", "K-4105", date(2026, 4, 17), 14, 159.90), ("R-2026-0406", "K-4106", date(2026, 5, 2), 30, 3890.00),
           ("R-2026-0407", "K-4107", date(2026, 4, 25), 14, 295.00), ("R-2026-0408", "K-4108", date(2026, 5, 6), 20, 870.00),
           ("R-2026-0409", "K-4101", date(2026, 3, 24), 14, 125.00), ("R-2026-0410", "K-4102", date(2026, 5, 1), 14, 410.40),
           ("R-2026-0411", "K-4103", date(2026, 5, 8), 30, 1640.00), ("R-2026-0412", "K-4104", date(2026, 4, 10), 14, 560.00)]
P3_TEL = [(1, "R-2026-0401", time(9, 5), time(9, 12)), (2, "R-2026-0402", time(9, 40), time(9, 58)), (3, "R-2026-0404", time(10, 15), time(10, 21)),
          (4, "R-2026-0409", time(11, 2), time(11, 37)), (5, "R-2026-0403", time(13, 30), time(13, 44)), (6, "R-2026-0405", time(14, 5), time(14, 9)),
          (7, "R-2026-0412", time(15, 20), time(15, 52)), (8, "R-2026-0407", time(16, 10), time(16, 25))]
STATI = ["offen", "Zahlungserinnerung", "1. Mahnung", "2. Mahnung"]


def p3(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Prüfungsnahe Aufgabe: Mahnwesen und Telefonzeiten",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Für das Mahnwesen soll zum Stichtag (Zelle B3) geprüft werden, welche Rechnungen überfällig sind. "
        "Regeln: Eine Rechnung ist ab dem Tag nach dem Fälligkeitsdatum überfällig. Bis 14 Tage Überfälligkeit erhält der Kunde eine „Zahlungserinnerung“, "
        "von 15 bis 30 Tagen eine „1. Mahnung“, ab 31 Tagen eine „2. Mahnung“. Rechnungen, die nicht überfällig sind, erhalten den Status „offen“. "
        "Für überfällige Rechnungen werden Verzugszinsen berechnet. Der Zinssatz steht in Zelle B4. Zinstage werden zwischen Fälligkeit und Stichtag mit 30-Tage-Monaten (360-Tage-Jahr) gezählt. "
        "Außerdem wurden Telefonate zu offenen Rechnungen geführt (Blatt „Telefonate“).",
        ["Berechnen Sie für jede Rechnung das Fälligkeitsdatum (F7 bis F18) und die Tage, die die Rechnung zum Stichtag überfällig ist (G7 bis G18). Ist die Rechnung nicht überfällig, soll 0 stehen. (12 Punkte)",
         "Ermitteln Sie für jede Rechnung den Status (H7 bis H18). (9 Punkte)",
         "Berechnen Sie für jede Rechnung die Verzugstage mit einer Funktion für 30-Tage-Monate (I7 bis I18) und die Verzugszinsen (J7 bis J18), auf Cent gerundet. Bei nicht überfälligen Rechnungen soll 0 stehen. (10 Punkte)",
         "Erstellen Sie die Auswertung (Zellen C21 bis C27): Anzahl der Rechnungen je Status, Summe der überfälligen Rechnungsbeträge, Summe der Verzugszinsen und größte Überfälligkeit in Tagen. (10 Punkte)",
         "Blatt „Telefonate“: Berechnen Sie die Dauer jedes Telefonats in Minuten (Spalte E) sowie die Gesamtdauer in Minuten, die durchschnittliche Dauer und die längste Dauer (E13 bis E15). (6 Punkte)",
         "Formatieren Sie die Blätter nach dem Corporate Design der BüroComfort GmbH und ergänzen Sie die Fußzeile. Datumsangaben im Format TT.MM.JJJJ, Uhrzeiten im Format hh:mm. (6 Punkte)"],
        "Insgesamt 53 Punkte. Achten Sie auf die Grenzwerte: Eine Rechnung, die am Stichtag fällig wird, ist noch nicht überfällig. " + BASIS,
        "Lösung: F7 =C7+D7; G7 =WENN($B$3>F7;$B$3-F7;0); H7 =WENN(G7=0;\"offen\";WENN(G7<=14;\"Zahlungserinnerung\";WENN(G7<=30;\"1. Mahnung\";\"2. Mahnung\"))); "
        "I7 =WENN(G7>0;TAGE360(F7;$B$3);0); J7 =RUNDEN(E7*$B$4*I7/360;2); Auswertung C21..C24 =ZÄHLENWENN($H$7:$H$18;B21); C25 =SUMMEWENN(G7:G18;\">0\";E7:E18); C26 =SUMME(J7:J18); C27 =MAX(G7:G18). "
        "Telefonate: E4 =STUNDE(D4-C4)*60+MINUTE(D4-C4); E13 =SUMME(E4:E11); E14 =MITTELWERT(E4:E11); E15 =MAX(E4:E11). "
        "Punkte: Fälligkeit/Überfälligkeit 12, Status 9, Zinsen 10, Auswertung 10, Telefonate 6, Gestaltung 6.", loesung)
    ws = wb.create_sheet("Offene Posten")
    titel_zeile(ws, "BüroComfort GmbH – Offene Posten und Mahnwesen", 10)
    beschriftung(ws, "A3", "Stichtag", True); beschriftung(ws, "A4", "Zinssatz p. a.", True)
    c = ws["B3"]; c.value = P3_STICHTAG; c.font = F10; c.border = RAHMEN; c.number_format = DATUM
    c = ws["B4"]; c.value = 0.05; c.font = F10; c.border = RAHMEN; c.number_format = PROZ
    tabelle_kopf(ws, 6, ["Rechnungs-Nr.", "Kunden-Nr.", "Rechnungsdatum", "Zahlungsziel in Tagen", "Betrag", "Fällig am", "Tage überfällig", "Status", "Verzugstage", "Verzugszinsen"])
    daten_zellen(ws, 7, P3_RECH, [None, None, DATUM, ZAHL, EURO])
    for r in range(7, 19):
        ziel(ws, f"F{r}", f"=C{r}+D{r}", DATUM, loesung)
        ziel(ws, f"G{r}", f"=IF($B$3>F{r},$B$3-F{r},0)", ZAHL, loesung)
        ziel(ws, f"H{r}", f'=IF(G{r}=0,"offen",IF(G{r}<=14,"Zahlungserinnerung",IF(G{r}<=30,"1. Mahnung","2. Mahnung")))', None, loesung)
        ziel(ws, f"I{r}", f"=IF(G{r}>0,DAYS360(F{r},$B$3),0)", ZAHL, loesung)
        ziel(ws, f"J{r}", f"=ROUND(E{r}*$B$4*I{r}/360,2)", EURO, loesung)
    tabelle_kopf(ws, 20, ["Auswertung", "Status", "Wert"], 1)
    ws["A20"].value = "Auswertung"; ws["B20"].value = "Gegenstand"
    for i, s in enumerate(STATI):
        r = 21 + i
        beschriftung(ws, f"A{r}", "Anzahl Rechnungen")
        beschriftung(ws, f"B{r}", s)
        ziel(ws, f"C{r}", f"=COUNTIF($H$7:$H$18,B{r})", ZAHL, loesung)
    for r, a, b in ((25, "Summe", "überfällige Rechnungsbeträge"), (26, "Summe", "Verzugszinsen"), (27, "Größte Überfälligkeit", "in Tagen")):
        beschriftung(ws, f"A{r}", a); beschriftung(ws, f"B{r}", b)
    ziel(ws, "C25", '=SUMIF(G7:G18,">0",E7:E18)', EURO, loesung)
    ziel(ws, "C26", "=SUM(J7:J18)", EURO, loesung)
    ziel(ws, "C27", "=MAX(G7:G18)", ZAHL, loesung)
    breiten(ws, {"A": 20, "B": 28, "C": 16, "D": 20, "E": 14, "F": 14, "G": 16, "H": 20, "I": 14, "J": 16})
    fuss(ws, loesung)
    t = wb.create_sheet("Telefonate")
    titel_zeile(t, "BüroComfort GmbH – Telefonate zu offenen Rechnungen", 5)
    tabelle_kopf(t, 3, ["Nr.", "Rechnungs-Nr.", "Beginn", "Ende", "Dauer in Minuten"])
    for r, (n, rn, b, e) in enumerate(P3_TEL, 4):
        for col, v, fmt in ((1, n, None), (2, rn, None), (3, b, "hh:mm"), (4, e, "hh:mm")):
            c = t.cell(row=r, column=col, value=v); c.font = F10; c.border = RAHMEN
            if fmt: c.number_format = fmt
        ziel(t, f"E{r}", f"=HOUR(D{r}-C{r})*60+MINUTE(D{r}-C{r})", ZAHL, loesung)
    for r, txt in ((13, "Gesamtdauer"), (14, "Durchschnittliche Dauer"), (15, "Längstes Telefonat")):
        beschriftung(t, f"A{r}", txt, True)
    ziel(t, "E13", "=SUM(E4:E11)", ZAHL, loesung)
    ziel(t, "E14", "=AVERAGE(E4:E11)", ZAHL, loesung)
    ziel(t, "E15", "=MAX(E4:E11)", ZAHL, loesung)
    breiten(t, {"A": 24, "B": 16, "C": 10, "D": 10, "E": 18})
    fuss(t, loesung)
    return speichern(wb, "p3-mahnwesen", loesung)


# ------------------------------------------------------------------ pm
def _pm_daten():
    rng = random.Random(77)
    zeilen = []
    start = date(2026, 4, 2)
    for i in range(24):
        tage = i * 3 + rng.randint(0, 2)
        d = date.fromordinal(start.toordinal() + tage)
        kunde = rng.choice([k[0] for k in KUNDEN])
        art = rng.choice([s[0] for s in STAMM])
        preis = {s[0]: s[2] for s in STAMM}[art]
        if preis > 100: menge = rng.randint(1, 6)
        elif preis > 20: menge = rng.randint(2, 20)
        else: menge = rng.randint(10, 120)
        zeilen.append((f"B-{1101 + i}", d, kunde, art, menge))
    return zeilen


PM_DATEN = _pm_daten()
PM_STICHTAG = date(2026, 6, 15)


def pm(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Meilenstein: Prüfungsnahe Gesamtaufgabe – Quartalsauswertung der Bestellungen",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Für das zweite Quartal liegen im Blatt „Bestellungen“ 24 Bestellungen vor. Bekannt sind nur Bestellnummer, Bestelldatum, Kunden-Nr., Artikelnummer und Menge. "
        "Alle weiteren Angaben holen Sie aus dem Blatt „Stammdaten“: Artikelstamm (Bezeichnung, Nettopreis), Kundenstamm (Name, Ort, Kundengruppe), "
        "Rabattstaffel auf den Nettowert (5 % ab 100,00 €, 10 % ab 500,00 €, 15 % ab 1.000,00 €) und Zahlungsziele je Kundengruppe. "
        "Die Umsatzsteuer beträgt 19 % (Zelle E2). Beträge werden kaufmännisch auf zwei Nachkommastellen gerundet. Der Stichtag für die Fälligkeitsprüfung steht in Zelle B2.",
        ["Blatt „Bestellungen“, Spalten F bis I: Holen Sie Kundenname, Kundengruppe, Artikelbezeichnung und Einzelpreis (netto) aus den Stammdaten. (12 Punkte)",
         "Spalten J bis O: Berechnen Sie Nettowert (Menge mal Einzelpreis, gerundet), Rabattsatz aus der Staffel, Rabatt in Euro (gerundet), Nettobetrag, Umsatzsteuer (gerundet) und Bruttobetrag. (18 Punkte)",
         "Spalten P und Q: Berechnen Sie das Fälligkeitsdatum (Bestelldatum plus Zahlungsziel der Kundengruppe) und den Status zum Stichtag: „überfällig“, wenn das Fälligkeitsdatum vor dem Stichtag liegt, sonst „offen“. (10 Punkte)",
         "Blatt „Auswertung“: Erstellen Sie die Auswertung je Kundengruppe (Anzahl, Bruttoumsatz, Anteil, Rang nach Umsatz) und die Kennzahlen (Anzahl der Bestellungen, Summe, Durchschnitt, größter und kleinster Bruttobetrag, Anzahl und Summe der überfälligen Bestellungen). (16 Punkte)",
         "Kontrollieren Sie Ihr Ergebnis in der Auswertung (Zelle B27: Differenz zwischen der Summe der Kundengruppen und dem Gesamtbetrag, auf Cent gerundet) und beantworten Sie in Zelle A29 in einem Satz: Welche Kundengruppe hat den höchsten Umsatz? (6 Punkte)",
         "Formatieren Sie beide Blätter nach dem Corporate Design der BüroComfort GmbH und ergänzen Sie die Fußzeile. (8 Punkte)"],
        "Insgesamt 70 Punkte. Planen Sie Ihre Zeit: Erst die Spalten der Bestellungen, dann die Auswertung. Prüfen Sie Ihre Ergebnisse mit Stichproben. " + BASIS,
        "Lösung: F5 =SVERWEIS(C5;Stammdaten!$F$4:$I$11;2;FALSCH); G5 Index 4; H5 =SVERWEIS(D5;Stammdaten!$A$4:$D$15;2;FALSCH); I5 Index 3; J5 =RUNDEN(E5*I5;2); "
        "K5 =SVERWEIS(J5;Stammdaten!$K$4:$L$7;2;WAHR); L5 =RUNDEN(J5*K5;2); M5 =J5-L5; N5 =RUNDEN(M5*$E$2;2); O5 =M5+N5; "
        "P5 =B5+SVERWEIS(G5;Stammdaten!$N$4:$O$6;2;FALSCH); Q5 =WENN(P5<$B$2;\"überfällig\";\"offen\"). Auswertung: Anzahl =ZÄHLENWENN, Umsatz =SUMMEWENN auf Spalte O, Anteil, Rang; "
        "Kennzahlen =ANZAHL2, =SUMME, =MITTELWERT, =MAX, =MIN, =ZÄHLENWENN(Q;\"überfällig\"), =SUMMEWENN(Q;\"überfällig\";O). Punkte: Stammdaten 12, Rechnung 18, Fälligkeit 10, Auswertung 16, Kontrolle 6, Gestaltung 8.", loesung)
    ws = wb.create_sheet("Bestellungen")
    titel_zeile(ws, "BüroComfort GmbH – Bestellungen 2. Quartal", 17)
    beschriftung(ws, "A2", "Stichtag", True); beschriftung(ws, "D2", "Umsatzsteuersatz", True)
    c = ws["B2"]; c.value = PM_STICHTAG; c.font = F10; c.border = RAHMEN; c.number_format = DATUM
    c = ws["E2"]; c.value = 0.19; c.font = F10; c.border = RAHMEN; c.number_format = PROZ
    tabelle_kopf(ws, 4, ["Bestell-Nr.", "Bestelldatum", "Kunden-Nr.", "Artikel-Nr.", "Menge", "Kundenname", "Kundengruppe", "Artikelbezeichnung",
                         "Einzelpreis netto", "Nettowert", "Rabattsatz", "Rabatt in €", "Nettobetrag", "Umsatzsteuer", "Bruttobetrag", "Fällig am", "Status"])
    daten_zellen(ws, 5, PM_DATEN, [None, DATUM, None, None, ZAHL])
    for r in range(5, 29):
        ziel(ws, f"F{r}", f"=VLOOKUP(C{r},Stammdaten!$F$4:$I$11,2,FALSE)", None, loesung)
        ziel(ws, f"G{r}", f"=VLOOKUP(C{r},Stammdaten!$F$4:$I$11,4,FALSE)", None, loesung)
        ziel(ws, f"H{r}", f"=VLOOKUP(D{r},Stammdaten!$A$4:$D$15,2,FALSE)", None, loesung)
        ziel(ws, f"I{r}", f"=VLOOKUP(D{r},Stammdaten!$A$4:$D$15,3,FALSE)", EURO, loesung)
        ziel(ws, f"J{r}", f"=ROUND(E{r}*I{r},2)", EURO, loesung)
        ziel(ws, f"K{r}", f"=VLOOKUP(J{r},Stammdaten!$K$4:$L$7,2,TRUE)", PROZ, loesung)
        ziel(ws, f"L{r}", f"=ROUND(J{r}*K{r},2)", EURO, loesung)
        ziel(ws, f"M{r}", f"=J{r}-L{r}", EURO, loesung)
        ziel(ws, f"N{r}", f"=ROUND(M{r}*$E$2,2)", EURO, loesung)
        ziel(ws, f"O{r}", f"=M{r}+N{r}", EURO, loesung)
        ziel(ws, f"P{r}", f"=B{r}+VLOOKUP(G{r},Stammdaten!$N$4:$O$6,2,FALSE)", DATUM, loesung)
        ziel(ws, f"Q{r}", f'=IF(P{r}<$B$2,"überfällig","offen")', None, loesung)
    breiten(ws, {"A": 12, "B": 14, "C": 12, "D": 12, "E": 8, "F": 28, "G": 14, "H": 28, "I": 16, "J": 14, "K": 11, "L": 12, "M": 14, "N": 14, "O": 14, "P": 13, "Q": 12})
    fuss(ws, loesung)
    a = wb.create_sheet("Auswertung")
    titel_zeile(a, "BüroComfort GmbH – Auswertung 2. Quartal", 5)
    tabelle_kopf(a, 3, ["Kundengruppe", "Anzahl", "Bruttoumsatz", "Anteil", "Rang"])
    for i, g in enumerate("ABC"):
        r = 4 + i
        beschriftung(a, f"A{r}", g)
        ziel(a, f"B{r}", f"=COUNTIF(Bestellungen!$G$5:$G$28,A{r})", ZAHL, loesung)
        ziel(a, f"C{r}", f"=SUMIF(Bestellungen!$G$5:$G$28,A{r},Bestellungen!$O$5:$O$28)", EURO, loesung)
        ziel(a, f"D{r}", f"=C{r}/$C$7", PROZ, loesung)
        ziel(a, f"E{r}", f"=RANK(C{r},$C$4:$C$6,0)", ZAHL, loesung)
    beschriftung(a, "A7", "Gesamt", True)
    ziel(a, "B7", "=SUM(B4:B6)", ZAHL, loesung, True)
    ziel(a, "C7", "=SUM(C4:C6)", EURO, loesung, True)
    ziel(a, "D7", "=SUM(D4:D6)", PROZ, loesung, True)
    tabelle_kopf(a, 10, ["Kennzahl", "Wert"])
    kz = [("Anzahl der Bestellungen", "=COUNTA(Bestellungen!A5:A28)", ZAHL), ("Summe der Bruttobeträge", "=SUM(Bestellungen!O5:O28)", EURO),
          ("Durchschnittlicher Bruttobetrag", "=AVERAGE(Bestellungen!O5:O28)", EURO), ("Größter Bruttobetrag", "=MAX(Bestellungen!O5:O28)", EURO),
          ("Kleinster Bruttobetrag", "=MIN(Bestellungen!O5:O28)", EURO), ("Anzahl überfälliger Bestellungen", '=COUNTIF(Bestellungen!Q5:Q28,"überfällig")', ZAHL),
          ("Summe überfälliger Bruttobeträge", '=SUMIF(Bestellungen!Q5:Q28,"überfällig",Bestellungen!O5:O28)', EURO)]
    for i, (t, f, fmt) in enumerate(kz):
        r = 11 + i
        beschriftung(a, f"A{r}", t)
        ziel(a, f"B{r}", f, fmt, loesung)
    tabelle_kopf(a, 26, ["Kontrolle", "Differenz"])
    beschriftung(a, "A27", "Gruppensumme minus Gesamtbetrag")
    ziel(a, "B27", "=ROUND(C7-B12,2)", EURO, loesung)
    beschriftung(a, "A29", "Antwort:", True)
    breiten(a, {"A": 36, "B": 16, "C": 18, "D": 12, "E": 10})
    fuss(a, loesung)
    w = wb.create_sheet("Stammdaten")
    titel_zeile(w, "BüroComfort GmbH – Stammdaten", 15)
    stamm_block(w, 3, 1); kunden_block(w, 3, 6)
    tabelle_kopf(w, 3, ["Nettowert ab", "Rabattsatz"], 11)
    daten_zellen(w, 4, RSTAFFEL, [EURO, PROZ], 11)
    ziele_block(w, 3, 14)
    breiten(w, {"A": 10, "B": 30, "C": 14, "D": 9, "E": 3, "F": 11, "G": 28, "H": 13, "I": 14, "J": 3, "K": 15, "L": 11, "M": 3, "N": 14, "O": 22})
    return speichern(wb, "pm-gesamtaufgabe", loesung)


if __name__ == "__main__":
    for fn in (p1, p2, p3, pm):
        for l in (False, True):
            print(fn(l))
