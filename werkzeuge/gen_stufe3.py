#!/usr/bin/env python3
"""Übungsdateien Stufe 3 (n1–n5, nm). Aufruf: python3 gen_stufe3.py <out_aufgabe> <out_loesung>"""
from datetime import date, time
from gen_stufe1 import *

DATUM = "DD.MM.YYYY"
BASIS = "Verwenden Sie in Ihren Formeln Zellbezüge und keine eingetippten Ergebnisse."

STAMM = [("A-1001", "Schreibtischstuhl Basic", 129.00, "Stück"), ("A-1002", "Schreibtisch höhenverstellbar", 489.90, "Stück"),
         ("A-1003", "Aktenschrank 2 Ordnerhöhen", 219.00, "Stück"), ("A-1004", "Ordner A4 breit", 2.49, "Stück"),
         ("A-1005", "Kugelschreiber", 0.89, "Stück"), ("A-1006", "Kopierpapier 500 Blatt", 4.95, "Paket"),
         ("A-1007", "Tacker", 7.90, "Stück"), ("A-1008", "Locher", 6.40, "Stück"),
         ("A-1009", "Notizblock A5", 1.35, "Stück"), ("A-1010", "Bürolampe LED", 34.90, "Stück"),
         ("A-1011", "Haftnotizen", 2.10, "Block"), ("A-1012", "Mauspad", 3.75, "Stück")]
KUNDEN = [("K-4101", "Beispielbau GmbH", "Rendsburg", "A"), ("K-4102", "Musterhandel Nord KG", "Kiel", "B"),
          ("K-4103", "Modellgarten e. K.", "Flensburg", "C"), ("K-4104", "Probe-Druck GmbH", "Neumünster", "A"),
          ("K-4105", "Demo-Pension Förde", "Kiel", "C"), ("K-4106", "Testwerk Elektro GmbH", "Eckernförde", "B"),
          ("K-4107", "Muster-Bäckerei Küste OHG", "Schleswig", "C"), ("K-4108", "Beispiel-Logistik AG", "Rendsburg", "A")]
STAFFEL = [(0, 0.0), (100, 0.05), (500, 0.10), (1000, 0.15)]
ZIELE = [("A", 30), ("B", 20), ("C", 14)]


def stamm_block(ws, r0=3, c0=1):
    tabelle_kopf(ws, r0, ["Art.-Nr.", "Bezeichnung", "Nettopreis", "Einheit"], c0)
    daten_zellen(ws, r0 + 1, STAMM, [None, None, EURO, None], c0)


def kunden_block(ws, r0=3, c0=1):
    tabelle_kopf(ws, r0, ["Kunden-Nr.", "Kundenname", "Ort", "Kundengruppe"], c0)
    daten_zellen(ws, r0 + 1, KUNDEN, [None, None, None, None], c0)


def staffel_block(ws, r0=3, c0=1):
    tabelle_kopf(ws, r0, ["Bestellwert ab", "Rabattsatz"], c0)
    daten_zellen(ws, r0 + 1, STAFFEL, [EURO, PROZ], c0)


def ziele_block(ws, r0=3, c0=1):
    tabelle_kopf(ws, r0, ["Kundengruppe", "Zahlungsziel in Tagen"], c0)
    daten_zellen(ws, r0 + 1, ZIELE, [None, ZAHL], c0)


# ------------------------------------------------------------------ n1
def n1(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Angebot mit dem Artikelstamm erstellen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Für ein Angebot sind nur die Artikelnummern und die Mengen bekannt. Bezeichnung, Einzelpreis und Einheit stehen im Artikelstamm (Tabellenblatt „Artikelstamm“). "
        "Das Angebot liegt im Tabellenblatt „Angebot“ vor.",
        ["Holen Sie für jede Position die Bezeichnung aus dem Artikelstamm (Spalte B).",
         "Holen Sie für jede Position den Einzelpreis aus dem Artikelstamm (Spalte D).",
         "Berechnen Sie den Gesamtpreis jeder Position (Menge mal Einzelpreis, Spalte E) und die Summe des Angebots (Zelle E12).",
         "Zusatzaufgabe: Holen Sie für jede Position die Einheit aus dem Artikelstamm (Spalte F).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Mit einer einzigen Nachschlage-Funktion können Sie die Daten aus dem Artikelstamm holen. Kopieren Sie die Formel nach unten: Der Bereich des Artikelstamms darf dabei nicht wandern. " + BASIS,
        "Lösung: B4 =SVERWEIS(A4;Artikelstamm!$A$4:$D$15;2;FALSCH); D4 mit Spaltenindex 3; F4 mit Spaltenindex 4; E4 =C4*D4; E12 =SUMME(E4:E11). "
        "Erwartet: Summe 2.862,20 €.", loesung)
    ws = wb.create_sheet("Angebot")
    titel_zeile(ws, "BüroComfort GmbH – Angebot (Entwurf)", 6)
    tabelle_kopf(ws, 3, ["Art.-Nr.", "Bezeichnung", "Menge", "Einzelpreis", "Gesamtpreis", "Einheit"])
    pos = [("A-1003", 2), ("A-1007", 5), ("A-1001", 4), ("A-1010", 6), ("A-1004", 40), ("A-1011", 25), ("A-1012", 10), ("A-1002", 3)]
    daten_zellen(ws, 4, [(a, None, m, None, None, None) for a, m in pos], [None, None, ZAHL])
    beschriftung(ws, "D12", "Summe", True)
    for r in range(4, 12):
        ziel(ws, f"B{r}", f"=VLOOKUP(A{r},Artikelstamm!$A$4:$D$15,2,FALSE)", None, loesung)
        ziel(ws, f"D{r}", f"=VLOOKUP(A{r},Artikelstamm!$A$4:$D$15,3,FALSE)", EURO, loesung)
        ziel(ws, f"E{r}", f"=C{r}*D{r}", EURO, loesung)
        ziel(ws, f"F{r}", f"=VLOOKUP(A{r},Artikelstamm!$A$4:$D$15,4,FALSE)", None, loesung)
    ziel(ws, "E12", "=SUM(E4:E11)", EURO, loesung, True)
    breiten(ws, {"A": 12, "B": 32, "C": 10, "D": 14, "E": 16, "F": 10})
    fuss(ws, loesung)
    ws2 = wb.create_sheet("Artikelstamm")
    titel_zeile(ws2, "BüroComfort GmbH – Artikelstamm", 4)
    stamm_block(ws2)
    breiten(ws2, {"A": 12, "B": 32, "C": 14, "D": 10})
    return speichern(wb, "n1-angebot", loesung)


# ------------------------------------------------------------------ n2
N2_BEST = [("B-9001", "K-4103", 99.99), ("B-9002", "K-4101", 100.00), ("B-9003", "K-4105", 499.99), ("B-9004", "K-4102", 500.00),
           ("B-9005", "K-4107", 999.99), ("B-9006", "K-4104", 1000.00), ("B-9007", "K-4106", 245.80), ("B-9008", "K-4108", 1830.00),
           ("B-9009", "K-4101", 62.40), ("B-9010", "K-4103", 720.00)]


def n2(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Rabatt und Kundendaten nachschlagen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Für zehn Bestellungen sind nur die Kunden-Nr. und der Bestellwert bekannt. Name und Ort des Kunden stehen im Blatt „Kunden“. "
        "Der Rabattsatz richtet sich nach der Rabattstaffel im Blatt „Staffel“: 0 % bis unter 100,00 €, 5 % ab 100,00 €, 10 % ab 500,00 €, 15 % ab 1.000,00 €. "
        "Die Bestellungen liegen im Tabellenblatt „Bestellungen“ vor.",
        ["Holen Sie den Kundennamen (Spalte D) und den Ort (Spalte E) aus dem Blatt „Kunden“.",
         "Holen Sie den Rabattsatz (Spalte F) aus der Rabattstaffel. Gesucht ist der Satz der passenden Stufe, nicht ein genauer Treffer.",
         "Berechnen Sie den Rabatt in Euro (Spalte G) und den Zahlbetrag (Spalte H).",
         "Zusatzaufgabe: Holen Sie die Kundengruppe (Spalte I) und das Zahlungsziel in Tagen (Spalte J). Das Zahlungsziel steht im Blatt „Zahlungsziele“.",
         "Prüfen Sie die Grenzfälle (99,99 €, 100,00 €, 499,99 €, 500,00 €, 999,99 € und 1.000,00 €).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Bei Name und Ort suchen Sie einen genauen Treffer. Bei der Rabattstaffel suchen Sie die passende Stufe. Das ist ein Unterschied im letzten Argument der Funktion. " + BASIS,
        "Lösung: D =SVERWEIS(B;Kunden!$A$4:$D$11;2;FALSCH); E Spaltenindex 3; F =SVERWEIS(C;Staffel!$A$4:$B$7;2;WAHR); G =C*F; H =C-G; "
        "I Spaltenindex 4; J =SVERWEIS(I;Zahlungsziele!$A$4:$B$6;2;FALSCH). Erwartet: B-9001 0 %, B-9002 5 %, B-9003 5 %, B-9004 10 %, B-9005 10 %, B-9006 15 %.", loesung)
    ws = wb.create_sheet("Bestellungen")
    titel_zeile(ws, "BüroComfort GmbH – Bestellungen mit Kundendaten", 10)
    tabelle_kopf(ws, 3, ["Bestell-Nr.", "Kunden-Nr.", "Bestellwert", "Kundenname", "Ort", "Rabattsatz", "Rabatt in €", "Zahlbetrag", "Kundengruppe", "Zahlungsziel in Tagen"])
    daten_zellen(ws, 4, [(a, b, c) for a, b, c in N2_BEST], [None, None, EURO])
    for r in range(4, 14):
        ziel(ws, f"D{r}", f"=VLOOKUP(B{r},Kunden!$A$4:$D$11,2,FALSE)", None, loesung)
        ziel(ws, f"E{r}", f"=VLOOKUP(B{r},Kunden!$A$4:$D$11,3,FALSE)", None, loesung)
        ziel(ws, f"F{r}", f"=VLOOKUP(C{r},Staffel!$A$4:$B$7,2,TRUE)", PROZ, loesung)
        ziel(ws, f"G{r}", f"=C{r}*F{r}", EURO, loesung)
        ziel(ws, f"H{r}", f"=C{r}-G{r}", EURO, loesung)
        ziel(ws, f"I{r}", f"=VLOOKUP(B{r},Kunden!$A$4:$D$11,4,FALSE)", None, loesung)
        ziel(ws, f"J{r}", f"=VLOOKUP(I{r},Zahlungsziele!$A$4:$B$6,2,FALSE)", ZAHL, loesung)
    breiten(ws, {"A": 12, "B": 12, "C": 14, "D": 28, "E": 14, "F": 12, "G": 14, "H": 14, "I": 14, "J": 20})
    fuss(ws, loesung)
    w = wb.create_sheet("Kunden"); titel_zeile(w, "BüroComfort GmbH – Kundenstamm", 4); kunden_block(w); breiten(w, {"A": 12, "B": 28, "C": 14, "D": 14})
    w = wb.create_sheet("Staffel"); titel_zeile(w, "BüroComfort GmbH – Rabattstaffel", 2); staffel_block(w); breiten(w, {"A": 16, "B": 12})
    w = wb.create_sheet("Zahlungsziele"); titel_zeile(w, "BüroComfort GmbH – Zahlungsziele", 2); ziele_block(w); breiten(w, {"A": 14, "B": 24})
    return speichern(wb, "n2-kundendaten", loesung)


# ------------------------------------------------------------------ n3
N3_PREISE = [("A-1004", "Ordner A4 breit", 2.49), ("A-1005", "Kugelschreiber", 0.89), ("A-1006", "Kopierpapier 500 Blatt", 4.95),
             ("A-1007", "Tacker", 7.90), ("A-1008", "Locher", 6.40), ("A-1009", "Notizblock A5", 1.35),
             ("A-1011", "Haftnotizen", 2.10), ("A-1012", "Mauspad", 3.75)]
N3_VERP = [("Ordner A4 breit", 130, 12), ("Kugelschreiber", 500, 100), ("Kopierpapier 500 Blatt", 37, 5),
           ("Haftnotizen", 85, 12), ("Notizblock A5", 250, 20), ("Mauspad", 18, 6)]


def n3(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Preise runden in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Verkaufspreise werden aus den Nettopreisen berechnet. Der Umsatzsteuersatz beträgt 19 %. Ein Preis darf nur zwei Nachkommastellen haben. "
        "Außerdem werden Artikel in Kartons geliefert. Die Berechnungen liegen in den Tabellenblättern „Preise“ und „Verpackung“ vor.",
        ["Berechnen Sie in Spalte D den Bruttopreis ohne Rundung (Nettopreis mal 1 plus Steuersatz aus Zelle B14).",
         "Berechnen Sie in Spalte E den Bruttopreis, gerundet auf zwei Nachkommastellen, und in Spalte F die Differenz zwischen D und E.",
         "Ermitteln Sie die Summen der Spalten D, E und F (Zeile 12).",
         "Zusatzaufgabe, Blatt „Verpackung“: Berechnen Sie, wie viele Kartons für jede Bestellmenge mindestens nötig sind (Spalte D) und wie viele Kartons ganz gefüllt werden (Spalte E).",
         "Zusatzaufgabe: Runden Sie den Jahresumsatz der Filiale Kiel (Blatt „Preise“, Zelle C18) auf ganze Euro (C19), auf Zehner (C20) und auf Hunderter (C21).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Für das Runden gibt es drei Funktionen: eine rundet kaufmännisch, eine immer auf, eine immer ab. Die Anzahl der Stellen steuert, wohin gerundet wird. " + BASIS,
        "Lösung: D4 =C4*(1+$B$14); E4 =RUNDEN(D4;2); F4 =D4-E4; Summen in Zeile 12; Verpackung D4 =AUFRUNDEN(B4/C4;0); E4 =ABRUNDEN(B4/C4;0); "
        "C19 =RUNDEN(C18;0) = 184.263; C20 =RUNDEN(C18;-1) = 184.260; C21 =RUNDEN(C18;-2) = 184.300. Summen: D12 = 35,4977 €; E12 = 35,50 €; F12 = -0,0023 €.", loesung)
    ws = wb.create_sheet("Preise")
    titel_zeile(ws, "BüroComfort GmbH – Bruttopreise", 6)
    tabelle_kopf(ws, 3, ["Art.-Nr.", "Artikel", "Nettopreis", "Brutto ungerundet", "Brutto gerundet", "Differenz"])
    daten_zellen(ws, 4, N3_PREISE, [None, None, '#,##0.0000 "€"'])
    for r in range(4, 12):
        ws[f"C{r}"].number_format = EURO
        ziel(ws, f"D{r}", f"=C{r}*(1+$B$14)", '#,##0.0000 "€"', loesung)
        ziel(ws, f"E{r}", f"=ROUND(D{r},2)", EURO, loesung)
        ziel(ws, f"F{r}", f"=D{r}-E{r}", '#,##0.0000 "€"', loesung)
    beschriftung(ws, "A12", "Summe", True)
    for col, fmt in (("D", '#,##0.0000 "€"'), ("E", EURO), ("F", '#,##0.0000 "€"')):
        ziel(ws, f"{col}12", f"=SUM({col}4:{col}11)", fmt, loesung, True)
    beschriftung(ws, "A14", "Umsatzsteuersatz")
    c = ws["B14"]; c.value = 0.19; c.font = F10; c.border = RAHMEN; c.number_format = PROZ
    tabelle_kopf(ws, 17, ["Wert", "", "Betrag"])
    beschriftung(ws, "A18", "Jahresumsatz Kiel")
    c = ws["C18"]; c.value = 184263.48; c.font = F10; c.border = RAHMEN; c.number_format = EURO
    for r, t, d in ((19, "auf ganze Euro gerundet", 0), (20, "auf Zehner gerundet", -1), (21, "auf Hunderter gerundet", -2)):
        beschriftung(ws, f"A{r}", t)
        ziel(ws, f"C{r}", f"=ROUND(C18,{d})", EURO, loesung)
    breiten(ws, {"A": 20, "B": 24, "C": 16, "D": 18, "E": 16, "F": 14})
    fuss(ws, loesung)
    w = wb.create_sheet("Verpackung")
    titel_zeile(w, "BüroComfort GmbH – Kartons für Lieferungen", 5)
    tabelle_kopf(w, 3, ["Artikel", "Bestellmenge", "Stück je Karton", "Kartons nötig", "Kartons voll gefüllt"])
    daten_zellen(w, 4, N3_VERP, [None, ZAHL, ZAHL])
    for r in range(4, 10):
        ziel(w, f"D{r}", f"=ROUNDUP(B{r}/C{r},0)", ZAHL, loesung)
        ziel(w, f"E{r}", f"=ROUNDDOWN(B{r}/C{r},0)", ZAHL, loesung)
    breiten(w, {"A": 26, "B": 14, "C": 16, "D": 14, "E": 20})
    fuss(w, loesung)
    return speichern(wb, "n3-runden", loesung)


# ------------------------------------------------------------------ n4
N4_DATEN = [("Schreibtischstuhl Basic", 48210), ("Schreibtisch höhenverstellbar", 61350), ("Aktenschrank 2 Ordnerhöhen", 22780),
            ("Ordner A4 breit", 15420), ("Kugelschreiber", 9870), ("Kopierpapier 500 Blatt", 31260), ("Tacker", 9870),
            ("Bürolampe LED", 27640), ("Notizblock A5", 12390), ("Haftnotizen", 8130)]


def n4(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Artikel nach Umsatz ordnen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Geschäftsleitung möchte wissen, welche Artikel im Jahresumsatz vorn und welche hinten liegen. Die Jahresumsätze liegen im Tabellenblatt „Jahresumsatz“ vor.",
        ["Ermitteln Sie in Spalte C den Rang jedes Artikels, wobei der Artikel mit dem höchsten Umsatz Rang 1 erhält.",
         "Ermitteln Sie in Spalte D den Rang jedes Artikels, wobei der Artikel mit dem niedrigsten Umsatz Rang 1 erhält.",
         "Zusatzaufgabe: Kennzeichnen Sie in Spalte E die drei umsatzstärksten Artikel mit „Top 3“.",
         "Zusatzaufgabe: Erklären Sie in Zelle B16 in einem Satz, was bei Kugelschreiber und Tacker auffällt.",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Der Rang einer Zahl hängt von allen anderen Zahlen der Liste ab. Setzen Sie deshalb die Liste in der Formel fest, bevor Sie nach unten kopieren. " + BASIS,
        "Lösung: C4 =RANG(B4;$B$4:$B$13;0); D4 =RANG(B4;$B$4:$B$13;1); E4 =WENN(C4<=3;\"Top 3\";\"\"). Schreibtisch höhenverstellbar Rang 1, Schreibtischstuhl 2, Kopierpapier 3. "
        "Kugelschreiber und Tacker haben denselben Umsatz und deshalb denselben Rang (8 absteigend, 2 aufsteigend); der nächste Rang wird übersprungen (Haftnotizen: 10 bzw. 1).", loesung)
    ws = wb.create_sheet("Jahresumsatz")
    titel_zeile(ws, "BüroComfort GmbH – Jahresumsatz je Artikel", 5)
    tabelle_kopf(ws, 3, ["Artikel", "Jahresumsatz", "Rang (absteigend)", "Rang (aufsteigend)", "Platzierung"])
    daten_zellen(ws, 4, N4_DATEN, [None, EURO])
    for r in range(4, 14):
        ziel(ws, f"C{r}", f"=RANK(B{r},$B$4:$B$13,0)", ZAHL, loesung)
        ziel(ws, f"D{r}", f"=RANK(B{r},$B$4:$B$13,1)", ZAHL, loesung)
        ziel(ws, f"E{r}", f'=IF(C{r}<=3,"Top 3","")', None, loesung)
    beschriftung(ws, "A16", "Auffälligkeit")
    breiten(ws, {"A": 32, "B": 18, "C": 18, "D": 18, "E": 14})
    fuss(ws, loesung)
    return speichern(wb, "n4-rang", loesung)


# ------------------------------------------------------------------ n5
N5_RECH = [("R-2026-0301", date(2026, 3, 2), 30), ("R-2026-0302", date(2026, 3, 9), 14), ("R-2026-0303", date(2026, 3, 16), 20),
           ("R-2026-0304", date(2026, 3, 23), 30), ("R-2026-0305", date(2026, 4, 6), 14), ("R-2026-0306", date(2026, 4, 14), 20)]
N5_ZINS = [("R-2026-0301", 1250.00, date(2026, 4, 1), date(2026, 4, 20)), ("R-2026-0302", 480.50, date(2026, 3, 23), date(2026, 4, 22)),
           ("R-2026-0303", 3320.00, date(2026, 4, 5), date(2026, 5, 18)), ("R-2026-0304", 815.90, date(2026, 4, 22), date(2026, 4, 22))]
N5_ZEIT = [("Montag", time(7, 45), time(16, 30)), ("Dienstag", time(8, 0), time(17, 15)), ("Mittwoch", time(7, 30), time(15, 45)),
           ("Donnerstag", time(8, 15), time(17, 0)), ("Freitag", time(7, 50), time(14, 20))]


def n5(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Fristen, Arbeitszeiten und Verzugstage in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Sie überwachen Zahlungsfristen, werten Arbeitszeiten aus und ermitteln Verzugszinsen. "
        "Die Aufgaben liegen in den Tabellenblättern „Rechnungen“, „Arbeitszeit“ und „Verzugszinsen“ vor.",
        ["Blatt „Rechnungen“: Berechnen Sie das Fälligkeitsdatum (Spalte D) und zerlegen Sie das Rechnungsdatum in Jahr, Monat und Tag (Spalten E, F, G).",
         "Zeigen Sie in Zelle B12 das heutige Datum an und berechnen Sie in Spalte H, wie viele Tage bis zur Fälligkeit noch bleiben (negativ = überfällig).",
         "Blatt „Arbeitszeit“: Berechnen Sie die Nettoarbeitszeit (Ende minus Beginn minus Pause, Spalte E), die ganzen Stunden (Spalte F), die Minuten (Spalte G) und die Arbeitszeit in Dezimalstunden (Spalte H).",
         "Blatt „Verzugszinsen“: Ermitteln Sie die Verzugstage zwischen Fälligkeit und Zahlungseingang mit einer Funktion, die mit 30-Tage-Monaten rechnet (Spalte E), und berechnen Sie die Verzugszinsen (Spalte F).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH. Datumsangaben im Format TT.MM.JJJJ, Uhrzeiten im Format hh:mm, Dezimalstunden mit zwei Nachkommastellen."],
        "Datum und Uhrzeit sind für Excel Zahlen. Deshalb können Sie damit rechnen. Prüfen Sie das Zellenformat, wenn statt eines Datums eine Zahl erscheint. " + BASIS,
        "Lösung: D4 =B4+C4; E =JAHR(B4); F =MONAT(B4); G =TAG(B4); B12 =HEUTE(); H4 =D4-$B$12 (hängt vom heutigen Datum ab). "
        "Arbeitszeit: E4 =C4-B4-D4; F =STUNDE(E4); G =MINUTE(E4); H =F4+G4/60. Verzugszinsen: E4 =TAGE360(C4;D4); F4 =B4*$B$11*E4/360. "
        "Erwartet: Fälligkeit R-2026-0301 am 01.04.2026; Nettoarbeitszeit Montag 08:15 = 8,25 Stunden; Verzugstage R-2026-0301: 19, Zins 1.250 € × 5 % × 19 / 360 = 3,30 €.", loesung)
    ws = wb.create_sheet("Rechnungen")
    titel_zeile(ws, "BüroComfort GmbH – Zahlungsfristen", 8)
    tabelle_kopf(ws, 3, ["Rechnungs-Nr.", "Rechnungsdatum", "Zahlungsziel in Tagen", "Fällig am", "Jahr", "Monat", "Tag", "Tage bis Fälligkeit"])
    daten_zellen(ws, 4, N5_RECH, [None, DATUM, ZAHL])
    beschriftung(ws, "A12", "Heutiges Datum", True)
    ziel(ws, "B12", "=TODAY()", DATUM, loesung)
    for r in range(4, 10):
        ziel(ws, f"D{r}", f"=B{r}+C{r}", DATUM, loesung)
        ziel(ws, f"E{r}", f"=YEAR(B{r})", "0", loesung)
        ziel(ws, f"F{r}", f"=MONTH(B{r})", ZAHL, loesung)
        ziel(ws, f"G{r}", f"=DAY(B{r})", ZAHL, loesung)
        ziel(ws, f"H{r}", f"=D{r}-$B$12", ZAHL, loesung)
    breiten(ws, {"A": 16, "B": 16, "C": 22, "D": 14, "E": 8, "F": 8, "G": 8, "H": 20})
    fuss(ws, loesung)
    w = wb.create_sheet("Arbeitszeit")
    titel_zeile(w, "BüroComfort GmbH – Arbeitszeit der Woche (Filiale Rendsburg)", 8)
    tabelle_kopf(w, 3, ["Tag", "Beginn", "Ende", "", "Nettoarbeitszeit", "Stunden", "Minuten", "Dezimalstunden"])
    w.cell(row=3, column=4).value = "Pause"
    for r, (t, b, e) in enumerate(N5_ZEIT, 4):
        for col, v, fmt in ((1, t, None), (2, b, "hh:mm"), (3, e, "hh:mm")):
            c = w.cell(row=r, column=col, value=v); c.font = F10; c.border = RAHMEN
            if fmt: c.number_format = fmt
        c = w.cell(row=r, column=4, value=time(0, 30)); c.font = F10; c.border = RAHMEN; c.number_format = "hh:mm"
    for r in range(4, 9):
        ziel(w, f"E{r}", f"=C{r}-B{r}-D{r}", "hh:mm", loesung)
        ziel(w, f"F{r}", f"=HOUR(E{r})", ZAHL, loesung)
        ziel(w, f"G{r}", f"=MINUTE(E{r})", ZAHL, loesung)
        ziel(w, f"H{r}", f"=F{r}+G{r}/60", "0.00", loesung)
    breiten(w, {"A": 14, "B": 10, "C": 10, "D": 10, "E": 18, "F": 10, "G": 10, "H": 18})
    fuss(w, loesung)
    z = wb.create_sheet("Verzugszinsen")
    titel_zeile(z, "BüroComfort GmbH – Verzugszinsen", 6)
    tabelle_kopf(z, 3, ["Rechnungs-Nr.", "Betrag", "Fällig am", "Zahlungseingang", "Verzugstage", "Verzugszinsen"])
    daten_zellen(z, 4, [(a, b, c, d) for a, b, c, d in N5_ZINS], [None, EURO, DATUM, DATUM])
    for r in range(4, 8):
        ziel(z, f"E{r}", f"=DAYS360(C{r},D{r})", ZAHL, loesung)
        ziel(z, f"F{r}", f"=B{r}*$B$11*E{r}/360", EURO, loesung)
    beschriftung(z, "A11", "Zinssatz p. a.")
    c = z["B11"]; c.value = 0.05; c.font = F10; c.border = RAHMEN; c.number_format = PROZ
    breiten(z, {"A": 16, "B": 14, "C": 14, "D": 18, "E": 14, "F": 16})
    fuss(z, loesung)
    return speichern(wb, "n5-datum-zeit", loesung)


# ------------------------------------------------------------------ nm
NM_POS = [("A-1001", 6), ("A-1010", 6), ("A-1004", 60), ("A-1006", 40), ("A-1009", 120), ("A-1011", 80), ("A-1007", 10), ("A-1012", 12)]


def nm(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Meilenstein: Rechnung mit Artikelstamm erstellen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Für die Rechnung an den Kunden K-4106 sind nur die Artikelnummern und Mengen bekannt. Alle weiteren Angaben holen Sie aus den Stammdaten (Tabellenblatt „Stammdaten“): "
        "Artikelstamm (Bezeichnung, Nettopreis), Kundenstamm (Name, Ort, Kundengruppe), Rabattstaffel auf die Nettosumme (5 % ab 100,00 €, 10 % ab 500,00 €, 15 % ab 1.000,00 €) "
        "und Zahlungsziele je Kundengruppe. Die Umsatzsteuer beträgt 19 %. Preise und Steuer werden kaufmännisch auf zwei Nachkommastellen gerundet.",
        ["Holen Sie Kundenname, Ort, Kundengruppe und Zahlungsziel in Tagen aus den Stammdaten (Zellen B6 bis B9) und berechnen Sie das Fälligkeitsdatum (Zelle B10).",
         "Holen Sie für jede Position Bezeichnung (Spalte C) und Nettopreis (Spalte E) aus dem Artikelstamm und berechnen Sie den Gesamtpreis der Position (Spalte F, auf zwei Nachkommastellen gerundet).",
         "Berechnen Sie die Summe netto (F23), den Rabattsatz aus der Staffel (F24), den Rabatt in Euro (F25), den Nettobetrag nach Rabatt (F26), die Umsatzsteuer (F27) und den Rechnungsbetrag (F28).",
         "Zusatzaufgabe: Zeigen Sie in Zelle B11 den Status der Rechnung an: „überfällig“, wenn das heutige Datum nach dem Fälligkeitsdatum liegt, sonst „noch nicht fällig“.",
         "Formatieren Sie die Rechnung nach dem Corporate Design der BüroComfort GmbH. Datumsangaben im Format TT.MM.JJJJ."],
        "Verwenden Sie in Ihren Formeln Zellbezüge und keine eingetippten Ergebnisse. Der Bereich der Stammdaten darf beim Kopieren nicht wandern. Prüfen Sie die Rechnung mit dem Taschenrechner nach.",
        "Lösung: B6 =SVERWEIS(B5;Stammdaten!$F$4:$I$11;2;FALSCH); B7 Index 3; B8 Index 4; B9 =SVERWEIS(B8;Stammdaten!$N$4:$O$6;2;FALSCH); B10 =B4+B9; "
        "C14 =SVERWEIS(B14;Stammdaten!$A$4:$C$15;2;FALSCH); E14 Index 3; F14 =RUNDEN(D14*E14;2); F23 =SUMME(F14:F21); F24 =SVERWEIS(F23;Stammdaten!$K$4:$L$7;2;WAHR); "
        "F25 =RUNDEN(F23*F24;2); F26 =F23-F25; F27 =RUNDEN(F26*E27;2); F28 =F26+F27; B11 =WENN(HEUTE()>B10;\"überfällig\";\"noch nicht fällig\"). "
        "Erwartet: Testwerk Elektro GmbH, Eckernförde, Gruppe B, 20 Tage, fällig 04.05.2026; Summe netto 1.784,80 €; Rabatt 15 % = 267,72 €; Nettobetrag 1.517,08 €; USt 288,25 €; Rechnungsbetrag 1.805,33 €.", loesung)
    ws = wb.create_sheet("Rechnung")
    titel_zeile(ws, "BüroComfort GmbH – Rechnung", 6)
    kopf = [("Rechnungs-Nr.", "R-2026-0412", None), ("Rechnungsdatum", date(2026, 4, 14), DATUM), ("Kunden-Nr.", "K-4106", None)]
    for i, (t, v, f) in enumerate(kopf):
        beschriftung(ws, f"A{3 + i}", t, True)
        c = ws[f"B{3 + i}"]; c.value = v; c.font = F10; c.border = RAHMEN
        if f: c.number_format = f
    labels = ["Kunde", "Ort", "Kundengruppe", "Zahlungsziel in Tagen", "Fällig am", "Status"]
    for i, t in enumerate(labels):
        beschriftung(ws, f"A{6 + i}", t, True)
    ziel(ws, "B6", "=VLOOKUP(B5,Stammdaten!$F$4:$I$11,2,FALSE)", None, loesung)
    ziel(ws, "B7", "=VLOOKUP(B5,Stammdaten!$F$4:$I$11,3,FALSE)", None, loesung)
    ziel(ws, "B8", "=VLOOKUP(B5,Stammdaten!$F$4:$I$11,4,FALSE)", None, loesung)
    ziel(ws, "B9", "=VLOOKUP(B8,Stammdaten!$N$4:$O$6,2,FALSE)", ZAHL, loesung)
    ziel(ws, "B10", "=B4+B9", DATUM, loesung)
    ziel(ws, "B11", '=IF(TODAY()>B10,"überfällig","noch nicht fällig")', None, loesung)
    tabelle_kopf(ws, 13, ["Pos.", "Art.-Nr.", "Bezeichnung", "Menge", "Einzelpreis netto", "Gesamtpreis netto"])
    # Spalten: A Pos, B Art.-Nr., C Bezeichnung, D Menge, E Einzelpreis, F Gesamtpreis
    for i, (a, m) in enumerate(NM_POS):
        r = 14 + i
        for col, v in ((1, i + 1), (2, a), (4, m)):
            c = ws.cell(row=r, column=col, value=v); c.font = F10; c.border = RAHMEN
        ziel(ws, f"C{r}", f"=VLOOKUP(B{r},Stammdaten!$A$4:$C$15,2,FALSE)", None, loesung)
        ziel(ws, f"E{r}", f"=VLOOKUP(B{r},Stammdaten!$A$4:$C$15,3,FALSE)", EURO, loesung)
        ziel(ws, f"F{r}", f"=ROUND(D{r}*E{r},2)", EURO, loesung)
    zs = [(23, "Summe netto", "=SUM(F14:F21)", EURO), (24, "Rabattsatz", "=VLOOKUP(F23,Stammdaten!$K$4:$L$7,2,TRUE)", PROZ),
          (25, "Rabatt", "=ROUND(F23*F24,2)", EURO), (26, "Nettobetrag nach Rabatt", "=F23-F25", EURO),
          (27, "Umsatzsteuer", "=ROUND(F26*E27,2)", EURO), (28, "Rechnungsbetrag", "=F26+F27", EURO)]
    for r, t, f, fmt in zs:
        beschriftung(ws, f"D{r}", t, r == 28)
        ziel(ws, f"F{r}", f, fmt, loesung, r == 28)
    c = ws["E27"]; c.value = 0.19; c.font = F10; c.border = RAHMEN; c.number_format = PROZ
    breiten(ws, {"A": 22, "B": 22, "C": 30, "D": 26, "E": 18, "F": 18})
    fuss(ws, loesung)
    w = wb.create_sheet("Stammdaten")
    titel_zeile(w, "BüroComfort GmbH – Stammdaten", 15)
    stamm_block(w, 3, 1)
    kunden_block(w, 3, 6)
    staffel_block(w, 3, 11)
    ziele_block(w, 3, 14)
    # Stammdaten: Artikelstamm hat 4 Spalten A:D; Kunden F:I; Staffel K:L; Ziele N:O
    breiten(w, {"A": 10, "B": 30, "C": 14, "D": 9, "E": 3, "F": 11, "G": 28, "H": 13, "I": 14, "J": 3, "K": 15, "L": 11, "M": 3, "N": 14, "O": 22})
    return speichern(wb, "nm-rechnung", loesung)


if __name__ == "__main__":
    for fn in (n1, n2, n3, n4, n5, nm):
        for l in (False, True):
            print(fn(l))
