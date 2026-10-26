#!/usr/bin/env python3
"""Übungsdateien Stufe 2 (z1, z2, z3, zm). Aufruf: python3 gen_stufe2.py <out_aufgabe> <out_loesung>"""
from gen_stufe1 import *

BASIS = "Verwenden Sie in Ihren Formeln Zellbezüge und keine eingetippten Ergebnisse."

# ------------------------------------------------------------------ z1
def z1(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Lagerbestand erfassen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Bei der Inventur wurde der Lagerbestand erfasst. Nicht alle Zeilen sind vollständig: Bei manchen Artikeln fehlt der gezählte Bestand, "
        "bei einem steht „n. v.“ (nicht vorhanden), bei manchen fehlt der Lieferant. Die Liste liegt im Tabellenblatt „Inventur“ vor.",
        ["Ermitteln Sie die Anzahl der Artikel in der Liste (Zelle C18).",
         "Ermitteln Sie, bei wie vielen Artikeln ein Bestand als Zahl erfasst wurde (Zelle C19).",
         "Ermitteln Sie, bei wie vielen Artikeln in der Spalte „Bestand“ überhaupt etwas steht (Zelle C20).",
         "Berechnen Sie, bei wie vielen Artikeln kein Bestand als Zahl vorliegt (Zelle C21).",
         "Ermitteln Sie, bei wie vielen Artikeln ein Lieferant eingetragen ist (Zelle C22).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Es gibt zwei Funktionen zum Zählen. Eine zählt nur Zahlen, die andere zählt alle Zellen, die nicht leer sind. Überlegen Sie bei jedem Auftrag, welche passt. " + BASIS,
        "Lösung: C18 =ANZAHL2(A4:A15) = 12; C19 =ANZAHL(C4:C15) = 9; C20 =ANZAHL2(C4:C15) = 10; C21 =C18-C19 = 3; C22 =ANZAHL2(D4:D15) = 9.", loesung)
    ws = wb.create_sheet("Inventur")
    titel_zeile(ws, "BüroComfort GmbH – Inventur Hauptlager", 4)
    tabelle_kopf(ws, 3, ["Art.-Nr.", "Artikel", "Bestand", "Lieferant"])
    daten = [("A-1001", "Schreibtischstuhl Basic", 24, "Möbelwerk Nord"), ("A-1002", "Schreibtisch höhenverstellbar", 8, "Möbelwerk Nord"),
             ("A-1003", "Aktenschrank 2 Ordnerhöhen", 5, "Möbelwerk Nord"), ("A-1004", "Ordner A4 breit", 340, "Papier & Co. KG"),
             ("A-1005", "Kugelschreiber", 1250, "Schreibwaren Lange"), ("A-1006", "Kopierpapier 500 Blatt", 96, "Papier & Co. KG"),
             ("A-1007", "Tacker", "n. v.", "Schreibwaren Lange"), ("A-1008", "Locher", None, None),
             ("A-1009", "Notizblock A5", 410, "Papier & Co. KG"), ("A-1010", "Bürolampe LED", 17, None),
             ("A-1011", "Haftnotizen", 780, "Schreibwaren Lange"), ("A-1012", "Mauspad", None, None)]
    daten_zellen(ws, 4, daten, [None, None, None, None])
    beschriftung(ws, "A18", "Kennzahlen", True)
    for r, t in zip(range(18, 23), ["Artikel in der Liste", "Bestand als Zahl erfasst", "Einträge in der Spalte Bestand",
                                    "Kein Bestand als Zahl", "Artikel mit Lieferant"]):
        beschriftung(ws, f"B{r}", t)
    ws["A18"].value = None
    ziel(ws, "C18", "=COUNTA(A4:A15)", None, loesung)
    ziel(ws, "C19", "=COUNT(C4:C15)", None, loesung)
    ziel(ws, "C20", "=COUNTA(C4:C15)", None, loesung)
    ziel(ws, "C21", "=C18-C19", None, loesung)
    ziel(ws, "C22", "=COUNTA(D4:D15)", None, loesung)
    ws["A17"].value = "Kennzahlen"
    ws["A17"].font = F10B
    breiten(ws, {"A": 12, "B": 34, "C": 12, "D": 22})
    fuss(ws, loesung)
    return speichern(wb, "z1-inventur", loesung)


# ------------------------------------------------------------------ z2
Z2_DATEN = [
    ("B-6001", "Rendsburg", "Rechnung", 245.00, "versendet"), ("B-6002", "Kiel", "Lastschrift", 89.90, "versendet"),
    ("B-6003", "Flensburg", "Vorkasse", 410.50, "offen"), ("B-6004", "Neumünster", "Rechnung", 200.00, "versendet"),
    ("B-6005", "Kiel", "Rechnung", 1250.00, "versendet"), ("B-6006", "Rendsburg", "Vorkasse", 64.20, "storniert"),
    ("B-6007", "Kiel", "Lastschrift", 199.99, "versendet"), ("B-6008", "Flensburg", "Rechnung", 520.00, "offen"),
    ("B-6009", "Rendsburg", "Lastschrift", 38.50, "versendet"), ("B-6010", "Neumünster", "Vorkasse", 200.00, "offen"),
    ("B-6011", "Kiel", "Vorkasse", 315.75, "versendet"), ("B-6012", "Rendsburg", "Rechnung", 780.00, "versendet"),
    ("B-6013", "Flensburg", "Lastschrift", 129.00, "storniert"), ("B-6014", "Kiel", "Rechnung", 205.40, "offen"),
    ("B-6015", "Neumünster", "Lastschrift", 56.90, "versendet"), ("B-6016", "Rendsburg", "Rechnung", 960.00, "versendet"),
    ("B-6017", "Flensburg", "Vorkasse", 275.30, "versendet"), ("B-6018", "Kiel", "Lastschrift", 143.00, "versendet"),
    ("B-6019", "Rendsburg", "Vorkasse", 92.80, "offen"), ("B-6020", "Neumünster", "Rechnung", 401.00, "versendet"),
]
FILIALEN = ["Rendsburg", "Kiel", "Flensburg", "Neumünster"]


def z2(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Bestellungen zählen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Geschäftsleitung möchte wissen, wie viele Bestellungen in der Kalenderwoche 13 eingegangen sind: je Filiale, nach Zahlungsart und nach Höhe. "
        "Die Bestellungen liegen im Tabellenblatt „Bestellungen KW 13“ vor. Die Auswertung steht rechts neben der Liste.",
        ["Ermitteln Sie die Anzahl der Bestellungen je Filiale (Zellen H4 bis H7) und die Gesamtzahl (Zelle H8).",
         "Berechnen Sie den Anteil jeder Filiale an allen Bestellungen (Zellen I4 bis I7).",
         "Ermitteln Sie die Anzahl der Bestellungen mit einem Bestellwert über 200,00 € (Zelle H11).",
         "Ermitteln Sie die Anzahl der Bestellungen mit Zahlungsart „Vorkasse“ (Zelle H12) und mit dem Status „offen“ (Zelle H13).",
         "Ermitteln Sie die Anzahl der Bestellungen ab dem Schwellenwert in Zelle H16. Schreiben Sie die Formel so, dass sie sich anpasst, wenn Sie den Schwellenwert ändern (Zelle H17).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Für Zählungen mit einer Bedingung gibt es eine eigene Funktion. Bedingungen mit Vergleichszeichen und Text stehen in Anführungszeichen. " + BASIS,
        "Lösung: H4 =ZÄHLENWENN($B$4:$B$23;G4) usw. = 6 / 6 / 4 / 4; H8 =SUMME(H4:H7) = 20; I4 =H4/$H$8; "
        "H11 =ZÄHLENWENN(D4:D23;\">200\"); H12 =ZÄHLENWENN(C4:C23;\"Vorkasse\"); H13 =ZÄHLENWENN(E4:E23;\"offen\"); "
        "H17 =ZÄHLENWENN(D4:D23;\">=\"&H16) = 7 bei 300,00 €. Erwartet: H11 = 10, H12 = 6, H13 = 5. Grenzfälle: zwei Bestellungen mit genau 200,00 € (B-6004, B-6010) zählen bei „>200“ nicht, bei „>=200“ schon.", loesung)
    ws = wb.create_sheet("Bestellungen KW 13")
    titel_zeile(ws, "BüroComfort GmbH – Bestellungen KW 13", 9)
    tabelle_kopf(ws, 3, ["Bestell-Nr.", "Filiale", "Zahlungsart", "Bestellwert", "Status"], 1)
    daten_zellen(ws, 4, Z2_DATEN, [None, None, None, EURO, None])
    # Auswertung G:I
    tabelle_kopf(ws, 3, ["Filiale", "Anzahl", "Anteil"], 7)
    for i, f in enumerate(FILIALEN):
        r = 4 + i
        beschriftung(ws, f"G{r}", f)
        ziel(ws, f"H{r}", f"=COUNTIF($B$4:$B$23,G{r})", None, loesung)
        ziel(ws, f"I{r}", f"=H{r}/$H$8", PROZ, loesung)
    beschriftung(ws, "G8", "Gesamt", True)
    ziel(ws, "H8", "=SUM(H4:H7)", None, loesung, True)
    tabelle_kopf(ws, 10, ["Auswertung", "Anzahl"], 7)
    beschriftung(ws, "G11", "Bestellwert über 200,00 €")
    ziel(ws, "H11", '=COUNTIF(D4:D23,">200")', None, loesung)
    beschriftung(ws, "G12", "Zahlungsart Vorkasse")
    ziel(ws, "H12", '=COUNTIF(C4:C23,"Vorkasse")', None, loesung)
    beschriftung(ws, "G13", "Status offen")
    ziel(ws, "H13", '=COUNTIF(E4:E23,"offen")', None, loesung)
    tabelle_kopf(ws, 15, ["Schwellenwert", "Wert"], 7)
    beschriftung(ws, "G16", "Schwellenwert (Bestellwert ab)")
    c = ws["H16"]; c.value = 300; c.font = F10; c.border = RAHMEN; c.number_format = EURO
    beschriftung(ws, "G17", "Bestellungen ab Schwellenwert")
    ziel(ws, "H17", '=COUNTIF(D4:D23,">="&H16)', None, loesung)
    breiten(ws, {"A": 12, "B": 14, "C": 14, "D": 14, "E": 12, "F": 3, "G": 32, "H": 14, "I": 12})
    fuss(ws, loesung)
    return speichern(wb, "z2-bestellungen", loesung)


# ------------------------------------------------------------------ z3
WG = ["Büromöbel", "Papier", "Schreibwaren", "Technik"]
Z3_DATEN = [
    ("V-7001", "Rendsburg", "Büromöbel", 489.90), ("V-7002", "Kiel", "Papier", 74.50), ("V-7003", "Flensburg", "Technik", 329.00),
    ("V-7004", "Neumünster", "Schreibwaren", 18.90), ("V-7005", "Kiel", "Büromöbel", 1218.00), ("V-7006", "Rendsburg", "Papier", 59.40),
    ("V-7007", "Kiel", "Technik", 899.00), ("V-7008", "Flensburg", "Schreibwaren", 42.30), ("V-7009", "Rendsburg", "Technik", 149.90),
    ("V-7010", "Neumünster", "Büromöbel", 639.00), ("V-7011", "Kiel", "Schreibwaren", 27.80), ("V-7012", "Flensburg", "Papier", 96.00),
    ("V-7013", "Rendsburg", "Schreibwaren", 35.60), ("V-7014", "Neumünster", "Technik", 249.00), ("V-7015", "Kiel", "Papier", 118.20),
    ("V-7016", "Flensburg", "Büromöbel", 770.50), ("V-7017", "Rendsburg", "Büromöbel", 219.00), ("V-7018", "Neumünster", "Papier", 64.90),
    ("V-7019", "Kiel", "Technik", 459.00), ("V-7020", "Flensburg", "Schreibwaren", 51.20), ("V-7021", "Rendsburg", "Technik", 79.99),
    ("V-7022", "Neumünster", "Schreibwaren", 23.40), ("V-7023", "Kiel", "Büromöbel", 345.00), ("V-7024", "Flensburg", "Technik", 500.00),
]


def z3(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Lernsituation: Umsätze nach Filiale und Warengruppe in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Für die Quartalsplanung benötigt die Geschäftsleitung die Umsätze der Kalenderwoche 14, getrennt nach Filiale und nach Warengruppe. "
        "Die einzelnen Verkäufe liegen im Tabellenblatt „Verkäufe KW 14“ vor. Die Auswertungen stehen rechts neben der Liste.",
        ["Ermitteln Sie den Umsatz je Filiale (Zellen H4 bis H7) und den Gesamtumsatz (Zelle H8).",
         "Ermitteln Sie den Umsatz je Warengruppe (Zellen H12 bis H15) und den Gesamtumsatz (Zelle H16).",
         "Prüfen Sie, ob beide Gesamtsummen mit der Summe der Einzelumsätze in Spalte D übereinstimmen (Zelle H18: Differenz).",
         "Ermitteln Sie den Umsatz aller Verkäufe ab 500,00 € (Zelle H21).",
         "Berechnen Sie den Anteil jeder Warengruppe am Gesamtumsatz (Zellen I12 bis I15).",
         "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design der BüroComfort GmbH."],
        "Für Summen mit einer Bedingung gibt es eine eigene Funktion. Sie braucht drei Angaben: wo geprüft wird, was geprüft wird und was addiert wird. " + BASIS,
        "Lösung: H4 =SUMMEWENN($B$4:$B$27;G4;$D$4:$D$27) usw.; H12 =SUMMEWENN($C$4:$C$27;G12;$D$4:$D$27); H8 und H16 =SUMME(…); "
        "H18 =H8-H16 = 0; H21 =SUMMEWENN(D4:D27;\">=500\") = 4.026,50 €; I12 =H12/$H$16. Erwartet: Gesamt 6.959,49 €; Büromöbel 3.681,40 €.", loesung)
    ws = wb.create_sheet("Verkäufe KW 14")
    titel_zeile(ws, "BüroComfort GmbH – Verkäufe KW 14", 9)
    tabelle_kopf(ws, 3, ["Beleg-Nr.", "Filiale", "Warengruppe", "Umsatz"], 1)
    daten_zellen(ws, 4, Z3_DATEN, [None, None, None, EURO])
    tabelle_kopf(ws, 3, ["Filiale", "Umsatz"], 7)
    for i, f in enumerate(FILIALEN):
        r = 4 + i
        beschriftung(ws, f"G{r}", f)
        ziel(ws, f"H{r}", f"=SUMIF($B$4:$B$27,G{r},$D$4:$D$27)", EURO, loesung)
    beschriftung(ws, "G8", "Gesamt", True)
    ziel(ws, "H8", "=SUM(H4:H7)", EURO, loesung, True)
    tabelle_kopf(ws, 11, ["Warengruppe", "Umsatz", "Anteil"], 7)
    for i, w in enumerate(WG):
        r = 12 + i
        beschriftung(ws, f"G{r}", w)
        ziel(ws, f"H{r}", f"=SUMIF($C$4:$C$27,G{r},$D$4:$D$27)", EURO, loesung)
        ziel(ws, f"I{r}", f"=H{r}/$H$16", PROZ, loesung)
    beschriftung(ws, "G16", "Gesamt", True)
    ziel(ws, "H16", "=SUM(H12:H15)", EURO, loesung, True)
    beschriftung(ws, "G18", "Differenz der beiden Gesamtsummen")
    ziel(ws, "H18", "=H8-H16", EURO, loesung)
    tabelle_kopf(ws, 20, ["Auswertung", "Umsatz"], 7)
    beschriftung(ws, "G21", "Umsatz aller Verkäufe ab 500,00 €")
    ziel(ws, "H21", '=SUMIF(D4:D27,">=500")', EURO, loesung)
    breiten(ws, {"A": 12, "B": 14, "C": 16, "D": 14, "E": 3, "F": 3, "G": 34, "H": 16, "I": 12})
    fuss(ws, loesung)
    return speichern(wb, "z3-verkaeufe", loesung)


# ------------------------------------------------------------------ zm
ZM_DATEN = [
    ("B-8001", "Rendsburg", "Büromöbel", "Rechnung", 640.00, "versendet"), ("B-8002", "Kiel", "Technik", "Lastschrift", 899.00, "versendet"),
    ("B-8003", "Flensburg", "Papier", "Vorkasse", 112.40, "versendet"), ("B-8004", "Neumünster", "Schreibwaren", "Rechnung", 46.80, "versendet"),
    ("B-8005", "Kiel", "Büromöbel", "Rechnung", 1480.00, "offen"), ("B-8006", "Rendsburg", "Technik", "Lastschrift", 329.00, "versendet"),
    ("B-8007", "Flensburg", "Büromöbel", "Rechnung", 770.50, "versendet"), ("B-8008", "Rendsburg", "Papier", "Vorkasse", 89.90, "storniert"),
    ("B-8009", "Neumünster", "Technik", "Lastschrift", 249.00, "versendet"), ("B-8010", "Kiel", "Schreibwaren", "Vorkasse", 33.60, "versendet"),
    ("B-8011", "Rendsburg", "Büromöbel", "Rechnung", 219.00, "versendet"), ("B-8012", "Flensburg", "Technik", "Lastschrift", 500.00, "offen"),
    ("B-8013", "Kiel", "Papier", "Rechnung", 205.30, "versendet"), ("B-8014", "Neumünster", "Büromöbel", "Vorkasse", 389.00, "storniert"),
    ("B-8015", "Rendsburg", "Schreibwaren", "Lastschrift", 71.20, "versendet"), ("B-8016", "Kiel", "Technik", "Rechnung", 1199.00, "versendet"),
    ("B-8017", "Flensburg", "Schreibwaren", "Rechnung", 58.90, "versendet"), ("B-8018", "Neumünster", "Papier", "Lastschrift", 143.50, "offen"),
    ("B-8019", "Rendsburg", "Technik", "Vorkasse", None, "offen"), ("B-8020", "Kiel", "Büromöbel", "Lastschrift", 925.00, "versendet"),
    ("B-8021", "Flensburg", "Papier", "Rechnung", 77.00, "versendet"), ("B-8022", "Neumünster", "Technik", "Rechnung", 615.00, "versendet"),
    ("B-8023", "Rendsburg", "Papier", "Lastschrift", 164.80, "versendet"), ("B-8024", "Kiel", "Schreibwaren", "Rechnung", 41.30, "storniert"),
    ("B-8025", "Flensburg", "Büromöbel", "Lastschrift", 1020.00, "versendet"), ("B-8026", "Neumünster", "Schreibwaren", "Vorkasse", 29.90, "versendet"),
    ("B-8027", "Rendsburg", "Technik", "Rechnung", 459.00, "versendet"), ("B-8028", "Kiel", "Papier", "Lastschrift", 98.20, "versendet"),
    ("B-8029", "Flensburg", "Technik", "Vorkasse", 275.00, "versendet"), ("B-8030", "Neumünster", "Büromöbel", "Rechnung", 540.00, "versendet"),
]


def zm(loesung):
    wb = Workbook()
    aufgabenblatt(
        wb, "Meilenstein: Bestellauswertung erstellen in der Abteilung Verwaltung",
        "Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. "
        "Die Geschäftsleitung braucht für die Quartalsbesprechung eine Auswertung aller 30 Bestellungen. "
        "Die Bestellungen liegen im Tabellenblatt „Bestellungen“ vor. Stornierte Bestellungen gehören nicht zum Umsatz. "
        "Bei einer Bestellung fehlt der Bestellwert; sie ist noch nicht bewertet.",
        ["Ermitteln Sie in Spalte G für jede Bestellung den Umsatz ohne Storno: Bei stornierten Bestellungen 0, sonst der Bestellwert.",
         "Ermitteln Sie die Kennzahlen (Zellen B37 bis B43): Anzahl aller Bestellungen, Anzahl der Bestellungen mit Bestellwert, Anzahl ohne Bestellwert, "
         "Anzahl der stornierten und der offenen Bestellungen sowie den Umsatz ohne stornierte Bestellungen, einmal mit SUMMEWENN (Zelle B42) und einmal als Summe der Spalte G (Zelle B43).",
         "Erstellen Sie die Auswertung je Filiale (Zeilen 47 bis 50 und Gesamtzeile 51): Anzahl der Bestellungen, Umsatz ohne Storno und Anteil am Gesamtumsatz.",
         "Erstellen Sie die Auswertung je Warengruppe (Zeilen 55 bis 58 und Gesamtzeile 59): Anzahl der Bestellungen, Umsatz ohne Storno und Anteil am Gesamtumsatz.",
         "Prüfen Sie Ihre Ergebnisse: Die Umsätze der Filialen und der Warengruppen müssen jeweils den Umsatz ohne Storno ergeben. Tragen Sie die Differenzen in die Zellen B62 und B63 ein.",
         "Formatieren Sie die Tabelle nach dem Corporate Design der BüroComfort GmbH."],
        "Nutzen Sie die Zähl- und Summenfunktionen, die Sie kennen. " + BASIS + " Prüfen Sie am Ende, ob die Gesamtzahlen zueinander passen.",
        "Lösung (Werte nach Recalc prüfen): G =WENN(F=\"storniert\";0;E); B37 =ANZAHL2(A4:A33); B38 =ANZAHL(E4:E33); B39 =B37-B38; B40 =ZÄHLENWENN(F4:F33;\"storniert\"); "
        "B41 =ZÄHLENWENN(F4:F33;\"offen\"); B42 =SUMMEWENN(F4:F33;\"<>storniert\";E4:E33); B43 =SUMME(G4:G33); Filiale: Anzahl =ZÄHLENWENN(B;A47), "
        "Umsatz =SUMMEWENN(B;A47;G), Anteil =C47/$C$51; Warengruppe analog mit Spalte C. B62 =C51-B43; B63 =C59-B43.", loesung)
    ws = wb.create_sheet("Bestellungen")
    titel_zeile(ws, "BüroComfort GmbH – Bestellauswertung Quartal", 7)
    tabelle_kopf(ws, 3, ["Bestell-Nr.", "Filiale", "Warengruppe", "Zahlungsart", "Bestellwert", "Status", "Umsatz ohne Storno"], 1)
    daten_zellen(ws, 4, ZM_DATEN, [None, None, None, None, EURO, None])
    for r in range(4, 34):
        ziel(ws, f"G{r}", f'=IF(F{r}="storniert",0,E{r})', EURO, loesung)
    tabelle_kopf(ws, 36, ["Kennzahl", "Wert"], 1)
    kz = [("Anzahl aller Bestellungen", "=COUNTA(A4:A33)"), ("Bestellungen mit Bestellwert", "=COUNT(E4:E33)"),
          ("Bestellungen ohne Bestellwert", "=B37-B38"), ("Stornierte Bestellungen", '=COUNTIF(F4:F33,"storniert")'),
          ("Offene Bestellungen", '=COUNTIF(F4:F33,"offen")'),
          ("Umsatz ohne Storno (SUMMEWENN)", '=SUMIF(F4:F33,"<>storniert",E4:E33)'), ("Umsatz ohne Storno (Summe Spalte G)", "=SUM(G4:G33)")]
    for i, (t, f) in enumerate(kz):
        r = 37 + i
        beschriftung(ws, f"A{r}", t)
        ziel(ws, f"B{r}", f, EURO if r >= 42 else None, loesung)
    tabelle_kopf(ws, 46, ["Filiale", "Anzahl", "Umsatz ohne Storno", "Anteil"], 1)
    for i, fi in enumerate(FILIALEN):
        r = 47 + i
        beschriftung(ws, f"A{r}", fi)
        ziel(ws, f"B{r}", f"=COUNTIF($B$4:$B$33,A{r})", None, loesung)
        ziel(ws, f"C{r}", f"=SUMIF($B$4:$B$33,A{r},$G$4:$G$33)", EURO, loesung)
        ziel(ws, f"D{r}", f"=C{r}/$C$51", PROZ, loesung)
    beschriftung(ws, "A51", "Gesamt", True)
    ziel(ws, "B51", "=SUM(B47:B50)", None, loesung, True)
    ziel(ws, "C51", "=SUM(C47:C50)", EURO, loesung, True)
    ziel(ws, "D51", "=SUM(D47:D50)", PROZ, loesung, True)
    tabelle_kopf(ws, 54, ["Warengruppe", "Anzahl", "Umsatz ohne Storno", "Anteil"], 1)
    for i, w in enumerate(WG):
        r = 55 + i
        beschriftung(ws, f"A{r}", w)
        ziel(ws, f"B{r}", f"=COUNTIF($C$4:$C$33,A{r})", None, loesung)
        ziel(ws, f"C{r}", f"=SUMIF($C$4:$C$33,A{r},$G$4:$G$33)", EURO, loesung)
        ziel(ws, f"D{r}", f"=C{r}/$C$59", PROZ, loesung)
    beschriftung(ws, "A59", "Gesamt", True)
    ziel(ws, "B59", "=SUM(B55:B58)", None, loesung, True)
    ziel(ws, "C59", "=SUM(C55:C58)", EURO, loesung, True)
    ziel(ws, "D59", "=SUM(D55:D58)", PROZ, loesung, True)
    tabelle_kopf(ws, 61, ["Kontrolle", "Differenz"], 1)
    beschriftung(ws, "A62", "Umsatz Filialen minus Umsatz ohne Storno")
    ziel(ws, "B62", "=C51-B43", EURO, loesung)
    beschriftung(ws, "A63", "Umsatz Warengruppen minus Umsatz ohne Storno")
    ziel(ws, "B63", "=C59-B43", EURO, loesung)
    breiten(ws, {"A": 40, "B": 14, "C": 18, "D": 14, "E": 14, "F": 12, "G": 20})
    fuss(ws, loesung)
    return speichern(wb, "zm-auswertung", loesung)


if __name__ == "__main__":
    for fn in (z1, z2, z3, zm):
        for l in (False, True):
            print(fn(l))
