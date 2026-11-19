#!/usr/bin/env python3
"""Erzeugt die Grundlagenmodule G1 (Formeln und Kopieren) und G2 (Feste Bezüge):
Übungsdateien (material/), Musterlösungen (lehrkraft/loesungen/) und die Aufgabenblöcke der Modultexte.

Aufruf aus dem Kursordner:  python3 werkzeuge/gen_grundlagen.py
Benötigt openpyxl und LibreOffice (soffice) zum Neuberechnen der Lösungsdateien.
Die erklärenden Textteile der Module stehen in werkzeuge/md_orig/g1.md und g2.md (Platzhalter {{AUFGABEN}}).
"""
import os, shutil, subprocess, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ext_framework import *          # setzt Stile, build_sheet, md_block ...
from gen_stufe1 import aufgabenblatt, fuss
from openpyxl import Workbook, load_workbook

KURS = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
ORIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "md_orig")

# ------------------------------------------------------------------ G1
G1 = [
dict(name="Übung 1 Rechnen", niveau="basis", min=20, kurz="Gesamtpreise berechnen",
 titel="Lernsituation: Gesamtpreise einer Sammelbestellung",
 sit="Die Einkaufsabteilung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, erfasst eine Sammelbestellung mit zwölf Positionen. Für jede Position soll der Gesamtpreis berechnet werden. Je Position kommt eine Verpackungspauschale von 2,50 € hinzu.",
 auftr=["Berechnen Sie in Spalte D den Gesamtpreis je Position (Menge mal Einzelpreis). Geben Sie die Formel nur in D14 ein und kopieren Sie sie mit dem Ausfüllkästchen nach unten.",
        "Berechnen Sie in Spalte E den Gesamtpreis mit der Verpackungspauschale. Tippen Sie die 2,50 € direkt in die Formel.",
        "Lassen Sie sich die Formeln anzeigen. Tragen Sie in B27 ein, welche Formel in E20 steht.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."],
 cols=[("Artikel", 28, None, "d", ""), ("Menge", 10, ZAHL, "d", ""), ("Einzelpreis", 14, EURO, "d", ""),
       ("Gesamtpreis", 16, EURO, "f", "=B{r}*C{r}"), ("mit Verpackung", 16, EURO, "f", "=D{r}+2.5")],
 rows=[("Kugelschreiber, 50er Pack", 20, 18.9), ("Kopierpapier, 5 x 500 Blatt", 15, 21.5), ("Ordner, 10er Pack", 8, 19.4),
       ("Haftnotizen, 12er Pack", 30, 9.95), ("Tacker", 6, 7.6), ("Heftklammern, Schachtel", 40, 1.85), ("Toner schwarz", 4, 64.9),
       ("Bürostuhl Standard", 3, 129.0), ("Schreibtischlampe LED", 5, 34.9), ("Aktenvernichter", 1, 89.0), ("Locher", 12, 6.45),
       ("Briefumschläge, 500er Pack", 10, 14.2)],
 extras=[("label", "A27", "Formel in E20", True),
         ("answer", "B27", "Die Formel lautet =D20+2,5. Die Bezüge wandern beim Kopieren mit, die 2,5 bleibt unverändert.", None, "B27:E27")],
 checks=["Ändern Sie die Menge in B14. D14 und E14 ändern sich sofort.", "Die Formel in D14 wurde nur einmal eingegeben und dann kopiert."],
 hilfe=[("Formel aufbauen", "Eine Formel beginnt immer mit =. Klicken Sie danach auf die erste Zelle, tippen Sie das Rechenzeichen und klicken Sie auf die zweite Zelle. Bestätigen Sie mit ++Enter++."),
        ("Kopieren", "Klicken Sie D14 an. Ziehen Sie das kleine Quadrat unten rechts an der Zelle (Ausfüllkästchen) bis D25 nach unten.")]),
dict(name="Übung 2 Punkt vor Strich", niveau="basis", min=15, kurz="Rechnung mit Rabatt und Steuer",
 titel="Lernsituation: Rechnungspositionen mit Rabatt und Umsatzsteuer",
 sit="Die Verwaltung der BüroComfort GmbH erstellt eine Rechnung über Büromöbel. Jede Position hat einen eigenen Rabatt. Zu jedem Nettobetrag kommen 19 % Umsatzsteuer.",
 auftr=["Berechnen Sie in Spalte E den Nettobetrag nach Rabatt: Menge mal Einzelpreis mal (1 minus Rabatt).",
        "Berechnen Sie in Spalte F die Umsatzsteuer mit 19 %. Tippen Sie den Steuersatz als 0,19 in die Formel.",
        "Berechnen Sie in Spalte G den Bruttobetrag.",
        "Erklären Sie in B25, warum die Formel =B14*C14*1-D14 ein falsches Ergebnis liefert.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."],
 cols=[("Artikel", 28, None, "d", ""), ("Menge", 10, ZAHL, "d", ""), ("Einzelpreis", 14, EURO, "d", ""), ("Rabatt", 10, PROZ, "d", ""),
       ("Nettobetrag", 16, EURO, "f", "=B{r}*C{r}*(1-D{r})"), ("Umsatzsteuer", 16, EURO, "f", "=E{r}*0.19"), ("Bruttobetrag", 16, EURO, "f", "=E{r}+F{r}")],
 rows=[("Bürostuhl Standard", 4, 129.0, 0.10), ("Schreibtisch höhenverstellbar", 2, 389.0, 0.05), ("Rollcontainer", 3, 119.9, 0.0),
       ("Konferenztisch", 1, 649.0, 0.15), ("Besucherstuhl", 6, 74.5, 0.10), ("Aktenschrank", 2, 259.0, 0.05), ("Garderobenständer", 3, 45.9, 0.0),
       ("Pinnwand", 5, 38.5, 0.10), ("Whiteboard", 2, 89.0, 0.05), ("Beistelltisch", 4, 52.0, 0.0)],
 extras=[("label", "A25", "Begründung", True),
         ("answer", "B25", "Excel rechnet Punkt vor Strich. In der falschen Formel wird zuerst B14*C14*1 berechnet und danach der Rabattsatz D14 (eine Zahl wie 0,1) abgezogen. Der Rabatt muss als (1-D14) in Klammern stehen.", None, "B25:G25")],
 checks=["Bei Rollcontainer, Garderobenständer und Beistelltisch (Rabatt 0 %) ist der Nettobetrag gleich Menge mal Einzelpreis."],
 hilfe=[("Klammern", "Was zuerst berechnet werden soll, steht in Klammern. Der Rabatt wird nicht abgezogen, sondern der Rest (1 minus Rabatt) wird mit Menge und Preis multipliziert.")]),
dict(name="Übung 3 Zur Seite kopieren", niveau="basis", min=15, kurz="Monats- und Halbjahreswerte",
 titel="Lernsituation: Umsätze im ersten Halbjahr",
 sit="Die Geschäftsleitung der BüroComfort GmbH möchte für das erste Halbjahr sehen, wie hoch der Umsatz je Warengruppe und je Monat ist.",
 auftr=["Berechnen Sie in Spalte H das Halbjahr je Warengruppe. Geben Sie die Formel nur in H14 ein und kopieren Sie sie nach unten.",
        "Berechnen Sie in Zeile 18 den Gesamtumsatz je Monat. Geben Sie die Formel nur in B18 ein und kopieren Sie sie nach rechts bis H18.",
        "Kontrollieren Sie das Ergebnis: Addieren Sie in B20 die Monatswerte B18 bis G18 und berechnen Sie in B21 die Differenz zu H18. Die Differenz muss 0 sein.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."],
 cols=[("Warengruppe", 18, None, "d", ""), ("Januar", 13, EURO, "d", ""), ("Februar", 13, EURO, "d", ""), ("März", 13, EURO, "d", ""),
       ("April", 13, EURO, "d", ""), ("Mai", 13, EURO, "d", ""), ("Juni", 13, EURO, "d", ""), ("Halbjahr", 15, EURO, "f", "=B{r}+C{r}+D{r}+E{r}+F{r}+G{r}")],
 rows=[("Schreibwaren", 4120, 3890, 4470, 4015, 4380, 3960), ("Büromöbel", 6840, 7210, 6525, 5980, 7655, 6210),
       ("Papier", 2950, 3105, 3340, 2875, 3190, 2815), ("Technik", 9380, 8760, 10120, 9215, 11040, 9870)],
 extras=[("label", "A18", "Gesamt", True)] +
        [("target", f"{c}18", f"={c}14+{c}15+{c}16+{c}17", EURO, True, f"Gesamt {n}") for c, n in zip("BCDEFGH", ["Januar", "Februar", "März", "April", "Mai", "Juni", "Halbjahr"])] +
        [("label", "A20", "Kontrolle: Summe B18 bis G18", True), ("target", "B20", "=B18+C18+D18+E18+F18+G18", EURO, False, "Summe der Monatswerte"),
         ("label", "A21", "Differenz zu H18", True), ("target", "B21", "=H18-B20", EURO, False, "Differenz")],
 widths={"A": 30},
 checks=["Die Formel in Zeile 18 wurde nur einmal eingegeben und nach rechts kopiert."],
 hilfe=[("Nach rechts kopieren", "Ziehen Sie das Ausfüllkästchen von B18 nach rechts bis H18. Aus B14+B15+B16+B17 wird dabei C14+C15+C16+C17 und so weiter.")]),
dict(name="Übung 4 Fehlersuche", niveau="plus", min=15, kurz="Fehlerhafte Formeln finden",
 titel="Lernsituation: Fehlerhafte Bestellliste prüfen",
 sit="Ein Kollege hat die Bestellliste der BüroComfort GmbH vorbereitet. Beim Kopieren der Formeln in Spalte D sind drei Fehler passiert. Die Ergebnisse sehen auf den ersten Blick richtig aus, stimmen aber nicht.",
 auftr=["Lassen Sie sich die Formeln der Spalte D anzeigen und suchen Sie die drei fehlerhaften Zellen.",
        "Korrigieren Sie die Zellen so, dass in der ganzen Spalte dieselbe Formel gilt.",
        "Tragen Sie in B25 die Zeilennummern der fehlerhaften Zellen ein und in B26 die Fehlerart.",
        "Berechnen Sie in Spalte E den Preis abzüglich 5 % Rabatt (Gesamtpreis mal 0,95)."],
 cols=[("Artikel", 28, None, "d", ""), ("Menge", 10, ZAHL, "d", ""), ("Einzelpreis", 14, EURO, "d", ""),
       ("Gesamtpreis", 16, EURO, "x", "=B{r}*C{r}",
        lambda r: 90 if r == 16 else (f"=B{r-1}*C{r}" if r == 19 else (f"=B{r}*C{r+1}" if r == 22 else f"=B{r}*C{r}")),
        "In D16 steht ein eingetippter Wert (90) statt der Formel. In D19 zeigt die Formel auf die Menge der Zeile darüber (B18), in D22 auf den Preis der Zeile darunter (C23)."),
       ("nach 5 % Rabatt", 16, EURO, "f", "=D{r}*0.95")],
 rows=[("Kugelschreiber, 50er Pack", 20, 18.9), ("Kopierpapier, 5 x 500 Blatt", 15, 21.5), ("Ordner, 10er Pack", 6, 15.5), ("Haftnotizen, 12er Pack", 30, 9.95),
       ("Tacker", 6, 7.6), ("Heftklammern, Schachtel", 40, 1.85), ("Toner schwarz", 4, 64.9), ("Locher", 12, 6.45), ("Briefumschläge, 500er Pack", 10, 14.2),
       ("Etiketten, 100er Pack", 25, 4.3)],
 extras=[("label", "A25", "Zeilen mit Fehlern", True), ("answer", "B25", "16, 19 und 22", None, "B25:E25"),
         ("label", "A26", "Fehlerarten", True),
         ("answer", "B26", "Zeile 16: eingetippter Wert statt Formel. Zeile 19: Bezug auf die Zeile darüber. Zeile 22: Bezug auf die Zeile darunter.", None, "B26:E26")],
 hilfe=[("Fehler finden", "Zeigen Sie die Formeln an. In einer richtigen Spalte steht in jeder Zeile dieselbe Formel mit der passenden Zeilennummer. Suchen Sie die Zeilen, die davon abweichen.")]),
]

# ------------------------------------------------------------------ G2
G2 = [
dict(name="Übung 1 Ein fester Wert", niveau="basis", min=20, kurz="Umsatzsteuer mit festem Satz",
 titel="Lernsituation: Bruttopreise mit Umsatzsteuersatz aus einer Zelle",
 sit="Die BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, führt eine Preisliste. Der Umsatzsteuersatz steht in Zelle H15. Ändert er sich, sollen sich alle Preise ohne weitere Eingabe anpassen.",
 auftr=["Berechnen Sie in C14 die Umsatzsteuer mit dem Satz aus H15 und kopieren Sie die Formel nach unten. Verwenden Sie zuerst den Bezug H15 ohne Dollarzeichen. Beschreiben Sie in B26, was beim Kopieren passiert.",
        "Setzen Sie den Bezug auf H15 mit der Taste F4 fest und kopieren Sie die Formel erneut nach unten.",
        "Berechnen Sie in Spalte D den Bruttopreis.",
        "Ändern Sie den Steuersatz in H15 testweise auf 7 %. Setzen Sie ihn danach wieder auf 19 %.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."],
 cols=[("Artikel", 28, None, "d", ""), ("Nettopreis", 14, EURO, "d", ""), ("Umsatzsteuer", 16, EURO, "f", "=B{r}*$H$15"), ("Bruttopreis", 16, EURO, "f", "=B{r}+C{r}")],
 rows=[("Kugelschreiber, 50er Pack", 18.9), ("Kopierpapier, 5 x 500 Blatt", 21.5), ("Ordner, 10er Pack", 19.4), ("Haftnotizen, 12er Pack", 9.95),
       ("Tacker", 7.6), ("Toner schwarz", 64.9), ("Bürostuhl Standard", 129.0), ("Schreibtischlampe LED", 34.9), ("Aktenvernichter", 89.0), ("Locher", 6.45)],
 extras=[("label", "H14", "Umsatzsteuersatz", True), ("input", "H15", 0.19, PROZ),
         ("label", "A26", "Beobachtung ohne Dollarzeichen", True),
         ("answer", "B26", "Der Bezug wandert mit: In C15 steht H16, in C16 steht H17 und so weiter. Diese Zellen sind leer, deshalb ist die Umsatzsteuer ab C15 gleich 0.", None, "B26:G26")],
 widths={"H": 20},
 checks=["Bei 7 % in H15 ändern sich nur die Spalten C und D."],
 hilfe=[("Dollarzeichen setzen", "Klicken Sie in der Bearbeitungsleiste in den Bezug H15 und drücken Sie ++F4++. Aus H15 wird $H$15. Bei manchen Laptops brauchen Sie dazu zusätzlich die Fn-Taste."),
        ("Formel", "Die Umsatzsteuer ist der Nettopreis mal der Steuersatz. Der Steuersatz steht immer in derselben Zelle.")]),
dict(name="Übung 2 Anteile", niveau="basis", min=15, kurz="Anteile am Gesamtumsatz",
 titel="Lernsituation: Anteile der Warengruppen am Gesamtumsatz",
 sit="Die Geschäftsleitung der BüroComfort GmbH möchte wissen, welchen Anteil jede Warengruppe am Gesamtumsatz hat und wie weit sie vom durchschnittlichen Umsatz abweicht.",
 auftr=["Berechnen Sie in B20 den Gesamtumsatz und in B21 den Mittelwert der Warengruppen.",
        "Berechnen Sie in Spalte C den Anteil jeder Warengruppe am Gesamtumsatz. Setzen Sie den Bezug auf B20 fest.",
        "Berechnen Sie in Spalte D die Abweichung vom Mittelwert. Setzen Sie den Bezug auf B21 fest.",
        "Berechnen Sie in C20 die Summe der Anteile. Das Ergebnis muss 100 % sein.",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."],
 cols=[("Warengruppe", 22, None, "d", ""), ("Umsatz", 16, EURO, "d", ""), ("Anteil", 12, PROZ, "f", "=B{r}/$B$20"), ("Abweichung vom Mittelwert", 20, EURO, "f", "=B{r}-$B$21")],
 rows=[("Büromöbel", 48250), ("Technik", 61940), ("Papier", 18730), ("Schreibwaren", 22410), ("Reinigung", 9860), ("Verpackung", 12390)],
 extras=[("label", "A20", "Gesamt", True), ("target", "B20", "=SUM(B14:B19)", EURO, True, "Gesamtumsatz"),
         ("target", "C20", "=SUM(C14:C19)", PROZ, True, "Summe der Anteile"),
         ("label", "A21", "Mittelwert", True), ("target", "B21", "=AVERAGE(B14:B19)", EURO, False, "Mittelwert")],
 checks=["Die Summe der Anteile in C20 ist 100,00 %."],
 hilfe=[("Anteil", "Anteil = Teil geteilt durch Ganzes. Der Teil (Spalte B) wandert beim Kopieren mit, das Ganze (B20) bleibt fest.")]),
dict(name="Übung 3 Zwei feste Werte", niveau="basis", min=15, kurz="Verkaufspreise kalkulieren",
 titel="Lernsituation: Verkaufspreise kalkulieren",
 sit="Der Einkauf der BüroComfort GmbH kalkuliert die Verkaufspreise. Auf den Einkaufspreis kommen Bezugskosten, auf den Bezugspreis ein Gewinnzuschlag. Beide Sätze stehen in der Tabelle und können sich ändern.",
 auftr=["Berechnen Sie in Spalte C den Bezugspreis: Einkaufspreis mal (1 plus Bezugskostensatz aus H15).",
        "Berechnen Sie in Spalte D den Verkaufspreis netto: Bezugspreis mal (1 plus Gewinnzuschlag aus H18).",
        "Berechnen Sie in Spalte E den Gewinn je Stück.",
        "Ändern Sie den Gewinnzuschlag in H18 auf 30 %. Beantworten Sie in B24: Warum ändert sich Spalte C nicht?",
        "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."],
 cols=[("Artikel", 26, None, "d", ""), ("Einkaufspreis", 15, EURO, "d", ""), ("Bezugspreis", 16, EURO, "f", "=B{r}*(1+$H$15)"),
       ("Verkaufspreis netto", 18, EURO, "f", "=C{r}*(1+$H$18)"), ("Gewinn je Stück", 16, EURO, "f", "=D{r}-C{r}")],
 rows=[("Bürostuhl Standard", 71.5), ("Schreibtisch höhenverstellbar", 214.0), ("Rollcontainer", 66.8), ("Besucherstuhl", 41.2),
       ("Aktenschrank", 142.6), ("Pinnwand", 21.3), ("Whiteboard", 49.0), ("Beistelltisch", 28.9)],
 extras=[("label", "H14", "Bezugskostensatz", True), ("input", "H15", 0.04, PROZ), ("label", "H17", "Gewinnzuschlag", True), ("input", "H18", 0.25, PROZ),
         ("label", "A24", "Begründung", True),
         ("answer", "B24", "Spalte C hängt nur vom Einkaufspreis und vom Bezugskostensatz in H15 ab. Der Gewinnzuschlag in H18 wird erst in Spalte D verwendet.", None, "B24:G24")],
 widths={"H": 20},
 hilfe=[("Zwei feste Zellen", "Jede Formel braucht einen eigenen festen Bezug: Spalte C den Satz aus H15, Spalte D den Satz aus H18.")]),
dict(name="Übung 4 Gemischte Bezüge", niveau="plus", min=20, kurz="Preistabelle mit Rabattstaffel",
 titel="Lernsituation: Preistabelle mit drei Rabattsätzen",
 sit="Der Vertrieb der BüroComfort GmbH möchte für acht Listenpreise sehen, wie sich ein Rabatt von 5 %, 10 % und 15 % auswirkt. Die Rabattsätze stehen in Zeile 12.",
 auftr=["Berechnen Sie in B14 den Preis nach 5 % Rabatt. Die Formel soll sich ohne Änderung nach rechts bis D14 und nach unten bis Zeile 21 kopieren lassen.",
        "Setzen Sie dazu bei der Formel den Bezug auf den Listenpreis so fest, dass nur die Spalte A feststeht. Den Bezug auf den Rabattsatz setzen Sie so fest, dass nur die Zeile 12 feststeht.",
        "Tragen Sie in B23 ein, wie die Formel in B14 lautet, und in B24, was $A14 bedeutet.",
        "Ändern Sie den Satz in C12 auf 12 %. Prüfen Sie, dass sich nur die Spalte C ändert."],
 cols=[("Listenpreis", 18, EURO, "d", ""), ("Preis bei Rabatt 1", 18, EURO, "f", "=$A{r}*(1-B$12)"),
       ("Preis bei Rabatt 2", 18, EURO, "f", "=$A{r}*(1-C$12)"), ("Preis bei Rabatt 3", 18, EURO, "f", "=$A{r}*(1-D$12)")],
 rows=[(25.0,), (50.0,), (75.0,), (100.0,), (150.0,), (200.0,), (250.0,), (500.0,)],
 extras=[("label", "A12", "Rabattsatz", True), ("input", "B12", 0.05, PROZ), ("input", "C12", 0.10, PROZ), ("input", "D12", 0.15, PROZ),
         ("label", "A23", "Formel in B14", True), ("answer", "B23", "Die Formel lautet =$A14*(1-B$12).", None, "B23:D23"),
         ("label", "A24", "Bedeutung von $A14", True), ("answer", "B24", "Die Spalte A ist festgesetzt, die Zeilennummer darf beim Kopieren nach unten wandern.", None, "B24:D24")],
 hilfe=[("Gemischte Bezüge", "Drücken Sie ++F4++ mehrmals hintereinander. Der Bezug wechselt: A14, $A$14, A$14, $A14. Wählen Sie die Variante, bei der nur ein Teil festgesetzt ist."),
        ("Welcher Teil bleibt fest?", "Beim Kopieren nach rechts bleibt die Spalte A fest. Beim Kopieren nach unten bleibt die Zeile 12 fest.")]),
dict(name="Übung 5 Fehlersuche", niveau="plus", min=15, kurz="Fehlende Dollarzeichen finden",
 titel="Lernsituation: Anteile und Provisionen prüfen",
 sit="Ein Kollege hat in der Auswertung der BüroComfort GmbH die Anteile und die Provisionen berechnet. Beim Kopieren nach unten stimmen die Ergebnisse nicht mehr: Es erscheinen Fehlermeldungen und Nullen.",
 auftr=["Lassen Sie sich die Formeln der Spalten C und D anzeigen. Finden Sie heraus, welche Bezüge beim Kopieren gewandert sind.",
        "Korrigieren Sie beide Spalten, sodass in jeder Zeile der richtige feste Bezug verwendet wird.",
        "Tragen Sie in B27 die Ursache des Fehlers ein."],
 cols=[("Warengruppe", 22, None, "d", ""), ("Umsatz", 16, EURO, "d", ""),
       ("Anteil", 12, PROZ, "x", "=B{r}/$B$24", lambda r: f"=B{r}/B{24 + (r - 14)}", "Der Bezug auf den Gesamtumsatz B24 ist nicht festgesetzt und wandert beim Kopieren nach unten (B25, B26 ...). Ab Zeile 15 wird durch leere Zellen geteilt (#DIV/0!)."),
       ("Provision", 14, EURO, "x", "=B{r}*$H$15", lambda r: f"=B{r}*H{15 + (r - 14)}", "Der Bezug auf den Provisionssatz H15 ist nicht festgesetzt und wandert nach unten (H16, H17 ...). Ab Zeile 15 wird mit leeren Zellen multipliziert, das Ergebnis ist 0.")],
 rows=[("Büromöbel", 48250), ("Technik", 61940), ("Papier", 18730), ("Schreibwaren", 22410), ("Reinigung", 9860), ("Verpackung", 12390),
       ("Software", 33200), ("Zubehör", 15840), ("Beleuchtung", 11975), ("Sonstiges", 7420)],
 extras=[("label", "A24", "Gesamt", True), ("target", "B24", "=SUM(B14:B23)", EURO, True, "Gesamtumsatz"),
         ("label", "H14", "Provisionssatz", True), ("input", "H15", 0.03, PROZ),
         ("label", "A27", "Ursache", True),
         ("answer", "B27", "In beiden Spalten fehlen die Dollarzeichen. Der Bezug auf B24 bzw. H15 wandert beim Kopieren nach unten mit und zeigt dann auf leere Zellen.", None, "B27:G27")],
 widths={"H": 18},
 hilfe=[("Fehler finden", "Vergleichen Sie die Formeln in C14 und C15. Welcher Teil der Formel hat sich verändert, obwohl er gleich bleiben sollte?")]),
]

DATEIEN = {
    "g1": ("g1-formeln-kopieren", "Formeln und Kopieren", G1),
    "g2": ("g2-feste-bezuege", "Feste Bezüge", G2),
}

AUFGABENSTELLUNG = {
    "g1": dict(titel="Lernsituation: Rechnen mit Zellbezügen und Formeln kopieren",
               sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. In diesem Modul lernen Sie, Formeln mit Zellbezügen aufzubauen und mit dem Ausfüllkästchen zu kopieren. Funktionen wie SUMME folgen erst im nächsten Modul.",
               hinweis="Jedes Tabellenblatt dieser Datei enthält eine eigene Ausgangssituation, Aufträge und eine Tabelle. Die Schwierigkeit steigt von Blatt zu Blatt."),
    "g2": dict(titel="Lernsituation: Feste Bezüge in Formeln",
               sit="Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg. In diesem Modul lernen Sie, Zellbezüge beim Kopieren festzuhalten, damit Sätze, Summen und Staffelwerte nicht mitwandern.",
               hinweis="Jedes Tabellenblatt dieser Datei enthält eine eigene Ausgangssituation, Aufträge und eine Tabelle. Die Schwierigkeit steigt von Blatt zu Blatt."),
}


def recalc(pfad):
    """Lösungsdatei mit LibreOffice neu berechnen (gespeicherte Werte)."""
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, pfad],
                   check=True, capture_output=True, timeout=180)
    shutil.copy(os.path.join(tmp, os.path.basename(pfad)), pfad)


def bauen(key):
    stem, titel, specs = DATEIEN[key]
    info = AUFGABENSTELLUNG[key]
    for loes in (False, True):
        wb = Workbook()
        aufgabenblatt(wb, info["titel"], info["sit"],
                      [f"Bearbeiten Sie das Blatt „{s['name']}“." for s in specs], info["hinweis"],
                      "Lösungsdatei mit Formeln und Kontrollwerten. Fußzeilen-Platzhalter ersetzen.", loes)
        for s in specs:
            build_sheet(wb, s, loes, len(wb.sheetnames))
        ordner = os.path.join(KURS, "lehrkraft/loesungen" if loes else "material")
        os.makedirs(ordner, exist_ok=True)
        pfad = os.path.join(ordner, f"{stem}{'-loesung' if loes else ''}.xlsx")
        wb.save(pfad)
        if loes:
            recalc(pfad)
    vals = load_workbook(os.path.join(KURS, "lehrkraft/loesungen", f"{stem}-loesung.xlsx"), data_only=True)
    bad = [(ws.title, c.coordinate, c.value) for ws in vals for row in ws.iter_rows() for c in row
           if isinstance(c.value, str) and c.value.startswith("#")]
    # Fehlerzellen sind nur in den Fehlersuche-Blättern der Aufgabe erwartet, nicht in der Lösung
    print(key, "Fehlerzellen in der Lösung:", bad)
    blocks = [md_block(s, i, vals[s["name"]], None) for i, s in enumerate(specs, 1)]
    t = open(os.path.join(ORIG, f"{key}.md"), encoding="utf-8").read()
    t = t.replace("{{AUFGABEN}}", "\n".join(blocks))
    open(os.path.join(KURS, "module", f"{key}.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    for k in DATEIEN:
        bauen(k)
