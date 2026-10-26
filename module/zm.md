---
titel: "Meilenstein: Bestellauswertung erstellen"
minuten: 65
meilenstein: true
review: entwurf
lernziele:
  - Bestellungen mit ANZAHL, ANZAHL2, ZÄHLENWENN und SUMMEWENN auswerten.
  - Stornierte und unvollständige Daten berücksichtigen.
  - Ergebnisse mit einer Gegenprobe kontrollieren.
abgabe:
  titel: Meilenstein Stufe 2 abgeben
  text: Laden Sie Ihre fertige Auswertung als Excel-Datei (.xlsx) hoch. Speichern Sie auch aus OnlyOffice als .xlsx.
  dateitypen: .xlsx
  dateien: 1
raster:
  - kriterium: Kennzahlen
    erreicht: Anzahl aller Bestellungen, Bestellungen mit und ohne Bestellwert, stornierte und offene Bestellungen sind richtig ermittelt.
    besonders_gut: Die Unterschiede zwischen ANZAHL und ANZAHL2 sind in einem Satz neben der Tabelle erklärt.
  - kriterium: Umsatz ohne Storno
    erreicht: Die Spalte G und der Umsatz ohne Storno stimmen, und beide Wege (SUMMEWENN und Summe der Spalte G) liefern dasselbe Ergebnis.
    besonders_gut: Die Abweichung durch die nicht bewertete Bestellung ist in einem Satz erklärt.
  - kriterium: Auswertung je Filiale und Warengruppe
    erreicht: Anzahl, Umsatz und Anteil sind für alle Filialen und Warengruppen richtig. Die Bereiche sind mit Dollarzeichen fixiert.
    besonders_gut: Die Formeln lassen sich ohne Änderung kopieren.
  - kriterium: Gegenprobe
    erreicht: Die Differenzen in B62 und B63 sind 0,00 €. Die Gesamtzahlen in den Auswertungen stimmen mit 30 Bestellungen überein.
    besonders_gut: Mögliche Fehlerquellen sind kurz benannt.
  - kriterium: Gestaltung nach Corporate Design
    erreicht: Arial 10, Überschrift zentriert 12 fett, Prozent mit zwei, Zahlen ohne und Euro mit zwei Nachkommastellen, Fußzeile mit Name und Klasse.
    besonders_gut: Die Tabelle ist druckfertig und ohne Erklärung gut lesbar.
quellen: [q-ms-anzahl, q-ms-anzahl2, q-ms-zaehlenwenn, q-ms-summewenn]
---
## 1. Ihr Auftrag

Die Geschäftsleitung der BüroComfort GmbH bereitet die Quartalsbesprechung vor. Sie erhalten die Liste aller 30 Bestellungen und sollen sie auswerten: **Wie viele Bestellungen gab es? Wie viel Umsatz ist angefallen? Wie verteilen sich Anzahl und Umsatz auf die Filialen und die Warengruppen?**

Zwei Besonderheiten gehören zur Aufgabe:

- **Stornierte Bestellungen** gehören nicht zum Umsatz.
- Bei **einer Bestellung fehlt der Bestellwert**. Sie ist noch nicht bewertet und darf Ihre Ergebnisse nicht verfälschen.

:::material material/zm-auswertung.xlsx Übungsdatei: Bestellauswertung Quartal
Tabellenblatt „Bestellungen“ mit 30 Bestellungen und den Auswertungstabellen darunter. Das Blatt „Aufgabenstellung“ enthält alle Aufträge.
:::

Lesen Sie zuerst das Blatt „Aufgabenstellung“.

## 2. Was Sie abgeben

- Spalte G: Umsatz ohne Storno für jede Bestellung
- Kennzahlen (B37 bis B43), einschließlich des Umsatzes ohne Storno auf zwei Wegen
- Auswertung je Filiale: Anzahl, Umsatz ohne Storno und Anteil
- Auswertung je Warengruppe: Anzahl, Umsatz ohne Storno und Anteil
- Gegenprobe in B62 und B63
- Formatierung nach dem Corporate Design der BüroComfort GmbH und eine Fußzeile mit Ihrem Namen und Ihrer Klasse

## 3. Bewertungsraster

Das Raster sehen Sie **vor** der Arbeit. So wissen Sie, worauf es ankommt.

:::raster
:::

## 4. Checkliste vor der Abgabe

:::checkliste Vor der Abgabe prüfen
- Alle 30 Zeilen haben in Spalte G ein Ergebnis. Bei stornierten Bestellungen steht 0,00 €.
- Es gibt 30 Bestellungen, davon 29 mit Bestellwert, 3 storniert und 4 offen.
- Die Umsatzsummen der Filialen und der Warengruppen sind gleich und stimmen mit dem Umsatz ohne Storno überein.
- Die Anteile ergeben jeweils 100 %.
- Ergebnisse sind Formeln. Es steht keine eingetippte Zahl in einer Ergebniszelle.
- Zahlenformate, Schrift und Überschrift entsprechen dem Corporate Design.
- Die Fußzeile enthält Name und Klasse (rechtsbündig).
- Die Datei ist als .xlsx gespeichert.
:::

:::tipp Wenn Sie nicht weiterkommen
Schlagen Sie in den Modulen nach: ANZAHL und ANZAHL2 in [Z1](modul:z1), ZÄHLENWENN in [Z2](modul:z2) und SUMMEWENN in [Z3](modul:z3). Für WENN hilft [R3](modul:r3).
:::

:::tipp Zur Gegenprobe
Zeigen die Zellen B62 und B63 nicht 0,00 €, vergleichen Sie die Filial- und Warengruppennamen mit der Liste. Ein Tippfehler im Namen oder ein überzähliges Leerzeichen genügt, damit eine Zeile fehlt.
:::

## Zusatz für Schnelle (freiwillig)

Wer vor dem Ende der Zeit fertig ist, kann die folgende Erweiterung bearbeiten. Sie wird nicht für die Bewertung des Meilensteins herangezogen. Die Erweiterung finden Sie im Blatt „Erweiterung Zahlungsarten“ derselben Datei.

:::aufgabe plus 25 Erweiterung · Zahlungsarten und Storno-Quote
Öffnen Sie das Blatt „Erweiterung Zahlungsarten“ und lösen Sie die Aufträge:

- Ermitteln Sie in Spalte F den Umsatz ohne Storno (bei stornierten Bestellungen 0) und in Spalte G eine Hilfsspalte: 1 bei stornierten Bestellungen, sonst 0.
- Werten Sie in I14 bis I17 für jede Filiale die Anzahl der Bestellungen, in J14 bis J17 den Umsatz ohne Storno, in K14 bis K17 die Storno-Quote (Anteil stornierter Bestellungen) und in L14 bis L17 den Umsatzrang aus.
- Werten Sie in I21 bis I23 für jede Zahlungsart die Anzahl, in J21 bis J23 den Umsatz ohne Storno und in K21 bis K23 den Anteil am Gesamtumsatz aus.
- Prüfen Sie in J24, ob die Umsätze der Zahlungsarten zusammen den Umsatz der Filialen ergeben (Text „stimmt“ oder „stimmt nicht“).
- Formatieren Sie Ihre Ergebnisse nach dem Corporate Design.

:::checkliste Selbstcheck
- Spalte F („Umsatz ohne Storno“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Spalte G („Storno (1/0)“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Alle Ergebnisse sind Formeln mit Zellbezügen.
- Eine Bestellung hat keinen Bestellwert (leere Zelle). Sie zählt bei der Anzahl mit, bringt aber 0 € Umsatz.
:::

:::loesung Kontrolle: Erwartete Ergebnisse
- Spalte F (F14 bis F37): 781,45 €; 734,61 €; 692,03 €; 1.474,00 €; 0,00 €; 317,70 €; 516,39 €; 0,00 €; 164,08 €; 1.493,71 €; 1.146,67 €; 862,19 €; 604,31 €; 0,00 €; 0,00 €; 959,00 €; 167,23 €; 236,51 €; 718,81 €; 782,94 €; 0,00 €; 40,69 €; 0,00 €; 1.484,38 €
- Spalte G (G14 bis G37): 0; 0; 0; 0; 1; 0; 0; 0; 0; 0; 0; 0; 0; 1; 1; 0; 0; 0; 0; 0; 1; 0; 1; 0
- Anzahl Rendsburg (I14): 10
- Anzahl Kiel (I15): 6
- Anzahl Flensburg (I16): 3
- Anzahl Neumünster (I17): 5
- Umsatz Rendsburg (J14): 7.100,86 €
- Umsatz Kiel (J15): 2.209,58 €
- Umsatz Flensburg (J16): 1.313,90 €
- Umsatz Neumünster (J17): 2.552,36 €
- Storno-Quote Rendsburg (K14): 10,00 %
- Storno-Quote Kiel (K15): 33,33 %
- Storno-Quote Flensburg (K16): 33,33 %
- Storno-Quote Neumünster (K17): 20,00 %
- Rang Rendsburg (L14): 1
- Rang Kiel (L15): 3
- Rang Flensburg (L16): 4
- Rang Neumünster (L17): 2
- Anzahl Rechnung (I21): 9
- Anzahl Lastschrift (I22): 8
- Anzahl Vorkasse (I23): 7
- Umsatz Rechnung (J21): 4.352,31 €
- Umsatz Lastschrift (J22): 3.998,99 €
- Umsatz Vorkasse (J23): 4.825,40 €
- Anteil Rechnung (K21): 33,03 %
- Anteil Lastschrift (K22): 30,35 %
- Anteil Vorkasse (K23): 36,62 %
- Kontrolle (J24): „stimmt“
:::

:::loesung Lösung
Spalte F (ab F14): `=WENN(D14="storniert";0;E14)`. Spalte G (ab G14): `=WENN(D14="storniert";1;0)`. I14: `=ZÄHLENWENN($B$14:$B$37;H14)`. I15: `=ZÄHLENWENN($B$14:$B$37;H15)`. I16: `=ZÄHLENWENN($B$14:$B$37;H16)`. I17: `=ZÄHLENWENN($B$14:$B$37;H17)`. J14: `=SUMMEWENN($B$14:$B$37;H14;$F$14:$F$37)`. J15: `=SUMMEWENN($B$14:$B$37;H15;$F$14:$F$37)`. J16: `=SUMMEWENN($B$14:$B$37;H16;$F$14:$F$37)`. J17: `=SUMMEWENN($B$14:$B$37;H17;$F$14:$F$37)`. K14: `=SUMMEWENN($B$14:$B$37;H14;$G$14:$G$37)/I14`. K15: `=SUMMEWENN($B$14:$B$37;H15;$G$14:$G$37)/I15`. K16: `=SUMMEWENN($B$14:$B$37;H16;$G$14:$G$37)/I16`. K17: `=SUMMEWENN($B$14:$B$37;H17;$G$14:$G$37)/I17`. L14: `=RANG(J14;$J$14:$J$17;0)`. L15: `=RANG(J15;$J$14:$J$17;0)`. L16: `=RANG(J16;$J$14:$J$17;0)`. L17: `=RANG(J17;$J$14:$J$17;0)`. I21: `=ZÄHLENWENN($C$14:$C$37;H21)`. I22: `=ZÄHLENWENN($C$14:$C$37;H22)`. I23: `=ZÄHLENWENN($C$14:$C$37;H23)`. J21: `=SUMMEWENN($C$14:$C$37;H21;$F$14:$F$37)`. J22: `=SUMMEWENN($C$14:$C$37;H22;$F$14:$F$37)`. J23: `=SUMMEWENN($C$14:$C$37;H23;$F$14:$F$37)`. K21: `=J21/SUMME($F$14:$F$37)`. K22: `=J22/SUMME($F$14:$F$37)`. K23: `=J23/SUMME($F$14:$F$37)`. J24: `=WENN(ABS(SUMME(J21:J23)-SUMME(J14:J17))<0,005;"stimmt";"stimmt nicht")`.
:::
:::


## 5. Abgabe

:::abgabe
:::
