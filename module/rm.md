---
titel: "Meilenstein: Bestellungen der KW 12 kalkulieren"
minuten: 60
meilenstein: true
review: entwurf
lernziele:
  - Summe, Durchschnitt, kleinsten und größten Wert mit Funktionen ermitteln.
  - Rabatt, Zusatzbonus und Versandart mit WENN, UND und ODER berechnen.
  - Eine komplette Kalkulation übersichtlich und nach Vorgabe gestalten.
abgabe:
  titel: Meilenstein Stufe 1 abgeben
  text: Laden Sie Ihre fertige Kalkulation als Excel-Datei (.xlsx) hoch. Speichern Sie auch aus OnlyOffice als .xlsx.
  dateitypen: .xlsx
  dateien: 1
raster:
  - kriterium: Kennzahlen
    erreicht: Summe, Durchschnitt, größter und kleinster Bestellwert sowie die Summen der Rabatte und Zusatzboni sind richtig berechnet.
    besonders_gut: Die Kennzahlen sind zusätzlich durch eine zweite Rechnung gegengeprüft, zum Beispiel Summe der Zahlbeträge plus Rabatte plus Boni.
  - kriterium: WENN-Logik
    erreicht: Rabattsatz und Rabatt stimmen für alle Bestellungen, auch an den Grenzen 100,00 € und 500,00 €.
    besonders_gut: Die Formel ist übersichtlich aufgebaut, die Grenzfälle sind kurz erklärt.
  - kriterium: UND und ODER
    erreicht: Zusatzbonus und Versandart stimmen für alle Bestellungen, auch bei genau 300,00 €.
    besonders_gut: Die Regeln sind in einem Satz neben der Tabelle verständlich beschrieben.
  - kriterium: Formeln statt Werte
    erreicht: Alle Ergebnisse sind Formeln mit Zellbezügen. Es wurde kein Ergebnis eingetippt.
    besonders_gut: Die Formeln lassen sich ohne Änderung nach unten kopieren.
  - kriterium: Gestaltung nach Corporate Design
    erreicht: Arial 10, Überschrift zentriert 12 fett, Prozent mit zwei, Zahlen ohne und Euro mit zwei Nachkommastellen, Fußzeile mit Name und Klasse.
    besonders_gut: Die Tabelle ist druckfertig und ohne Erklärung gut lesbar.
quellen: [q-ms-wenn, q-ms-und, q-oder]
---
## 1. Ihr Auftrag

Sie arbeiten in der Abteilung Verwaltung der BüroComfort GmbH. Die Buchhaltung benötigt für die Kalenderwoche 12 eine vollständige Kalkulation aller Bestellungen.

:::material material/rm-kalkulation.xlsx Übungsdatei: Bestellungen KW 12
Tabellenblatt „Bestellungen KW 12“ mit 15 Bestellungen. Das Blatt „Aufgabenstellung“ enthält alle Regeln und Aufträge.
:::

Lesen Sie zuerst das Blatt „Aufgabenstellung“. Es nennt die Regeln:

| Regel | Inhalt |
| --- | --- |
| Rabatt | 5 % ab 100,00 €, 10 % ab 500,00 € |
| Zusatzbonus | 2 % des Bestellwerts für Bonuskunden ab 300,00 € Bestellwert |
| Versandart | Deutschland oder Österreich = „Standard“, sonst „Ausland“ |

## 2. Was Sie abgeben

- Rabattsatz und Rabatt in Euro für jede Bestellung
- Zusatzbonus in Euro
- Versandart
- Zahlbetrag (Bestellwert minus Rabatt minus Zusatzbonus)
- Kennzahlen: Summe der Bestellwerte, Durchschnitt, größter und kleinster Bestellwert, Summe der Rabatte, Summe der Zusatzboni
- Formatierung nach dem Corporate Design der BüroComfort GmbH und eine Fußzeile mit Ihrem Namen und Ihrer Klasse

## 3. Bewertungsraster

Das Raster sehen Sie **vor** der Arbeit. So wissen Sie, worauf es ankommt.

:::raster
:::

## 4. Checkliste vor der Abgabe

:::checkliste Vor der Abgabe prüfen
- Alle 15 Zeilen haben Rabattsatz, Rabatt, Zusatzbonus, Versandart und Zahlbetrag.
- Die Grenzfälle sind richtig: B-5015 (500,00 €), B-5011 (300,00 €) und B-5013 (100,00 €).
- Die Kennzahlen stehen in der Kennzahlen-Tabelle ab Zeile 21.
- Ergebnisse sind Formeln. Es steht keine eingetippte Zahl in einer Ergebniszelle.
- Zahlenformate, Schrift und Überschrift entsprechen dem Corporate Design.
- Die Fußzeile enthält Name und Klasse (rechtsbündig).
- Die Datei ist als .xlsx gespeichert.
:::

:::tipp Wenn Sie nicht weiterkommen
Schlagen Sie in den Modulen nach: SUMME und MITTELWERT in [R1](modul:r1), MIN und MAX in [R2](modul:r2), WENN in [R3](modul:r3), verschachteltes WENN in [R4](modul:r4), UND und ODER in [R5](modul:r5) und [R6](modul:r6).
:::

## Zusatz für Schnelle (freiwillig)

Wer vor dem Ende der Zeit fertig ist, kann die folgende Erweiterung bearbeiten. Sie wird nicht für die Bewertung des Meilensteins herangezogen. Die Erweiterung finden Sie im Blatt „Erweiterung KW 13“ derselben Datei.

:::aufgabe plus 20 Erweiterung · Kalkulation mit Express-Versand
Öffnen Sie das Blatt „Erweiterung KW 13“ und lösen Sie die Aufträge:

- Ermitteln Sie in Spalte F den Rabattsatz (5 % ab 100,00 €, 10 % ab 500,00 €) und in Spalte G den Rabatt in Euro.
- Berechnen Sie in Spalte H den Zusatzbonus (2 % des Bestellwerts für Bonuskunden ab 300,00 €).
- Legen Sie in Spalte I die Versandart fest (Express, Standard oder Ausland).
- Berechnen Sie in Spalte J den Zahlbetrag.
- Ermitteln Sie in den Zellen B28 bis B31 die Summe der Zahlbeträge, den Durchschnitt, den größten Zahlbetrag und die Anzahl der Expresssendungen.
- Formatieren Sie Ihre Ergebnisse nach dem Corporate Design.

:::checkliste Selbstcheck
- Spalte F („Rabattsatz“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Spalte G („Rabatt in €“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Spalte H („Zusatzbonus in €“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Spalte I („Versandart“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Spalte J („Zahlbetrag“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Alle Ergebnisse sind Formeln mit Zellbezügen.
- Grenzfälle: B-5106 (500,00 €, Bonuskunde, DE) wird Express. B-5109 (300,00 €) erhält den Zusatzbonus. B-5108 (299,99 €) nicht.
:::

:::loesung Kontrolle: Erwartete Ergebnisse
- Spalte F (F14 bis F25): 10,00 %; 0,00 %; 10,00 %; 5,00 %; 5,00 %; 10,00 %; 10,00 %; 5,00 %; 5,00 %; 10,00 %; 5,00 %; 10,00 %
- Spalte G (G14 bis G25): 62,00 €; 0,00 €; 125,00 €; 24,00 €; 17,00 €; 50,00 €; 100,00 €; 15,00 €; 15,00 €; 76,04 €; 5,00 €; 189,00 €
- Spalte H (H14 bis H25): 12,40 €; 0,00 €; 25,00 €; 0,00 €; 6,80 €; 10,00 €; 0,00 €; 0,00 €; 6,00 €; 0,00 €; 0,00 €; 37,80 €
- Spalte I (I14 bis I25): „Express“; „Standard“; „Express“; „Ausland“; „Ausland“; „Express“; „Standard“; „Standard“; „Standard“; „Standard“; „Standard“; „Ausland“
- Spalte J (J14 bis J25): 545,60 €; 95,50 €; 1.100,00 €; 456,00 €; 316,20 €; 440,00 €; 900,00 €; 284,99 €; 279,00 €; 684,36 €; 95,00 €; 1.663,20 €
- Summe Zahlbeträge (B28): 6.859,85 €
- Durchschnitt Zahlbetrag (B29): 571,65 €
- Größter Zahlbetrag (B30): 1.663,20 €
- Expresssendungen (B31): 3
:::

:::loesung Lösung
Spalte F (ab F14): `=WENN(E14>=500;0,1;WENN(E14>=100;0,05;0))`. Spalte G (ab G14): `=E14*F14`. Spalte H (ab H14): `=WENN(UND(C14="Ja";E14>=300);E14*0,02;0)`. Spalte I (ab I14): `=WENN(UND(D14="DE";C14="Ja";E14>=500);"Express";WENN(ODER(D14="DE";D14="AT");"Standard";"Ausland"))`. Spalte J (ab J14): `=E14-G14-H14`. B28: `=SUMME(J14:J25)`. B29: `=MITTELWERT(J14:J25)`. B30: `=MAX(J14:J25)`. B31: `=ZÄHLENWENN(I14:I25;"Express")`.
:::
:::


## 5. Abgabe

:::abgabe
:::
