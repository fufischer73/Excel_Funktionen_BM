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
    erreicht: Rabattsatz und Rabatt stimmen für alle Bestellungen, auch an den Grenzen 100,00 € und 1.000,00 €.
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
| Rabatt | 5 % ab 100,00 €, 10 % ab 500,00 €, 15 % ab 1.000,00 € |
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
- Die Grenzfälle sind richtig: B-5008 (1.000,00 €), B-5011 (300,00 €) und B-5013 (100,00 €).
- Die Kennzahlen stehen in der Kennzahlen-Tabelle ab Zeile 21.
- Ergebnisse sind Formeln. Es steht keine eingetippte Zahl in einer Ergebniszelle.
- Zahlenformate, Schrift und Überschrift entsprechen dem Corporate Design.
- Die Fußzeile enthält Name und Klasse (rechtsbündig).
- Die Datei ist als .xlsx gespeichert.
:::

:::tipp Wenn Sie nicht weiterkommen
Schlagen Sie in den Modulen nach: SUMME und MITTELWERT in [R1](modul:r1), MIN und MAX in [R2](modul:r2), WENN in [R3](modul:r3), verschachteltes WENN in [R4](modul:r4), UND und ODER in [R5](modul:r5) und [R6](modul:r6).
:::

## 5. Abgabe

:::abgabe
:::
