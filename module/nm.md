---
titel: "Meilenstein: Rechnung mit Artikelstamm erstellen"
minuten: 60
meilenstein: true
review: entwurf
lernziele:
  - Eine vollständige Rechnung mit SVERWEIS aus Stammdaten aufbauen.
  - Rabattstaffel, Rundung und Fälligkeit in einer Rechnung kombinieren.
  - Eine Rechnung kontrollieren und nach Vorgabe gestalten.
abgabe:
  titel: Meilenstein Stufe 3 abgeben
  text: Laden Sie Ihre fertige Rechnung als Excel-Datei (.xlsx) hoch. Speichern Sie auch aus OnlyOffice als .xlsx.
  dateitypen: .xlsx
  dateien: 1
raster:
  - kriterium: Stammdaten nachschlagen
    erreicht: Kundenname, Ort, Kundengruppe, Zahlungsziel, Bezeichnung und Nettopreis stammen per SVERWEIS aus den Stammdaten und stimmen. Die Matrizen sind fixiert.
    besonders_gut: Die Wahl von FALSCH und WAHR ist in einem Satz begründet.
  - kriterium: Rechnungsrechnung
    erreicht: Gesamtpreise, Summe netto, Rabattsatz, Rabatt, Nettobetrag, Umsatzsteuer und Rechnungsbetrag sind richtig. Preise und Steuer sind auf zwei Nachkommastellen gerundet.
    besonders_gut: Die Rechnung wurde mit einer zweiten Rechnung (zum Beispiel Taschenrechner) gegengeprüft.
  - kriterium: Datum und Frist
    erreicht: Das Fälligkeitsdatum ist richtig berechnet und als Datum formatiert.
    besonders_gut: Der Status wird mit WENN und HEUTE angezeigt.
  - kriterium: Formeln statt Werte
    erreicht: Alle Ergebnisse sind Formeln mit Zellbezügen. Es wurde kein Ergebnis eingetippt. Die Rechnung enthält keine Fehlermeldung.
    besonders_gut: Wenn Sie die Kunden-Nr. ändern, passt sich die Rechnung korrekt an.
  - kriterium: Gestaltung nach Corporate Design
    erreicht: Arial 10, Überschrift zentriert 12 fett, Prozent mit zwei, Zahlen ohne und Euro mit zwei Nachkommastellen, Datum als TT.MM.JJJJ, Fußzeile mit Name und Klasse.
    besonders_gut: Die Rechnung ist druckfertig und ohne Erklärung gut lesbar.
quellen: [q-ms-sverweis, q-ms-runden, q-ms-heute]
---
## 1. Ihr Auftrag

Die Abteilung Verwaltung der BüroComfort GmbH erstellt eine Rechnung für den Kunden **K-4106**. Vorgegeben sind nur Rechnungsnummer, Rechnungsdatum, Kunden-Nr., Artikelnummern und Mengen. Alles andere holen Sie aus den **Stammdaten**: dem Artikelstamm, dem Kundenstamm, der Rabattstaffel und den Zahlungszielen.

:::material material/nm-rechnung.xlsx Übungsdatei: Rechnung mit Stammdaten
Tabellenblätter „Rechnung“ und „Stammdaten“. Das Blatt „Aufgabenstellung“ enthält alle Aufträge und Regeln.
:::

Lesen Sie zuerst das Blatt „Aufgabenstellung“. Es nennt die Regeln:

| Regel | Inhalt |
| --- | --- |
| Rabatt | Rabattstaffel auf die **Nettosumme**: 5 % ab 100,00 €, 10 % ab 500,00 €, 15 % ab 1.000,00 € |
| Zahlungsziel | Aus der Kundengruppe, steht im Blatt „Stammdaten“ |
| Umsatzsteuer | 19 %, kaufmännisch auf Cent gerundet |
| Rundung | Gesamtpreis je Position und Rabatt auf zwei Nachkommastellen runden |

## 2. Was Sie abgeben

- Kundenangaben und Fälligkeitsdatum (B6 bis B10)
- Alle Positionen mit Bezeichnung, Einzelpreis und Gesamtpreis
- Rechnungsrechnung: Summe netto, Rabattsatz, Rabatt, Nettobetrag, Umsatzsteuer, Rechnungsbetrag
- Status der Rechnung (B11), als Zusatzaufgabe
- Formatierung nach dem Corporate Design der BüroComfort GmbH und eine Fußzeile mit Ihrem Namen und Ihrer Klasse

## 3. Bewertungsraster

Das Raster sehen Sie **vor** der Arbeit. So wissen Sie, worauf es ankommt.

:::raster
:::

## 4. Checkliste vor der Abgabe

:::checkliste Vor der Abgabe prüfen
- Kunde, Ort, Kundengruppe und Zahlungsziel stehen in B6 bis B9. Das Fälligkeitsdatum steht in B10 und hat das Format TT.MM.JJJJ.
- Alle acht Positionen haben Bezeichnung, Einzelpreis und Gesamtpreis.
- Alle SVERWEIS-Formeln haben eine feste Matrix. Bei Artikelnummern, Kunden-Nr. und Kundengruppe steht FALSCH, bei der Rabattstaffel WAHR.
- Summe netto, Rabatt und Umsatzsteuer sind gerundet.
- Es steht keine Fehlermeldung wie #NV in der Rechnung.
- Ergebnisse sind Formeln. Es steht keine eingetippte Zahl in einer Ergebniszelle.
- Zahlenformate, Schrift und Überschrift entsprechen dem Corporate Design.
- Die Fußzeile enthält Name und Klasse (rechtsbündig).
- Die Datei ist als .xlsx gespeichert.
:::

:::loesung Kontrolle: Erwartete Ergebnisse
Der Kunde ist die Testwerk Elektro GmbH in Eckernförde (Gruppe B, 20 Tage Zahlungsziel). Fällig ist die Rechnung am 04.05.2026. Die Summe netto beträgt 1.784,80 €. Der Rabatt von 15 % ist 267,72 €. Der Nettobetrag nach Rabatt ist 1.517,08 €. Die Umsatzsteuer beträgt 288,25 €. Der Rechnungsbetrag ist 1.805,33 €.
:::

:::tipp Wenn Sie nicht weiterkommen
Schlagen Sie in den Modulen nach: SVERWEIS mit genauer Suche in [N1](modul:n1), SVERWEIS mit Staffel in [N2](modul:n2), Runden in [N3](modul:n3) und Datum in [N5](modul:n5).
:::

## Zusatz für Schnelle (freiwillig)

Wer vor dem Ende der Zeit fertig ist, kann die folgende Erweiterung bearbeiten. Sie wird nicht für die Bewertung des Meilensteins herangezogen. Die Erweiterung finden Sie im Blatt „Erweiterung Skonto“ derselben Datei.

:::aufgabe plus 25 Erweiterung · Zweite Rechnung mit Skonto
Öffnen Sie das Blatt „Erweiterung Skonto“ und lösen Sie die Aufträge:

- Holen Sie für jede Position Bezeichnung (Spalte C) und Einzelpreis (Spalte D) aus dem Artikelstamm und berechnen Sie in Spalte E den Gesamtpreis (gerundet).
- Holen Sie in H16 die Kundengruppe des Kunden aus H15 und in H17 das Zahlungsziel dieser Gruppe. Berechnen Sie in H18 das Fälligkeitsdatum und in H19 das Ende der Skontofrist.
- Berechnen Sie in H20 bis H27 die Summe netto, den Rabattsatz (ungefähre Suche in der Staffel), den Rabatt, den Nettobetrag nach Rabatt, die Umsatzsteuer (19 %), den Rechnungsbetrag, den Skontobetrag und den Zahlbetrag bei Skonto.
- Formatieren Sie die Rechnung nach dem Corporate Design. Datumsangaben im Format TT.MM.JJJJ.

:::checkliste Selbstcheck
- Spalte C („Bezeichnung“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Spalte D („Einzelpreis netto“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Spalte E („Gesamtpreis netto“) enthält in jeder Zeile eine Formel und keine eingetippten Werte.
- Alle Ergebnisse sind Formeln mit Zellbezügen. Bezüge, die beim Kopieren nicht wandern sollen, sind mit Dollarzeichen festgesetzt.
- Der Kunde gehört zur Gruppe A und hat 30 Tage Zahlungsziel.
:::

:::loesung Kontrolle: Erwartete Ergebnisse
- Spalte C (C14 bis C20): „Schreibtisch höhenverstellbar“; „Locher“; „Kugelschreiber“; „Mauspad“; „Aktenschrank 2 Ordnerhöhen“; „Haftnotizen“; „Tacker“
- Spalte D (D14 bis D20): 489,90 €; 6,40 €; 0,89 €; 3,75 €; 219,00 €; 2,10 €; 7,90 €
- Spalte E (E14 bis E20): 979,80 €; 160,00 €; 275,90 €; 150,00 €; 219,00 €; 105,00 €; 118,50 €
- Kundengruppe (H16): „A“
- Zahlungsziel in Tagen (H17): 30
- Fällig am (H18): 20.05.2026
- Skontofrist bis (H19): 30.04.2026
- Summe netto (H20): 2.008,20 €
- Rabattsatz (H21): 15,00 %
- Rabatt (H22): 301,23 €
- Nettobetrag nach Rabatt (H23): 1.706,97 €
- Umsatzsteuer (H24): 324,32 €
- Rechnungsbetrag (H25): 2.031,29 €
- Skonto (H26): 40,63 €
- Zahlbetrag bei Skonto (H27): 1.990,66 €
:::

:::loesung Lösung
Spalte C (ab C14): `=SVERWEIS(A14;Stammdaten!$A$4:$D$15;2;FALSCH)`. Spalte D (ab D14): `=SVERWEIS(A14;Stammdaten!$A$4:$D$15;3;FALSCH)`. Spalte E (ab E14): `=RUNDEN(B14*D14;2)`. H16: `=SVERWEIS(H15;Stammdaten!$F$4:$I$11;4;FALSCH)`. H17: `=SVERWEIS(H16;Stammdaten!$N$4:$O$6;2;FALSCH)`. H18: `=H14+H17`. H19: `=H14+10`. H20: `=SUMME(E14:E20)`. H21: `=SVERWEIS(H20;Stammdaten!$K$4:$L$7;2;WAHR)`. H22: `=RUNDEN(H20*H21;2)`. H23: `=H20-H22`. H24: `=RUNDEN(H23*0,19;2)`. H25: `=H23+H24`. H26: `=RUNDEN(H25*0,02;2)`. H27: `=H25-H26`.
:::
:::


## 5. Abgabe

:::abgabe
:::
