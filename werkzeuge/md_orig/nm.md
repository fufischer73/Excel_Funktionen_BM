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

## 5. Abgabe

:::abgabe
:::
