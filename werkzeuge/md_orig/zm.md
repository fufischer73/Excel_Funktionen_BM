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

## 5. Abgabe

:::abgabe
:::
