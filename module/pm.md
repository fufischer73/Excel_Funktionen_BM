---
titel: "Meilenstein: Prüfungsnahe Gesamtaufgabe lösen"
minuten: 55
meilenstein: true
review: entwurf
lernziele:
  - Alle Funktionen des Lernpfads in einer umfangreichen Aufgabe kombinieren.
  - Stammdaten, Rechnung, Fristen und Auswertung in einer Datei verbinden.
  - Die eigene Arbeit selbst kontrollieren und Prüfungsanforderungen erfüllen.
abgabe:
  titel: Meilenstein Stufe 4 abgeben
  text: Laden Sie Ihre fertige Datei als Excel-Datei (.xlsx) hoch. Speichern Sie auch aus OnlyOffice als .xlsx.
  dateitypen: .xlsx
  dateien: 1
raster:
  - kriterium: Stammdaten nachschlagen
    erreicht: Kundenname, Kundengruppe, Artikelbezeichnung und Einzelpreis stammen per SVERWEIS mit FALSCH aus den Stammdaten und stimmen. Die Bereiche sind fixiert.
    besonders_gut: Die Wahl von FALSCH und WAHR ist in einem Satz begründet.
  - kriterium: Rechnung und Rundung
    erreicht: Nettowert, Rabattsatz (WAHR), Rabatt, Nettobetrag, Umsatzsteuer und Bruttobetrag sind für alle 24 Bestellungen richtig und kaufmännisch gerundet.
    besonders_gut: Die Grenzfälle der Rabattstaffel sind geprüft und kurz benannt.
  - kriterium: Fälligkeit und Status
    erreicht: Fälligkeitsdatum und Status sind für alle Bestellungen richtig. Das Zahlungsziel stammt aus der Kundengruppe.
    besonders_gut: Der Status ändert sich korrekt, wenn Sie den Stichtag ändern.
  - kriterium: Auswertung
    erreicht: Anzahl, Umsatz, Anteil und Rang je Kundengruppe sowie alle Kennzahlen sind richtig. Die Kontrolle in der Auswertung zeigt 0,00 €.
    besonders_gut: Der Antwortsatz ist vollständig und nennt Kundengruppe und Anteil.
  - kriterium: Formeln statt Werte
    erreicht: Alle Ergebnisse sind Formeln mit Zellbezügen. Es wurde kein Ergebnis eingetippt. Die Datei enthält keine Fehlermeldung.
    besonders_gut: Die Formeln lassen sich ohne Änderung nach unten kopieren.
  - kriterium: Gestaltung nach Corporate Design
    erreicht: Arial 10, Überschrift zentriert 12 fett, Prozent mit zwei, Zahlen ohne und Euro mit zwei Nachkommastellen, Datum als TT.MM.JJJJ, Fußzeile mit Name und Klasse auf beiden Blättern.
    besonders_gut: Die Tabellen sind druckfertig und ohne Erklärung gut lesbar.
quellen: [q-ms-sverweis, q-ms-runden, q-ms-zaehlenwenn, q-ms-summewenn, q-ms-rang, q-ms-wenn]
---
## 1. Ihr Auftrag

Dies ist die Abschlussaufgabe des Lernpfads. Sie ähnelt in Aufbau und Umfang einer Prüfungsaufgabe: mehrere Blätter, viele Aufträge mit Punkten, keine Schritt-für-Schritt-Anleitung. Sie brauchen fast alle Funktionen, die Sie kennengelernt haben.

Die Abteilung Verwaltung der BüroComfort GmbH wertet die 24 Bestellungen des zweiten Quartals aus. Sie berechnen für jede Bestellung Nettowert, Rabatt, Umsatzsteuer und Bruttobetrag, prüfen, welche Bestellungen zum Stichtag überfällig sind, und erstellen eine Auswertung je Kundengruppe.

:::material material/pm-gesamtaufgabe.xlsx Übungsdatei: Quartalsauswertung der Bestellungen
Tabellenblätter „Bestellungen“, „Auswertung“ und „Stammdaten“. Das Blatt „Aufgabenstellung“ enthält Situation, Regeln und alle Aufträge mit Punkten.
:::

Lesen Sie zuerst das Blatt „Aufgabenstellung“. Bearbeitungszeit: etwa 50 Minuten.

| Auftrag | Inhalt | Punkte |
| --- | --- | --- |
| 1 | Stammdaten holen (Kunde, Gruppe, Artikel, Preis) | 12 |
| 2 | Nettowert, Rabatt, Umsatzsteuer, Bruttobetrag | 18 |
| 3 | Fälligkeit und Status zum Stichtag | 10 |
| 4 | Auswertung je Kundengruppe und Kennzahlen | 16 |
| 5 | Kontrolle und Antwortsatz | 6 |
| 6 | Gestaltung nach Corporate Design, Fußzeile | 8 |
| | **Gesamt** | **70** |

:::tipp Zeitplan
Planen Sie Ihre Zeit: Etwa 10 Minuten für das Lesen und Planen, 30 Minuten für die Bestellungen, 10 Minuten für die Auswertung, 10 Minuten für Kontrolle und Gestaltung. Arbeiten Sie die Aufträge der Reihe nach ab. Kommen Sie bei einem Auftrag nicht weiter, gehen Sie zum nächsten und kehren später zurück.
:::

## 2. Bewertungsraster

Das Raster sehen Sie **vor** der Arbeit. So wissen Sie, worauf es ankommt.

:::raster
:::

## 3. Checkliste vor der Abgabe

:::checkliste Vor der Abgabe prüfen
- Alle 24 Zeilen haben in den Spalten F bis Q Ergebnisse. Es steht keine Fehlermeldung wie #NV in der Datei.
- Alle SVERWEIS haben feste Bereiche. Bei Nummern steht FALSCH, bei der Rabattstaffel WAHR.
- Nettowert, Rabatt und Umsatzsteuer sind mit RUNDEN auf Cent gerundet.
- Die Auswertung enthält Anzahl, Umsatz, Anteil und Rang je Kundengruppe und alle sieben Kennzahlen.
- Die Kontrollzelle B27 zeigt 0,00 €. Die Anteile ergeben 100 %.
- Der Antwortsatz in A29 ist vollständig.
- Alle Ergebnisse sind Formeln. Es steht keine eingetippte Zahl in einer Ergebniszelle.
- Zahlenformate, Schrift und Überschriften stimmen auf beiden Blättern mit dem Corporate Design überein. Die Fußzeile enthält Name und Klasse.
- Die Datei ist als .xlsx gespeichert.
:::

:::loesung Hilfe: Vorgehen
Erledigen Sie zuerst die Stammdaten in den Spalten F bis I. Alle weiteren Spalten rechnen mit diesen Ergebnissen. Für die Fälligkeit brauchen Sie die Kundengruppe: Das Zahlungsziel steht im Blatt „Stammdaten“ in der Tabelle der Zahlungsziele. Die Auswertung auf dem zweiten Blatt greift auf das Blatt „Bestellungen“ zu: Schreiben Sie vor den Bereich den Blattnamen, zum Beispiel `Bestellungen!$G$5:$G$28`.
:::

:::loesung Kontrolle: Erwartete Ergebnisse
Gesamt: 24 Bestellungen mit einer Summe der Bruttobeträge von 12.786,51 €. Durchschnitt: 532,77 €. Größter Bruttobetrag: 2.477,66 €. Kleinster: 13,77 €. Überfällig zum Stichtag 15.06.2026 sind 17 Bestellungen mit 9.142,62 €. Kundengruppen: A 8 Bestellungen, 6.503,07 € (50,86 %, Rang 1). B 10 Bestellungen, 4.313,56 € (33,74 %, Rang 2). C 6 Bestellungen, 1.969,88 € (15,41 %, Rang 3). Die Kontrolle zeigt 0,00 €. Beispiel: B-1101 (Testwerk Elektro GmbH, Gruppe B, 40 Ordner) hat einen Nettowert von 99,60 € (kein Rabatt, knapp unter 100,00 €), 18,92 € Umsatzsteuer, 118,52 € Brutto und ist am 23.04.2026 fällig, also überfällig.
:::

:::loesung Lösung
F5: `=SVERWEIS(C5;Stammdaten!$F$4:$I$11;2;FALSCH)`. G5: wie F5 mit Spaltenindex 4. H5: `=SVERWEIS(D5;Stammdaten!$A$4:$D$15;2;FALSCH)`. I5: wie H5 mit Spaltenindex 3. J5: `=RUNDEN(E5*I5;2)`. K5: `=SVERWEIS(J5;Stammdaten!$K$4:$L$7;2;WAHR)`. L5: `=RUNDEN(J5*K5;2)`. M5: `=J5-L5`. N5: `=RUNDEN(M5*$E$2;2)`. O5: `=M5+N5`. P5: `=B5+SVERWEIS(G5;Stammdaten!$N$4:$O$6;2;FALSCH)`. Q5: `=WENN(P5<$B$2;"überfällig";"offen")`. Auswertung: B4: `=ZÄHLENWENN(Bestellungen!$G$5:$G$28;A4)`. C4: `=SUMMEWENN(Bestellungen!$G$5:$G$28;A4;Bestellungen!$O$5:$O$28)`. D4: `=C4/$C$7`. E4: `=RANG(C4;$C$4:$C$6;0)`. B11: `=ANZAHL2(Bestellungen!A5:A28)`. B12: `=SUMME(Bestellungen!O5:O28)`. B13: `=MITTELWERT(Bestellungen!O5:O28)`. B14: `=MAX(Bestellungen!O5:O28)`. B15: `=MIN(Bestellungen!O5:O28)`. B16: `=ZÄHLENWENN(Bestellungen!Q5:Q28;"überfällig")`. B17: `=SUMMEWENN(Bestellungen!Q5:Q28;"überfällig";Bestellungen!O5:O28)`. B27: `=RUNDEN(C7-B12;2)`.
:::

## 4. Abgabe

:::abgabe
:::
