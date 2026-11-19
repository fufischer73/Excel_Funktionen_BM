# Übergabe: Lernpfad „Excel-Funktionen“

Stand: 25.10.2026 · Auftraggeber: Jan, BBZ Rendsburg-Eckernförde · Werkzeug: Lernpfad-Baukasten (GitHub `dclausen01/lernpfad-baukasten`)

## 1. Auftrag und Rahmen

- **Thema:** Abschnitt „Funktionen“ im Bereich Tabellenkalkulation der Kann-Liste (Lernportfolio v3).
- **Zielgruppe:** Kaufleute für Büromanagement. Vorkenntnis: Zellen formatieren und Tabellen nach Vorlage gestalten (das ist ein eigener, früherer Lernpfad).
- **Umfang:** Einheiten à 90 Minuten, keine Hausaufgaben. Geplant waren 10 Doppelstunden, die Zeitansätze sollten bewusst großzügig sein.
- **Niveaus:** Basis (alle), Plus und Kreativ (Schnelle). Zusätzlich mehrere Aufgabenblätter je Datei mit steigender Schwierigkeit (siehe 4).
- **Modellunternehmen:** BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, Abteilung Verwaltung. Kundennamen sind fiktiv.
- **Ansprache:** Sie-Form.

## 2. Vorgaben

### Software und Formate
- **Microsoft Excel** ist die Hauptvariante, weil die Prüfung der Büromanagement-Auszubildenden damit gestaltet ist.
- **ONLYOFFICE** ist das künftige Werkzeug der Schule. Alle Übungsdateien sind `.xlsx` und damit in ONLYOFFICE zu öffnen. Die Module sind als Variante Excel/OnlyOffice angelegt (Umschalter im Kurs).
- Formeln stehen in den Dateien mit englischen Funktionsnamen und Komma (Speicherformat). Excel und ONLYOFFICE zeigen sie in der deutschen Oberfläche deutsch an (z. B. `=WENN(A1>5;"ja";"nein")`). In den Modultexten stehen die deutschen Namen mit Semikolon.
- Wichtig für ONLYOFFICE: Die Dateien enthalten keine gespeicherten Berechnungswerte in den Aufgabendateien. Die Programme rechnen beim Öffnen. Menünamen und das Verhalten von Blattbezügen in ONLYOFFICE sind noch nicht getestet (siehe 6).
- **Corporate Design der BüroComfort GmbH** (in jeder Aufgabe genannt): Arial 10, Überschriften zentriert 12 fett, Prozent 2 Nachkommastellen, andere Zahlen 0 Nachkommastellen, Euro `#,##0.00 "€"`. Zielzellen in den Aufgabendateien sind unformatiert („Standard“). Die Fußzeile (Name, Klasse, rechtsbündig) steht nur in den Lösungsdateien als Platzhalter.

### Moodle
- Moodle Version 4, Freischaltung der Stufen **nacheinander**, Einstufung über Meilensteine, Bestehensgrenze Einstufungstest 80 %.
- Aus dem Kurs entsteht über den Baukasten ein Moodle-Paket (Skript `alles.py`, Übergabe nach `moodle.md`). Das ist noch **nicht** gebaut, weil erst die Freigabe nötig ist.
- Pro Modul gibt es Kurz-Check (Moodle-Test), bei Bedarf Abgabe (Moodle-Aufgabe mit Punkten) und bei Meilensteinen ein Bewertungsraster.
- Die Lösungsdateien (`lehrkraft/loesungen`) werden **nicht** in Web oder Moodle ausgeliefert.

### Inhaltliche Regeln
- Nichts erfinden: Quellen werden belegt, Unsicheres steht als `zu_pruefen`.
- Keine Aufgaben der Kammer (Nüra-Prüfungsdateien) nachbauen. Sie dienen nur als Orientierung für Tabellenarten. Quellen von easy4me.info dürfen mit Quellenangabe genutzt oder neu gestaltet werden.
- Keine personenbezogenen Daten.
- Review-Haltepunkt: Nur Jan setzt `review: freigegeben`. Alle 21 Module stehen auf `entwurf`.

## 3. Struktur des Lernpfads (21 Module, 4 Stufen)

| Stufe | Module (Minuten Pflicht) |
| --- | --- |
| 1 Rechnen und Entscheiden | R1 SUMME/MITTELWERT, R2 MIN/MAX, R3 WENN, R4 WENN verschachtelt, R5 UND/ODER, R6 UND/ODER mit WENN, RM Meilenstein „Bestellungen KW 12 kalkulieren“ |
| 2 Zählen und Summieren | Z1 ANZAHL/ANZAHL2, Z2 ZÄHLENWENN, Z3 SUMMEWENN, ZM Meilenstein „Bestellauswertung“ |
| 3 Nachschlagen | N1 SVERWEIS, N2 SVERWEIS im Einsatz, N3 Runden, N4 RANG, N5 Datum und Zeit, NM Meilenstein „Rechnung mit Artikelstamm“ |
| 4 Prüfungsfit | P1 Angebot und Rechnung, P2 Auswertung, P3 Zeiten und Fristen, PM Meilenstein „prüfungsnahe Gesamtaufgabe“ |

Dazu: Selbstcheck am Anfang, Einstufungstest je Stufe.

**Zeitbilanz (Pflichtzeit laut Modulkopf): 1.094 Minuten, etwa 12,2 Doppelstunden** (vorher 905 Minuten). Plus- und Kreativaufgaben sowie die Meilenstein-Erweiterungen kommen freiwillig hinzu.

## 4. Aufbau der Übungsdateien

Jede Datei in `material/` hat das Blatt „Aufgabenstellung“, das ursprüngliche Aufgabenblatt und (Stufen 1 bis 3) zusätzliche Blätter:

| Blatt | Niveau | Inhalt |
| --- | --- | --- |
| Aufgabe 1 (bisheriges Blatt) | Basis | Einstieg, Grenzfälle |
| Aufgabe 2 Standard | Basis (Pflicht) | neue Situation, Kontrollzellen |
| Aufgabe 3 Anspruchsvoll | Plus | Funktionen kombiniert, feste Eingabezellen, Begründungsfragen |
| Aufgabe 4 Transfer | Plus | wenig Anleitung. Bei N1 mit fehlerhaften Formeln zur Fehlersuche |

Abweichungen: R1, R2, R5, N3, N4 haben nur zwei Zusatzblätter. N5 hat „Aufgabe 6 Anspruchsvoll“ und „Aufgabe 7 Transfer“. RM, ZM, NM haben je ein freiwilliges Blatt „Erweiterung …“. Stufe 4 blieb unverändert (bereits mehrblättrig).

Die Modultexte verweisen auf die Blattnamen und enthalten je Aufgabe Selbstcheck, „Kontrolle: Erwartete Ergebnisse“ und „Lösung“ (aufklappbar).

## 5. Erstellte Dateien

Kursordner `excel-funktionen/` (Cloud-Arbeitskopie; Kopie als ZIP in OneDrive):

- `kurs.yaml`, `medien.yaml` (Herstellerlinks, Linkprüfung 25.10.2026), `quellen.yaml`
- `module/*.md`: 21 Module
- `material/*.xlsx`: 21 Übungsdateien (r1 bis r6, rm, z1 bis z3, zm, n1 bis n5, nm, p1 bis p3, pm)
- `lehrkraft/loesungen/*-loesung.xlsx`: 21 Musterlösungen (mit Formeln und Fußzeilen-Platzhalter, per Neuberechnung geprüft, 0 Fehler)
- `werkzeuge/`: Generatoren der Übungsdateien und Modul-Aufgabenblöcke mit README
- `ausgabe/web/`: gebaute Web-Version (nie von Hand ändern)
- `ausgabe/pruefung/`: Prüfbericht und Screenshots (nicht im ZIP)

**Ablage auf Jans Rechner (OneDrive):**
`C:\Users\fufis\OneDrive\007_Schule\02_Fächer\02_Bürokaufleute\02_Excel\000_Excel_Grundlagenseminar\Lernpfad_Excel-Funktionen\`
- `lernpfad_update_v2.zip` ist der aktuelle Gesamtstand (143 Dateien). Bitte selbst entpacken, dann `ausgabe\web\index.html` öffnen.
- `lernpfad_update.zip` ist der ältere Stand vor den Zusatzblättern. Daneben liegen frühere Einzeldateien der Web-Version (Stufe-1-Stand).

Werkzeuge: Skripte im Baukasten (`pruefen.py`, `bauen.py`, `durchklicken.py`, später `alles.py`). Letzter Lauf: Prüfung ohne Fehler (6 Hinweise zu Quellen), Web gebaut, Durchklick ohne Probleme.

## 6. Offene Punkte

1. **Quellen prüfen:** 6 Einträge in `quellen.yaml` stehen auf `zu_pruefen`: `q-autosumme-pfeil`, `q-mittelwert-null`, `q-oder`, `q-oo-deutsch`, `q-zeit-bruchteil`, `q-zinsformel`.
2. **ONLYOFFICE testen:** Übungsdateien und Formeln in der Schulinstallation öffnen. Menünamen, Semikolon als Trenner und Blattbezüge (z. B. `Artikelstamm_3!$A$4:$E$13`) prüfen. Die OnlyOffice-Anleitung ist bisher nur in R1 ausgearbeitet. Hilfen und Lösungen der übrigen Module sind in Excel-Schreibweise.
3. **Zeitplan entscheiden:** Pflichtzeit 12,2 statt 10 Doppelstunden. Optionen: Standardblätter freiwillig machen oder Module kürzen. Danach in `kurs.yaml` das Badge „ca. 10 Doppelstunden“ anpassen (steht noch dort).
4. **Stufe 4** (P1 bis P3, PM) bei Bedarf im Niveau anheben (nicht beauftragt).
5. **Inhaltliche Durchsicht und Freigabe:** nur Jan setzt `review: freigegeben`. Danach Moodle-Build (`alles.py`, Version 4) und Übergabe nach `moodle.md`.
6. **Optional:** Dateinamen der Web-Seiten mit Umlauten über `datei:` setzen, easy4me-Material oder Videos mit Quellenangabe ergänzen, Selbstcheck im Kurs testen.
7. **Hinweis zu Meilenstein-Texten:** Das Bewertungsraster bezieht sich nur auf die Hauptaufgabe. Die Erweiterungsblätter gehen nicht in die Bewertung ein.

## 7. Hinweise für die Weiterarbeit

- Die Generatoren (Python, `openpyxl`) liegen im Kursordner unter `werkzeuge/`. Aufruf und Reihenfolge stehen in `werkzeuge/README.md`. Die Modultexte der Stufen 1 bis 3 werden aus `werkzeuge/md_orig/` neu geschrieben. Handänderungen an den Aufgabenblöcken gehen beim erneuten Lauf verloren.
- Nach jeder Änderung: `pruefen.py`, `bauen.py`, `durchklicken.py` laufen lassen. Lösungsdateien nach Änderungen neu berechnen (Recalc).
- Bekannte Besonderheit: TAGE360 setzt Enddaten am Monatsende um. Beispiele daher ohne Stichtag am Monatsletzten.

## 8. Ergänzung 19.11.2026: Grundlagenmodule und Kann-Liste

- **Neu:** Modul **G1 Formeln und Kopieren** (vor R1) und **G2 Feste Bezüge** (nach R2, vor R3) mit Übungsdateien `material/g1-formeln-kopieren.xlsx` und `material/g2-feste-bezuege.xlsx` (je 4 bzw. 5 Übungsblätter, davon eine Fehlersuche), Lösungen in `lehrkraft/loesungen/`. Erzeugt mit `werkzeuge/gen_grundlagen.py` (Textvorlagen: `werkzeuge/md_orig/g1.md`, `g2.md`; Aufgabenblöcke werden eingesetzt). Dateien sind mit LibreOffice berechnet und fehlerfrei. **Noch nicht in Excel getestet.**
- **kurs.yaml:** Stufe 1 lautet jetzt `[g1, r1, r2, g2, r3, ...]`. Der Selbstcheck empfiehlt Einsteigern G1.
- **Noch zu tun:** (1) Web-Build mit dem Baukasten (`pruefen.py`, `bauen.py`); die Seitennamen `g1-formeln-und-kopieren.html` und `g2-feste-bezuge.html` in `portfolio/kann-liste.html` sind vermutet und nach dem Build zu prüfen. (2) Neue Module haben keine `quellen:`/`medien:`-Einträge. (3) Die Plus-Blätter von R1 (Aufgabe 3) und R2 (Aufgabe 3) verlangen schon ein festes `$`; entweder Hinweis auf G2 ergänzen oder die Blätter anpassen. (4) Zeitbilanz: G1 und G2 bringen je 60 Minuten dazu (Pflichtzeit nun ca. 1.214 Minuten).
- **Kann-Liste (Prototyp):** `portfolio/kann-liste.html` verlinkt die Funktionen-Kompetenzen mit den Modulen und erzeugt den Reiter „Mein Lernplan“. Die sechs neuen Kompetenzen zu Formeln und Zellbezügen stehen bewusst am Ende des Bereichs Tabellenkalkulation, damit gespeicherte Ampeln (Schlüssel nach Position) gültig bleiben.
- **Beschlossen (Jan):** Verschachtelte WENN höchstens zwei WENN insgesamt (betrifft R4 Aufgabe 3/4, Rabattstaffel, R6 Aufgabe 4). Weitere Bereiche: Gestalten und Formatieren, Diagramme, Rechenoperationen. Prüfungsnahe Aufgaben mit Excel-Vorlage steuert Jan bei. Ziel-Software ist Microsoft Excel.
- **R4 und RM auf höchstens zwei WENN umgebaut (19.11.2026):** Rabattstaffel jetzt 0 % / 5 % ab 100 € / 10 % ab 500 €; Provision 0 % / 1 % ab 10.000 € / 2 % ab 30.000 €; Versandkosten nach Gewicht in drei Stufen (bis 5 kg, bis 10 kg, darüber; Sperrgut entfällt); Lieferantennoten A (ab 80), B (ab 60), C (darunter). RM hat dieselbe Staffel (B-5015 steht jetzt genau auf 500,00 €, Summe der Bestellwerte 7.581,15 €). Noch mit mehr als zwei WENN: R6 Aufgabe 4 (E und F), N4 Aufgabe 3 (Prämie), P3 (Mahnstatus). `ext_run.py` rechnet ohne die Skill-Umgebung mit LibreOffice neu.
- **Alle Formeln jetzt höchstens zwei WENN (19.11.2026):** R6 neu gegliedert in leicht (Aufgabe 2 UND, Aufgabe 3 ODER), Standard (4), Kombination (5), anspruchsvoll (6), Transfer (7, vereinfacht). N4 Aufgabe 3: Prämie Rang 1 = 100 €, Rang 2 und 3 = 50 €. P3: Status nur noch „offen“, „Zahlungserinnerung“ (bis 14 Tage), „Mahnung“ (ab 15 Tage); Auswertung C21 bis C26. Prüfung über alle Lösungsdateien: keine Formel mit mehr als zwei WENN.
- **Neuer Aufbau der Übungsdateien, Stufe 1 (19.11.2026):** Alle Dateien von R1 bis R6 und RM haben jetzt ein Startblatt, den Auftrag oben auf jedem Arbeitsblatt (Tabellen beginnen in Zeile 14, die Zelladressen in den Modulen sind angepasst) und ein Blatt „Corporate Design und Abgabe“. Das Blatt „Aufgabenstellung“ entfällt. Stufen 2 bis 4 folgen. Erzeugt mit `werkzeuge/gen_stufe1_neu.py`.
