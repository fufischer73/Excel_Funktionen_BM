# Planung: Lernpfad Excel für Kaufleute für Büromanagement

Stand: 19.11.2026 · Ziel: Prüfung in ca. 12 Wochen, Lernen auch zu Hause · Software: Microsoft Excel (ONLYOFFICE nur Nebenvariante)
Arbeitsbranch: `claude/tender-rubin-foxelr`. Hinweise zu Werkzeugen und Aufbau: `uebergabe.md`, `werkzeuge/README.md`.

## 1. Beschlossene Regeln

- Höchstens **zwei WENN** ineinander (drei Ergebnisse). Längere Staffeln über SVERWEIS (Stufe 3).
- Jede Übungsdatei: **Start**, Arbeitsblätter mit **Auftrag oben auf dem Blatt** (Tabelle ab Zeile 14), Blatt **Corporate Design und Abgabe**. Kein Blatt „Aufgabenstellung“ mehr.
- Steigende Komplexität nach **Kopierfähigkeit und Adressierung**: relativ (K1/K2) → eine feste Zelle (K3) → fester Bereich (K4) → Bereich auf anderem Blatt (K5) → gemischt (K6, optional) → kombiniert (K7).
- Funktionen mit **Blattbezügen** üben: SVERWEIS, SUMMEWENN, ZÄHLENWENN, ANZAHL, ANZAHL2.
- **Prüfungsnaher Bereich** getrennt von den Funktionsmodulen. Aufgaben kommen von Jan (Aufgabe mit Excel-Vorlage als Ausgangsdatei).
- Prüfungsdaten (Kammer) nur als **Muster für Aufbau und Spalten**, nicht kopieren, nicht ins Repo. Eigene fiktive Daten (BüroComfort GmbH), keine personenbezogenen Daten.
- Nur Jan setzt `review: freigegeben`.

## 2. Erledigt

| Was | Stand |
| --- | --- |
| Lernpfad mit 21 Modulen, Übungsdateien, Lösungen, Webversion | vorhanden (Entwurf) |
| G1 Formeln und Kopieren, G2 Feste Bezüge (je 60 Min.) | neu, Übungsdatei und Lösung erzeugt |
| R4, RM auf höchstens zwei WENN | fertig |
| R6 neu gegliedert (leicht UND, leicht ODER, Standard, Kombination, anspruchsvoll, Transfer), N4 Aufgabe 3, P3 vereinfacht | fertig |
| Neuer Dateiaufbau für **Stufe 1** (R1 bis R6, RM) | fertig |
| Kann-Liste-Prototyp (`portfolio/kann-liste.html`): Links zu Modulen, Reiter „Mein Lernplan“, 6 neue Kompetenzen | Prototyp |

## 3. Offene Arbeitspakete

Reihenfolge ist ein Vorschlag. Aufwand: S = klein, M = mittel, L = groß.

| Nr. | Paket | Inhalt | Aufwand |
| --- | --- | --- | --- |
| 1 | **Stufe 2 im neuen Aufbau** (Z1 bis Z3, ZM) | Layout wie Stufe 1; neue Aufgaben mit Daten auf einem Blatt und Auswertung auf einem anderen (ANZAHL, ANZAHL2, ZÄHLENWENN, SUMMEWENN) | L |
| 2 | **Stufe 3 im neuen Aufbau** (N1 bis N5, NM) | Layout; SVERWEIS exakt und mit WAHR (Staffeln); Blattbezüge sind teils schon da | L |
| 3 | **Stufe 4 im neuen Aufbau** (P1 bis P3, PM) | Layout; Anbindung an den Prüfungsbereich klären | M |
| 4 | **Hinweis zu festen Bezügen** in R1 und R2 | Plus-Blätter verlangen `$` vor G2: Hinweis auf G2 oder Blätter anpassen | S |
| 5 | **G3 Gestalten und Formatieren** | Corporate Design, Zahlenformate, Druckeinstellungen, Fußzeile; passend zur Kann-Liste | M |
| 6 | **Rechenoperationen** (Dreisatz, Durchschnitte, gewogener Durchschnitt, Prozent, prozentuale Veränderung) | eigene Module mit Übungsdateien | M bis L |
| 7 | **Diagramme** (10 Kompetenzen der Kann-Liste) | Modul(e) mit Übungsdateien (Achsen, Beschriftung, Legende, Trendlinie …) | L |
| 8 | **Prüfungsnaher Bereich** | Themenpakete mit Stammdaten, Bewegungsdaten und Auswertungsblatt (z. B. Lager, Kunden, Einkauf/Lieferanten, Filialen, Personal/Entgelt, Zeit und Fristen); Jans Aufgaben einarbeiten | L |
| 9 | **Kann-Liste fertigstellen** | Stufe 4 und Meilensteine zuordnen, neue Bereiche (G3, Rechnen, Diagramme, Prüfung) verknüpfen, Fortschritt aus dem Lernpfad anzeigen, Texte für Lernende prüfen | M |
| 10 | **Auslieferung** | Web-Build mit dem Baukasten, Moodle-Paket (SCORM), Hosting bei privatem Repo klären, Datenschutz mit der Schule abstimmen | M |
| 11 | **Qualitätssicherung** | alle Dateien in Excel öffnen und Ergebnisse prüfen, 6 Quellen auf `zu_pruefen`, Quellen und Medien für neue Module, Kurz-Checks und Einstufungstests nachziehen, Zeitbilanz und Badge „ca. 10 Doppelstunden“ anpassen, Freigabe durch Jan | M |
| 12 | **12-Wochen-Plan** | Wochenplan mit Unterricht und Zuhause-Lernen, passend zu Kann-Liste und Prüfungstermin | S bis M |

## 4. Offene Entscheidungen (Jan)

1. Prüfungsdaten: eigene Daten mit gleicher Struktur (Empfehlung) oder die Originaldaten 1:1?
2. Reihenfolge: zuerst Stufen 2 bis 4 umstellen (Pakete 1 bis 3) oder zuerst den Prüfungsbereich (Paket 8)?
3. Auslieferung: Moodle (Empfehlung bei privatem Repo und Zuhause-Lernen) oder Web?
4. Gemeinsames Blatt „Corporate Design und Abgabe“ oder zwei getrennte Blätter?
5. Welche Prüfungsdateien zuerst (Vorschlag: Lager und Kunden)?

## 5. Ideensammlung

Hier sammeln wir Ideen, ohne sie sofort umzusetzen. Jede Idee bekommt später eine Nummer in der Tabelle oben oder wird gestrichen.

**Von Jan:**
- (offen, bitte ergänzen)

**Vorschläge von Claude:**
- Formelsammlung/Spickzettel zum Ausdrucken (Funktionen, Tastenkürzel wie F4 und Strg+1, typische Fehlermeldungen `#NV`, `#BEZUG!`, `#WERT!`).
- Fehlersuche-Serie: je Stufe ein Blatt mit eingebauten Fehlern.
- Mini-Prüfungen mit Zeitvorgabe und Punkten, steigend bis zur Prüfungsdauer.
- Prüfwerkzeug für Lehrkräfte: Schülerdatei mit der Musterlösung vergleichen und Punkte vorschlagen.
- Glossar der Fachbegriffe (Zellbezug, Matrix, Kriterium, Stichtag).
- Kann-Liste: Empfehlung des nächsten Moduls, Lernplan als PDF, Wochenziele.
- Kurze Erklärvideos zu den Funktionen (Links in `medien.yaml`).
- Lösungswege in Varianten für Excel und ONLYOFFICE.

## 6. So arbeiten wir weiter

- Eine Sitzung = ein Paket aus Abschnitt 3. Am Ende: Dateien prüfen, committen, pushen, `uebergabe.md` kurz ergänzen.
- Dateien der Prüfung (`Tabellen.zip`) liegen **nicht** im Repo. Bei Bedarf erneut im Chat hochladen.
- Neue Ideen: Abschnitt 5 ergänzen.
