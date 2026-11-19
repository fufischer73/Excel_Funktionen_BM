# Abweichungen vom Modellunternehmen

Stand: 19.11.2026. Vergleich der erzeugten Lösungsdateien mit `stammdaten.yaml`
(Prüfskript: `python3 werkzeuge/stammdaten_pruefen.py`). Die Abweichungen werden beim Umstellen der Stufen 2 bis 4
(Paket 1 bis 3 der Planung) behoben, indem die Generatoren die Daten aus `werkzeuge/unternehmen.py` lesen.

## Automatisch gefunden

### Artikel (Nummer, Bezeichnung, Preis weichen vom Stamm ab)

| Art.-Nr. | in der Datei | Stamm | wo |
| --- | --- | --- | --- |

### Kunden (Nummer oder Name weicht vom Stamm ab)

| Kunden-Nr. | in der Datei | Stamm | wo |
| --- | --- | --- | --- |
| K-4001 | Beispielbau GmbH | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4002 | Muster-Elektro GmbH | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4003 | Test-Gastro e. K. | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4004 | Demo-Möbel GmbH | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4005 | Probe-Druck OHG | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4006 | Beispiel-Spedition AG | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4007 | Modell-Handel KG | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4008 | Muster-Verlag GmbH | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4009 | Beispiel-Bäckerei e. K. | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |
| K-4010 | Test-Reisen GmbH | nicht im Stamm | z1-inventur/Aufgabe 2 Standard |

Behoben am 19.11.2026: Preisliste in R2 (Artikelnummern und Preise jetzt aus dem Stamm).

## Von Hand festgestellt

| Thema | Befund | Vorschlag |
| --- | --- | --- |
| Kundennummern | Die Aufgaben verwenden verschiedene Nummernkreise: K-2301 (R3), K-3101 (R6), K-4001 bis K-4010 (Z1, R5), K-4101 (RM, N-Stufe, P-Stufe). Namen mit „Beispiel…“, „Muster…“ kommen mit mehreren Nummern vor. | Ein Nummernkreis K-4101 bis K-4112 mit festen Namen. Aufgaben, die nur Nummern brauchen, nutzen diese. |
| Versandregel | R3: über 50 € frei, sonst 4,95 €. R4 Aufgabe 3: frei ab 150 € mit Gewichtsstaffel. P1: 7,90 € unter 200 €. | Drei Regeln bewusst verschieden lassen, aber im Steckbrief als „Regel je Aufgabe“ führen, oder auf eine angleichen. |
| Rabatt | R4 und RM: einfache Staffel (0, 5, 10 %). P1, N2, PM: vierstufige Staffel (bis 15 %) mit SVERWEIS. | Passt zur Regel „höchstens zwei WENN, längere Staffeln mit SVERWEIS“. Im Steckbrief beide Staffeln führen (ist geschehen). |
| Artikelmengen | Z1 (Inventur) und andere Dateien bilden eigene Artikellisten (Briefumschläge, Toner, Etiketten …). | Artikel aus dem Stamm verwenden, fehlende Artikel dort ergänzen. |
| Filialen | Rendsburg, Kiel, Flensburg, Neumünster sind überall gleich. Eckernförde und Schleswig kommen nur als Kundenorte vor. | Passt. |
| Personen | Es werden keine Namen von Personen verwendet („Mitarbeit 1“, „Person 1“). | Beibehalten. |
| Lieferanten | Lieferantennamen stehen nur in R4 Aufgabe 4 und Stufe 3. | Aus dem Stamm (`lieferanten`) verwenden. |
