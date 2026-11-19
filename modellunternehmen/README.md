# Modellunternehmen

Hier liegen die Stammdaten des Modellunternehmens, aus denen alle Übungsdateien und Modultexte entstehen sollen.

| Datei | Zweck |
| --- | --- |
| `stammdaten.yaml` | **Einzige Quelle**: Firma, Standorte, Warengruppen, Artikel, Kunden, Lieferanten, Geschäftsregeln, Corporate Design |
| `steckbrief.md` | Lesbare Fassung für Lehrkräfte (wird erzeugt, nicht von Hand ändern) |
| `abweichungen.md` | Stellen, an denen die erzeugten Dateien noch nicht zu den Stammdaten passen |

## Werkzeuge

- `werkzeuge/unternehmen.py`: lädt die Stammdaten für die Generatoren (`FIRMA`, `ADRESSE`, `ARTIKEL`, `KUNDEN`, `LIEFERANTEN` …).
- `werkzeuge/steckbrief_erzeugen.py`: schreibt `steckbrief.md`.
- `werkzeuge/stammdaten_pruefen.py`: vergleicht die Lösungsdateien mit den Stammdaten.
- `werkzeuge/firma_tauschen.py`: tauscht Namen in den Quelldateien aus (Trockenlauf, dann `--anwenden`).

## Anderes Modellunternehmen verwenden

1. Rechte klären (siehe `planung.md`, Entscheidung 9).
2. In `stammdaten.yaml` die Angaben unter `firma` anpassen und im Abschnitt `tausch` die alten und neuen Schreibweisen eintragen.
3. `python3 werkzeuge/firma_tauschen.py` (Trockenlauf), dann mit `--anwenden`.
4. Übungsdateien und Module neu erzeugen (Generatoren, siehe `werkzeuge/README.md`) und `steckbrief_erzeugen.py` ausführen.

Hinweis: Sind Stufe 2 bis 4 noch nicht auf `unternehmen.py` umgestellt, kommen Namen dort aus dem Quelltext der Generatoren. `firma_tauschen.py` ersetzt sie trotzdem.
