# Werkzeuge: Generatoren der Übungsdateien

Python-Skripte (benötigen `openpyxl`), die die Excel-Dateien in `material/` und `lehrkraft/loesungen/` erzeugen. Aufruf immer aus dem Kursordner.

## Reihenfolge
1. Grunddateien erzeugen (Stufen 1 bis 4):
   `python3 werkzeuge/gen_stufe1.py material lehrkraft/loesungen` (ebenso gen_stufe2.py, gen_stufe3.py, gen_stufe4.py)
2. Zusatzblätter einfügen (Stufen 1 bis 3, Meilenstein-Erweiterungen), rechnet die Lösungen neu und schreibt die Aufgabenblöcke in `module/*.md`:
   `python3 werkzeuge/ext_run.py ext_s1` (ebenso ext_s2, ext_s3, ext_s4)
3. Prüfen und bauen: `pruefen.py`, `bauen.py`, `durchklicken.py` aus dem Lernpfad-Baukasten.

## Dateien
- `gen_stufe1.py` bis `gen_stufe4.py`: Grunddateien je Stufe (Stufe 4 enthält die prüfungsnahen Dateien p1 bis p3, pm).
- `ext_framework.py`: Gerüst für Zusatzblätter und Modultexte. `ext_run.py`: Ablauf. `ext_s1.py` bis `ext_s4.py`: Aufgabenspezifikationen (Daten, Formeln, Aufträge).
- `md_orig/`: unveränderte Modultexte als Ausgangsbasis. `ext_run.py` schreibt die Module daraus neu und ist daher wiederholbar.
- `check4.py`: unabhängige Nachrechnung der Werte von Stufe 4.

## Wichtig
- Vor jedem Lauf von `ext_run.py` die Grunddateien der Stufe neu erzeugen (Schritt 1), sonst bleiben alte Blätter bestehen.
- Handänderungen in `module/*.md` an den Aufgabenblöcken gehen beim erneuten Lauf verloren. Dann `md_orig/` und die Spezifikationen anpassen.
- Pfade in `ext_run.py` (`KURS`, `ORIG`) und `ext_framework.py` (Import von `gen_stufe1`) sind auf diese Umgebung eingestellt. Bei anderem Speicherort anpassen.
- Neuberechnung der Lösungen: `/mnt/skills/public/xlsx/scripts/recalc.py` (nur in dieser Umgebung vorhanden, sonst LibreOffice).
