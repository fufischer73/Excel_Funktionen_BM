#!/usr/bin/env python3
"""Vergleicht die erzeugten Lösungsdateien mit den Stammdaten des Modellunternehmens.

Meldet Artikelnummern mit anderer Bezeichnung oder anderem Preis, Kundennummern mit anderem Namen
und Orte, die nicht zum Unternehmen passen. Ausgabe als Markdown (für modellunternehmen/abweichungen.md).
Aufruf aus dem Kursordner:  python3 werkzeuge/stammdaten_pruefen.py
"""
import glob, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import unternehmen as U
from openpyxl import load_workbook

KURS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
art = {a[0]: a for a in U.ARTIKEL}
kun = {k[0]: k for k in U.KUNDEN}
treffer_a, treffer_k = {}, {}
for pfad in sorted(glob.glob(os.path.join(KURS, "lehrkraft/loesungen/*.xlsx"))):
    datei = os.path.basename(pfad).replace("-loesung.xlsx", "")
    wb = load_workbook(pfad, data_only=True)
    for ws in wb:
        for row in ws.iter_rows():
            v = [c.value for c in row]
            for i, x in enumerate(v):
                if isinstance(x, str) and re.fullmatch(r"A-\d{4}", x) and i + 2 < len(v) and isinstance(v[i + 1], str):
                    preis = v[i + 2] if isinstance(v[i + 2], (int, float)) else None
                    # Preis nur prüfen, wenn die Spaltenüberschrift „Preis“ enthält (sonst ist es z. B. eine Menge)
                    kopf = ""
                    for rr in range(row[0].row - 1, 0, -1):
                        h = ws.cell(row=rr, column=i + 3).value
                        if isinstance(h, str):
                            kopf = h
                            break
                    if "preis" not in kopf.lower():
                        preis = None
                    a = art.get(x)
                    if a is None or v[i + 1] != a[1] or (preis is not None and abs(preis - a[2]) > 0.001):
                        treffer_a.setdefault((x, v[i + 1], preis), set()).add(f"{datei}/{ws.title}")
                if isinstance(x, str) and re.fullmatch(r"K-\d{4}", x) and i + 1 < len(v) and isinstance(v[i + 1], str) and not v[i + 1].startswith("A-") and v[i + 1] not in ("Ja", "Nein"):
                    k = kun.get(x)
                    if k is None or v[i + 1] != k[1]:
                        treffer_k.setdefault((x, v[i + 1]), set()).add(f"{datei}/{ws.title}")
print("## Artikel (Nummer, Bezeichnung, Preis weichen vom Stamm ab)\n")
print("| Art.-Nr. | in der Datei | Stamm | wo |\n| --- | --- | --- | --- |")
for (nr, name, preis), orte in sorted(treffer_a.items()):
    a = art.get(nr)
    print(f"| {nr} | {name}, {preis} | {a[1]}, {a[2]}" if a else f"| {nr} | {name}, {preis} | nicht im Stamm", "|", ", ".join(sorted(orte)), "|")
print("\n## Kunden (Nummer oder Name weicht vom Stamm ab)\n")
print("| Kunden-Nr. | in der Datei | Stamm | wo |\n| --- | --- | --- | --- |")
for (nr, name), orte in sorted(treffer_k.items()):
    k = kun.get(nr)
    print(f"| {nr} | {name} | {k[1] if k else 'nicht im Stamm'} | {', '.join(sorted(orte))} |")
