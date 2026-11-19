#!/usr/bin/env python3
"""Tauscht Namen des Modellunternehmens in den Quelldateien aus.

1. In modellunternehmen/stammdaten.yaml den Abschnitt `tausch` füllen: Liste von {alt: ..., neu: ...}
   (längere Texte zuerst, zum Beispiel erst „BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg“, dann „BüroComfort GmbH“).
2. Trockenlauf:   python3 werkzeuge/firma_tauschen.py
   Anwenden:      python3 werkzeuge/firma_tauschen.py --anwenden
3. Danach alle Übungsdateien und Module neu erzeugen (Generatoren, siehe werkzeuge/README.md).
Es werden nur Quelldateien geändert (werkzeuge/*.py, werkzeuge/md_orig, kurs.yaml, README, portfolio, stammdaten.yaml ohne den Abschnitt `tausch`).
"""
import glob, os, sys, yaml

KURS = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
yml = os.path.join(KURS, "modellunternehmen", "stammdaten.yaml")
tausch = yaml.safe_load(open(yml, encoding="utf-8")).get("tausch") or []
if not tausch:
    sys.exit("Der Abschnitt `tausch` in stammdaten.yaml ist leer. Nichts zu tun.")
dateien = (glob.glob(os.path.join(KURS, "werkzeuge", "*.py")) + glob.glob(os.path.join(KURS, "werkzeuge", "md_orig", "*.md")) +
           [os.path.join(KURS, n) for n in ("kurs.yaml", "README.md", "uebergabe.md", "planung.md")] +
           glob.glob(os.path.join(KURS, "portfolio", "*.html")) + [yml])
anwenden = "--anwenden" in sys.argv
for pfad in sorted(dateien):
    t = open(pfad, encoding="utf-8").read()
    kopf, sep, rest = (t.split("\ntausch:", 1) + [""])[:2] if pfad == yml else (t, "", "")
    neu, n = kopf, 0
    for x in tausch:
        n += neu.count(x["alt"])
        neu = neu.replace(x["alt"], x["neu"])
    if n:
        print(f"{n:4d} Ersetzungen: {os.path.relpath(pfad, KURS)}")
        if anwenden:
            open(pfad, "w", encoding="utf-8").write(neu + (("\ntausch:" + rest) if pfad == yml else ""))
print("Angewendet." if anwenden else "Trockenlauf. Mit --anwenden ausführen.")
