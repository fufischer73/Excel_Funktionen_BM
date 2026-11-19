"""Stammdaten des Modellunternehmens für alle Generatoren.

Liest modellunternehmen/stammdaten.yaml. Verwendung in einem Generator:
    from unternehmen import FIRMA, ADRESSE, ARTIKEL, KUNDEN
Die Tupel haben dieselbe Form wie früher STAMM und KUNDEN in gen_stufe4.py.
"""
import os
import yaml

PFAD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "modellunternehmen", "stammdaten.yaml")


def laden():
    with open(PFAD, encoding="utf-8") as f:
        return yaml.safe_load(f)


D = laden()
_f = D["firma"]
FIRMA = _f["name"]
KURZNAME = _f["kurzname"]
ADRESSE = f"{FIRMA}, {_f['strasse']}, {_f['plz']} {_f['ort']}"
ABTEILUNG = _f["abteilung_lernende"]
FILIALEN = [x["name"] for x in D["filialen"]]
WARENGRUPPEN = list(D["warengruppen"])
ARTIKEL = [(a["nr"], a["bezeichnung"], float(a["preis"]), a["einheit"]) for a in D["artikel"]]
ARTIKEL_WG = {a["nr"]: a["wg"] for a in D["artikel"]}
KUNDEN = [(k["nr"], k["name"], k["ort"], k["gruppe"]) for k in D["kunden"]]
ZAHLUNGSZIELE = [(g["gruppe"], g["zahlungsziel_tage"]) for g in D["kundengruppen"]]
LIEFERANTEN = [(l["nr"], l["name"], l["ort"]) for l in D["lieferanten"]]
R = D["regeln"]
CD = D["corporate_design"]
KOMM = D["kommunikation"]
USt = R["umsatzsteuer"]
RABATT_EINFACH = [(x["ab"], x["satz"]) for x in R["mengenrabatt_einfach"]]
RABATT_STAFFEL = [(x["ab"], x["satz"]) for x in R["mengenrabatt_staffel"]]


def artikel(nr):
    """Artikel nach Nummer: (nr, bezeichnung, preis, einheit)."""
    return next(a for a in ARTIKEL if a[0] == nr)


def kunde(nr):
    return next(k for k in KUNDEN if k[0] == nr)


if __name__ == "__main__":
    print(ADRESSE, len(ARTIKEL), "Artikel,", len(KUNDEN), "Kunden")
