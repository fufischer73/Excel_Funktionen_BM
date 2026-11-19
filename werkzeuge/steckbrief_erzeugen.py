#!/usr/bin/env python3
"""Erzeugt modellunternehmen/steckbrief.md aus stammdaten.yaml (nicht von Hand ändern)."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import unternehmen as U

def euro(x): return f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
def pct(x): return f"{x * 100:.0f} %".replace(".", ",")

f = U.D["firma"]
z = []
z.append(f"# Steckbrief: {U.FIRMA}\n")
z.append("Diese Datei wird aus `stammdaten.yaml` erzeugt (`python3 werkzeuge/steckbrief_erzeugen.py`). Alle Angaben sind fiktiv.\n")
z.append("## Das Unternehmen\n")
z.append(f"- **Name:** {U.FIRMA}\n- **Anschrift:** {f['strasse']}, {f['plz']} {f['ort']}\n- **Branche:** {f['branche']}\n- **Lernende arbeiten in:** Abteilung {U.ABTEILUNG} ({f['beruf']})\n")
z.append(f["beschreibung"].strip() + "\n")
z.append("## Standorte\n\n| Standort | Art |\n| --- | --- |")
for x in U.D["filialen"]: z.append(f"| {x['name']} | {x['art']} |")
z.append("\n**Abteilungen:** " + ", ".join(U.D["abteilungen"]) + "\n")
z.append("**Warengruppen:** " + ", ".join(U.WARENGRUPPEN) + "\n")
z.append("## Artikelstamm\n\n| Art.-Nr. | Bezeichnung | Warengruppe | Einheit | Nettopreis |\n| --- | --- | --- | --- | ---: |")
for a in U.D["artikel"]: z.append(f"| {a['nr']} | {a['bezeichnung']} | {a['wg']} | {a['einheit']} | {euro(a['preis'])} |")
z.append("\n## Kunden\n\n| Kunden-Nr. | Name | Ort | Gruppe |\n| --- | --- | --- | --- |")
for k in U.D["kunden"]: z.append(f"| {k['nr']} | {k['name']} | {k['ort']} | {k['gruppe']} |")
z.append("\n| Kundengruppe | Zahlungsziel in Tagen |\n| --- | ---: |")
for g, t in U.ZAHLUNGSZIELE: z.append(f"| {g} | {t} |")
z.append("\n## Lieferanten\n\n| Nr. | Name | Ort |\n| --- | --- | --- |")
for l in U.D["lieferanten"]: z.append(f"| {l['nr']} | {l['name']} | {l['ort']} |")
r = U.R
z.append("\n## Geschäftsregeln\n")
z.append(f"- **Umsatzsteuer:** {pct(r['umsatzsteuer'])} (ermäßigt {pct(r['umsatzsteuer_ermaessigt'])})")
z.append(f"- **Versand:** über {euro(r['versand']['frei_ab_bestellwert'])} Bestellwert versandkostenfrei, sonst {euro(r['versand']['pauschale'])}")
z.append(f"- **Skonto:** {pct(r['skonto']['satz'])} bei Zahlung innerhalb von {r['skonto']['frist_tage']} Tagen")
z.append("- **Mengenrabatt (einfach, höchstens zwei WENN):** " + "; ".join(f"ab {euro(x['ab'])}: {pct(x['satz'])}" for x in r["mengenrabatt_einfach"]))
z.append("- **Mengenrabatt (Staffel, ab Stufe 3 mit SVERWEIS):** " + "; ".join(f"ab {euro(x['ab'])}: {pct(x['satz'])}" for x in r["mengenrabatt_staffel"]))
z.append(f"- **Verzugszinsen:** {pct(r['zinssatz_verzug'])} pro Jahr (Zinstage nach der 360-Tage-Methode)\n")
cd = U.D["corporate_design"]
z.append("## Corporate Design\n")
z.append(f"Schrift {cd['schrift']} {cd['schriftgroesse']}; Überschriften {cd['ueberschrift']['ausrichtung']}, {cd['ueberschrift']['groesse']}, fett; "
         f"Prozent mit {cd['prozent_nachkommastellen']} Nachkommastellen; andere Zahlen mit {cd['zahl_nachkommastellen']}; Euro mit Tausenderpunkt und zwei Nachkommastellen; Datum {cd['datum']}; Uhrzeit {cd['uhrzeit']}.\n")
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "modellunternehmen", "steckbrief.md")
open(out, "w", encoding="utf-8").write("\n".join(z))
print("geschrieben:", os.path.normpath(out))
