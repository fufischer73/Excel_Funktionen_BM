from ext_framework import *
SPECS = {}
def lab(cell, text, bold=False): return ("label", cell, text, bold)
def tgt(cell, f, fmt, label=None, bold=False): return ("target", cell, f, fmt, bold, label)
def inp(cell, v, fmt=None): return ("input", cell, v, fmt)
def ans(cell, text, label=None, rng=None): return ("answer", cell, text, label, rng)
CD = "Formatieren Sie Ihre Ergebnisse nach dem Corporate Design."

# ================================================================ z1
kd = [("K-4001","Beispielbau GmbH","04331 555-101","info@beispielbau.example",12800),("K-4002","Muster-Elektro GmbH","04331 555-102",None,4999.99),
 ("K-4003","Test-Gastro e. K.",None,"kontakt@test-gastro.example",7300),("K-4004","Demo-Möbel GmbH","04331 555-104","mail@demo-moebel.example",None),
 ("K-4005","Probe-Druck OHG",None,None,2150),("K-4006","Beispiel-Spedition AG","04331 555-106","service@beispiel-spedition.example",18900),
 ("K-4007","Modell-Handel KG","04331 555-107",None,None),("K-4008","Muster-Verlag GmbH","04331 555-108","verlag@muster-verlag.example",9100),
 ("K-4009","Beispiel-Bäckerei e. K.",None,"info@beispiel-baeckerei.example",3400),("K-4010","Test-Reisen GmbH","04331 555-110","reisen@test-reisen.example",6150)]
vollst = sum(1 for r in kd if r[2] and r[3])
um = [r[4] for r in kd if r[4] is not None]
um_alle_null = sum(um) / len(kd)
um_mw = sum(um) / len(um)
um_beide = len([r for r in kd if r[4] is None])
um_vollst = [r[1] for r in kd if r[2] and r[3]]

um_ant = [(1,"Service"),(2,"Lieferung"),(3,"Preis")]
um_rows = [("T-01",2,1,3),("T-02",1,"k. A.",2),("T-03",3,2,None),("T-04",2,2,2),("T-05",None,None,None),("T-06",1,1,1),("T-07",4,3,"k. A."),("T-08",2,None,3),
           ("T-09",5,4,4),("T-10",2,2,3),("T-11",1,2,2),("T-12",None,3,None),("T-13",3,3,2),("T-14",2,1,1)]
voll_umfrage = sum(1 for r in um_rows if all(isinstance(x,(int,float)) for x in r[1:]))

SPECS["z1"] = [
dict(name="Aufgabe 2 Standard", niveau="basis", min=15, kurz="Kundenkartei prüfen",
 titel="Lernsituation: Kundenkartei auf Vollständigkeit prüfen",
 sit="Die Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, pflegt eine Kundenkartei. Vor einem Mailing soll geprüft werden, wie vollständig die Kontaktdaten sind und für wie viele Kunden ein Vorjahresumsatz vorliegt. Fehlende Angaben sind leere Zellen.",
 auftr=["Zählen Sie in Spalte F, wie viele Kontaktangaben (Telefon und E-Mail) je Kunde vorhanden sind.","Zeigen Sie in Spalte G mit „ja“ oder „nein“ an, ob ein Umsatz erfasst ist.",
        "Ermitteln Sie in den Zellen B25 bis B30 die Anzahl der Kunden, der Kunden mit und ohne Telefon sowie mit und ohne Umsatzangabe.",
        "Berechnen Sie in B31 den durchschnittlichen Umsatz der Kunden mit Umsatzangabe und in B32 den Wert, der sich ergäbe, wenn die leeren Umsätze als 0 mitgezählt würden.",
        "Beantworten Sie in B33: Bei wie vielen Kunden sind Telefon und E-Mail vorhanden?", CD],
 cols=[("Kunden-Nr.",12,None,"d",""),("Firma",26,None,"d",""),("Telefon",16,None,"d",""),("E-Mail",32,None,"d",""),("Umsatz Vorjahr",16,EURO,"d",""),
       ("Kontaktangaben",14,ZAHL,"f","=COUNTA(C{r}:D{r})"),("Umsatz erfasst?",14,None,"f",'=IF(COUNT(E{r})=1,"ja","nein")')],
 rows=kd,
 extras=[lab("A25","Kunden insgesamt",True),tgt("B25","=COUNTA(A14:A23)",ZAHL,"Kunden insgesamt"),lab("A26","Kunden mit Telefon",True),tgt("B26","=COUNTA(C14:C23)",ZAHL,"Kunden mit Telefon"),
   lab("A27","Kunden ohne Telefon",True),tgt("B27","=B25-B26",ZAHL,"Kunden ohne Telefon"),lab("A28","Kunden mit Umsatzangabe",True),tgt("B28","=COUNT(E14:E23)",ZAHL,"Kunden mit Umsatzangabe"),
   lab("A29","Kunden ohne Umsatzangabe",True),tgt("B29","=B25-B28",ZAHL,"Kunden ohne Umsatzangabe"),lab("A30","Mit Telefon und E-Mail",True),
   lab("A31","Durchschnitt (nur erfasste)",True),tgt("B31","=AVERAGE(E14:E23)",EURO,"Durchschnitt der erfassten Umsätze"),
   lab("A32","Durchschnitt (leer als 0)",True),tgt("B32","=SUM(E14:E23)/B25",EURO,"Durchschnitt mit leeren Zellen als 0"),
   lab("A33","Telefon und E-Mail vorhanden",True),ans("B33",f"{vollst} Kunden",None,"B33:D33")],
 widths={"A":30},
 checks=["B25 + B27 stimmen: Kunden mit Telefon und ohne Telefon ergeben zusammen alle Kunden.","Der Durchschnitt mit leeren Zellen als 0 ist niedriger als der Durchschnitt der erfassten Umsätze."]),
dict(name="Aufgabe 3 Transfer", niveau="plus", min=20, kurz="Kundenumfrage auswerten",
 titel="Lernsituation: Kundenumfrage auswerten",
 sit="Die BüroComfort GmbH hat 14 Kunden gebeten, den Service, die Lieferung und den Preis mit Schulnoten von 1 bis 6 zu bewerten. Manche Kunden haben Fragen nicht beantwortet (leere Zelle) oder „k. A.“ eingetragen.",
 auftr=["Berechnen Sie je Teilnehmer in Spalte E die Anzahl der Noten (nur Zahlen) und in Spalte F die Anzahl der Einträge (auch Text).","Zeigen Sie in Spalte G mit „ja“ oder „nein“ an, ob alle drei Fragen mit einer Note beantwortet wurden.",
        "Ermitteln Sie in den Zeilen 30 bis 34 für jede Frage die Anzahl der Noten, die Anzahl der Einträge, die Rücklaufquote (Noten im Verhältnis zur Teilnehmerzahl) und die Durchschnittsnote.",
        "Beantworten Sie in B36 und B37 die Fragen unter der Tabelle.", CD],
 cols=[("Teilnehmer",12,None,"d",""),("Service",12,ZAHL,"d",""),("Lieferung",12,ZAHL,"d",""),("Preis",12,ZAHL,"d",""),
       ("Noten",10,ZAHL,"f","=COUNT(B{r}:D{r})"),("Einträge",10,ZAHL,"f","=COUNTA(B{r}:D{r})"),("Vollständig?",13,None,"f",'=IF(E{r}=3,"ja","nein")')],
 rows=um_rows,
 extras=[lab("A29","Teilnehmer",True),tgt("B29","=COUNTA(A14:A27)",ZAHL,"Teilnehmer"),lab("A30","Noten",True),lab("A31","Einträge",True),lab("A32","Rücklaufquote",True),lab("A33","Durchschnittsnote",True)]+
   [tgt(f"{c}30",f"=COUNT({c}14:{c}27)",ZAHL,f"Noten {n}") for c,n in zip("BCD",["Service","Lieferung","Preis"])]+
   [tgt(f"{c}31",f"=COUNTA({c}14:{c}27)",ZAHL,f"Einträge {n}") for c,n in zip("BCD",["Service","Lieferung","Preis"])]+
   [tgt(f"{c}32",f"={c}30/$B$29",PROZ,f"Rücklaufquote {n}") for c,n in zip("BCD",["Service","Lieferung","Preis"])]+
   [tgt(f"{c}33",f"=AVERAGE({c}14:{c}27)","0.00",f"Durchschnittsnote {n}") for c,n in zip("BCD",["Service","Lieferung","Preis"])]+
   [lab("A36","Vollständige Fragebögen",True),ans("B36",f"{voll_umfrage} Teilnehmer",None,"B36:D36"),
    lab("A37","Warum ist bei „Lieferung“ die Zahl der Einträge höher als die der Noten?",True),ans("B37","ANZAHL2 zählt auch den Text „k. A.“, ANZAHL zählt nur Zahlen.",None,"B37:G37")],
 widths={"A":30},
 checks=["Bei T-02 (Lieferung „k. A.“) ist die Zahl der Einträge größer als die Zahl der Noten.","T-05 hat keinen Eintrag: Noten 0, Einträge 0."]),
]

# ================================================================ z2
rk = [("RK-01","Transportschaden","Nord",3,"erledigt"),("RK-02","Falschlieferung","Süd",2,"erledigt"),("RK-03","Qualität","Ost",8,"offen"),("RK-04","Verspätung","West",1,"erledigt"),
 ("RK-05","Transportschaden","Süd",6,"erledigt"),("RK-06","Qualität","Nord",4,"erledigt"),("RK-07","Falschlieferung","Ost",5,"erledigt"),("RK-08","Verspätung","Nord",2,"erledigt"),
 ("RK-09","Transportschaden","Ost",9,"offen"),("RK-10","Qualität","West",7,"offen"),("RK-11","Falschlieferung","Nord",3,"erledigt"),("RK-12","Verspätung","Süd",1,"erledigt"),
 ("RK-13","Transportschaden","West",5,"erledigt"),("RK-14","Qualität","Süd",10,"offen"),("RK-15","Falschlieferung","West",2,"erledigt"),("RK-16","Verspätung","Ost",4,"erledigt")]
bs = [("B-6001","Beispielbau GmbH","Rendsburg",840.5,"Rechnung"),("B-6002","Muster-Elektro GmbH","Kiel",1250,"Vorkasse"),("B-6003","Test-Gastro e. K.","Eckernförde",96.8,"Rechnung"),
 ("B-6004","Demo-Möbel GmbH","Rendsburg",3120,"Rechnung"),("B-6005","Probe-Druck OHG","Büdelsdorf",215.9,"Vorkasse"),("B-6006","Beispiel-Spedition AG","Kiel",4890,"Rechnung"),
 ("B-6007","Modell-Handel KG","Rendsburg",499.99,"Lastschrift"),("B-6008","Muster-Verlag GmbH","Eckernförde",650,"Lastschrift"),("B-6009","Beispiel-Bau & Garten e. K.","Büdelsdorf",305.4,"Rechnung"),
 ("B-6010","Test-Reisen GmbH","Rendsburg",78.2,"Vorkasse"),("B-6011","Demo-Tiefbau GmbH","Kiel",2100,"Rechnung"),("B-6012","Probe-Hausbau OHG","Rendsburg",500,"Lastschrift"),
 ("B-6013","Modell-Elektro KG","Büdelsdorf",129,"Vorkasse"),("B-6014","Muster-Bäckerei e. K.","Eckernförde",45.9,"Lastschrift")]
ret = [("R-01","Technik","defekt",85),("R-02","Papier","Fehlbestellung",12.5),("R-03","Büromöbel","Transportschaden",240),("R-04","Technik","defekt",130),("R-05","Schreibwaren","Fehlbestellung",9.9),
 ("R-06","Technik","Transportschaden",310),("R-07","Büromöbel","defekt",175),("R-08","Papier","Transportschaden",22),("R-09","Technik","Fehlbestellung",49),("R-10","Schreibwaren","defekt",6.5),
 ("R-11","Büromöbel","Fehlbestellung",98),("R-12","Technik","defekt",160),("R-13","Papier","Fehlbestellung",18),("R-14","Büromöbel","defekt",205),("R-15","Technik","Transportschaden",75),
 ("R-16","Schreibwaren","Transportschaden",14),("R-17","Büromöbel","Transportschaden",330),("R-18","Papier","defekt",8),("R-19","Technik","defekt",99),("R-20","Büromöbel","Fehlbestellung",120)]
from collections import Counter
cw = Counter(r[1] for r in ret); top_wg = cw.most_common(1)[0]
n_wg = {"Büromöbel":0,"Schreibwaren":0,"Papier":0,"Technik":0}; n_wg.update(cw)

SPECS["z2"] = [
dict(name="Aufgabe 2 Standard", niveau="basis", min=15, kurz="Reklamationen auswerten",
 titel="Lernsituation: Reklamationen auswerten",
 sit="Die Qualitätssicherung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, hat 16 Reklamationen erfasst. Die Geschäftsleitung möchte wissen, welche Gründe, Filialen und Status am häufigsten vorkommen. Die Begriffe stehen links in Spalte H, die Ergebnisse gehören in Spalte I.",
 auftr=["Zählen Sie in I14 bis I17 die Reklamationen je Grund. Verwenden Sie als Suchkriterium jeweils die Zelle in Spalte H und setzen Sie den Bereich fest.","Zählen Sie in I19 bis I22 die Reklamationen der Filiale Nord, die offenen Reklamationen und die Reklamationen mit einer Bearbeitungsdauer von mehr als 5 Tagen.",
        "Berechnen Sie in I23 den Anteil der offenen Reklamationen an allen Reklamationen.","Zeigen Sie in Spalte F die Priorität an: „hoch“ bei mehr als 5 Tagen Bearbeitung, sonst „normal“.","Berechnen Sie in I18 die Summe der Zählungen je Grund. Sie muss der Zahl aller Reklamationen entsprechen.", CD],
 cols=[("Nr.",10,None,"d",""),("Grund",20,None,"d",""),("Filiale",10,None,"d",""),("Bearbeitung (Tage)",14,ZAHL,"d",""),("Status",12,None,"d",""),("Priorität",12,None,"f",'=IF(D{r}>5,"hoch","normal")')],
 rows=rk,
 extras=[lab("H13","Auswertung",True),lab("I13","Anzahl",True)]+
   [lab(f"H{14+i}",g) for i,g in enumerate(["Transportschaden","Falschlieferung","Qualität","Verspätung"])]+
   [tgt(f"I{14+i}",f"=COUNTIF($B$14:$B$29,H{14+i})",ZAHL,f"Anzahl {g}") for i,g in enumerate(["Transportschaden","Falschlieferung","Qualität","Verspätung"])]+
   [lab("H18","Summe",True),tgt("I18","=SUM(I14:I17)",ZAHL,"Summe je Grund"),
    lab("H19","Nord"),tgt("I19","=COUNTIF($C$14:$C$29,H19)",ZAHL,"Filiale Nord"),lab("H20","offen"),tgt("I20","=COUNTIF($E$14:$E$29,H20)",ZAHL,"Status offen"),
    lab("H21","Mehr als 5 Tage"),tgt("I21",'=COUNTIF($D$14:$D$29,">5")',ZAHL,"Mehr als 5 Tage"),lab("H22","Alle Reklamationen"),tgt("I22","=COUNTA(A14:A29)",ZAHL,"Alle Reklamationen"),
    lab("H23","Anteil offen",True),tgt("I23","=I20/I22",PROZ,"Anteil offen")],
 widths={"H":26,"I":12},
 checks=["I18 stimmt mit I22 überein.","Prüfen Sie, dass sich die Zählungen ändern, wenn Sie in Spalte E einen Status von „offen“ auf „erledigt“ ändern."]),
dict(name="Aufgabe 3 Anspruchsvoll", niveau="plus", min=20, kurz="Platzhalter und Schwellenwert",
 titel="Lernsituation: Bestellungen mit Kriterien auswerten",
 sit="Der Vertrieb der BüroComfort GmbH wertet 14 Bestellungen aus. Teils stehen Kriterien als Text mit Platzhaltern, teils als Vergleich mit einem Wert. Der Schwellenwert steht in Zelle I15 und kann geändert werden.",
 auftr=["Zählen Sie in I17 die Firmen mit „bau“ im Namen (Platzhalter *).","Zählen Sie in I18 die Bestellungen aus Orten, die mit „R“ beginnen.","Zählen Sie in I19 die Bestellungen mit einem Bestellwert über dem Schwellenwert in I15 und in I20 die Bestellungen mit einem Bestellwert mindestens in Höhe des Schwellenwerts. Das Kriterium muss auf die Zelle I15 verweisen.",
        "Zählen Sie in I21 die Bestellungen, die nicht aus Rendsburg kommen, und in I22 die Bestellungen, die nicht per Rechnung bezahlt werden.","Zählen Sie in I23 die Bestellungen mit einem Wert von 100,00 € bis einschließlich 500,00 €. Überlegen Sie, wie Sie aus zwei Zählungen die Anzahl dazwischen erhalten.",
        "Ändern Sie den Schwellenwert in I15 auf 1.000 €. Welche Zellen ändern sich?"],
 cols=[("Bestellung",12,None,"d",""),("Firma",30,None,"d",""),("Ort",14,None,"d",""),("Bestellwert",14,EURO,"d",""),("Zahlungsart",14,None,"d","")],
 rows=bs,
 extras=[lab("H14","Auswertung",True),lab("H15","Schwellenwert (Bestellwert)",True),inp("I15",500,EURO),
   lab("H17","Firma enthält „bau“"),tgt("I17",'=COUNTIF($B$14:$B$27,"*bau*")',ZAHL,"Firma enthält „bau“"),
   lab("H18","Ort beginnt mit „R“"),tgt("I18",'=COUNTIF($C$14:$C$27,"R*")',ZAHL,"Ort beginnt mit „R“"),
   lab("H19","Wert über Schwellenwert"),tgt("I19",'=COUNTIF($D$14:$D$27,">"&$I$15)',ZAHL,"Wert über Schwellenwert"),
   lab("H20","Wert mindestens Schwellenwert"),tgt("I20",'=COUNTIF($D$14:$D$27,">="&$I$15)',ZAHL,"Wert mindestens Schwellenwert"),
   lab("H21","Ort nicht Rendsburg"),tgt("I21",'=COUNTIF($C$14:$C$27,"<>Rendsburg")',ZAHL,"Ort nicht Rendsburg"),
   lab("H22","Zahlungsart nicht Rechnung"),tgt("I22",'=COUNTIF($E$14:$E$27,"<>Rechnung")',ZAHL,"Zahlungsart nicht Rechnung"),
   lab("H23","Wert von 100 bis 500 €"),tgt("I23",'=COUNTIF($D$14:$D$27,">=100")-COUNTIF($D$14:$D$27,">500")',ZAHL,"Wert von 100 bis 500 €")],
 widths={"H":32,"I":12},
 checks=["Die Bestellung B-6012 hat genau 500,00 €: Sie zählt bei „mindestens“ mit, bei „über“ nicht.","Bei einem Schwellenwert von 1.000 € ändern sich nur I19 und I20."]),
dict(name="Aufgabe 4 Transfer", niveau="plus", min=20, kurz="Retouren auswerten",
 titel="Lernsituation: Retouren nach Warengruppe und Grund",
 sit="Die BüroComfort GmbH hat 20 Retouren erfasst. Die Geschäftsleitung möchte wissen, welche Warengruppe und welcher Grund besonders häufig vorkommt. Die Auswertungstabellen sind mit den Suchbegriffen vorbereitet.",
 auftr=["Zählen Sie in I14 bis I17 die Retouren je Warengruppe und in I21 bis I23 die Retouren je Grund. Verwenden Sie als Kriterium jeweils den Suchbegriff in Spalte H.","Berechnen Sie in den Spalten J den Anteil an allen Retouren.",
        "Bilden Sie in Zeile 18 und Zeile 24 die Summen und prüfen Sie in I19, ob die Summe der Warengruppen mit der Zahl aller Retouren übereinstimmt (Text „stimmt“ oder „stimmt nicht“).","Beantworten Sie in B36, welche Warengruppe die meisten Retouren hat und wie hoch ihr Anteil ist.", CD],
 cols=[("Retoure",10,None,"d",""),("Warengruppe",16,None,"d",""),("Grund",18,None,"d",""),("Wert",12,EURO,"d","")],
 rows=ret,
 extras=[lab("H13","Warengruppe",True),lab("I13","Anzahl",True),lab("J13","Anteil",True)]+
   [lab(f"H{14+i}",g) for i,g in enumerate(["Büromöbel","Schreibwaren","Papier","Technik"])]+
   [tgt(f"I{14+i}",f"=COUNTIF($B$14:$B$33,H{14+i})",ZAHL,f"Anzahl {g}") for i,g in enumerate(["Büromöbel","Schreibwaren","Papier","Technik"])]+
   [tgt(f"J{14+i}",f"=I{14+i}/$I$18",PROZ,f"Anteil {g}") for i,g in enumerate(["Büromöbel","Schreibwaren","Papier","Technik"])]+
   [lab("H18","Summe",True),tgt("I18","=SUM(I14:I17)",ZAHL,"Summe Warengruppen"),tgt("J18","=SUM(J14:J17)",PROZ,"Summe Anteile"),
    lab("H19","Kontrolle"),tgt("I19",'=IF(I18=COUNTA(A14:A33),"stimmt","stimmt nicht")',None,"Kontrolle"),
    lab("H20","Grund",True),lab("I20","Anzahl",True),lab("J20","Anteil",True)]+
   [lab(f"H{21+i}",g) for i,g in enumerate(["defekt","Fehlbestellung","Transportschaden"])]+
   [tgt(f"I{21+i}",f"=COUNTIF($C$14:$C$33,H{21+i})",ZAHL,f"Anzahl {g}") for i,g in enumerate(["defekt","Fehlbestellung","Transportschaden"])]+
   [tgt(f"J{21+i}",f"=I{21+i}/$I$24",PROZ,f"Anteil {g}") for i,g in enumerate(["defekt","Fehlbestellung","Transportschaden"])]+
   [lab("H24","Summe",True),tgt("I24","=SUM(I21:I23)",ZAHL,"Summe Gründe"),
    lab("A36","Warengruppe mit den meisten Retouren",True),ans("B36",f"{top_wg[0]} mit {top_wg[1]} von 20 Retouren ({top_wg[1]*5} %)",None,"B36:G36")],
 widths={"H":18,"I":10,"J":10,"A":34},
 checks=["Die Summe der Anteile ergibt 100,00 %."]),
]

# ================================================================ z3
sp = [("S-01","Fahrt","Verkauf",84.5),("S-02","Material","Verwaltung",36.9),("S-03","Bewirtung","Verkauf",112.4),("S-04","Fahrt","Einkauf",47.2),("S-05","Unterkunft","Verkauf",189),
 ("S-06","Fahrt","Verkauf",61.8),("S-07","Material","Einkauf",254.3),("S-08","Bewirtung","Verwaltung",68),("S-09","Fahrt","Verwaltung",22.5),("S-10","Unterkunft","Einkauf",156),
 ("S-11","Material","Verkauf",19.9),("S-12","Fahrt","Einkauf",93.1),("S-13","Bewirtung","Einkauf",74.6),("S-14","Unterkunft","Verkauf",210),("S-15","Material","Verwaltung",48.75),("S-16","Fahrt","Verkauf",118.4)]
au = [("A-101","Beispielbau GmbH","erledigt",1840),("A-102","Muster-Elektro GmbH","offen",620.5),("A-103","Test-Gastro e. K.","storniert",310),("A-104","Demo-Möbel GmbH","erledigt",3250),
 ("A-105","Probe-Druck OHG","erledigt",480),("A-106","Beispiel-Spedition AG","offen",5120),("A-107","Modell-Handel KG","storniert",890),("A-108","Muster-Verlag GmbH","erledigt",1275.4),
 ("A-109","Beispiel-Bau & Garten e. K.","offen",735),("A-110","Test-Reisen GmbH","erledigt",500),("A-111","Demo-Tiefbau GmbH","erledigt",2680),("A-112","Probe-Hausbau OHG","storniert",1450),
 ("A-113","Modell-Elektro KG","offen",205),("A-114","Muster-Bäckerei e. K.","erledigt",96.8)]
vk = [("V-01","V2",2200),("V-02","V1",1250),("V-03","V4",1500),("V-04","V3",640),("V-05","V1",980),("V-06","V2",3150),("V-07","V3",720),("V-08","V4",2480),("V-09","V1",1820),
 ("V-10","V2",1890),("V-11","V3",1380),("V-12","V4",1620),("V-13","V1",2350),("V-14","V3",910),("V-15","V2",860),("V-16","V4",1740),("V-17","V1",1040),("V-18","V3",550)]
SPECS["z3"] = [
dict(name="Aufgabe 2 Standard", niveau="basis", min=15, kurz="Spesen nach Kostenart",
 titel="Lernsituation: Spesen nach Kostenart und Abteilung auswerten",
 sit="Die Buchhaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, hat 16 Spesenbelege erfasst. Für die Kostenplanung werden die Beträge je Kostenart und je Abteilung benötigt.",
 auftr=["Zeigen Sie in Spalte E „Belegprüfung“ an, wenn der Betrag über 100,00 € liegt, sonst „-“.","Berechnen Sie in I14 bis I17 die Summe der Beträge je Kostenart und in J14 bis J17 den Anteil an der Gesamtsumme. Verwenden Sie als Kriterium die Zelle in Spalte H.",
        "Berechnen Sie in I21 bis I23 die Summe der Beträge je Abteilung.","Bilden Sie die Gesamtsummen in I18 und I24 und kontrollieren Sie in I25, ob beide Summen mit der Summe der Spalte D übereinstimmen (Text „stimmt“ oder „stimmt nicht“).", CD],
 cols=[("Beleg",10,None,"d",""),("Kostenart",14,None,"d",""),("Abteilung",14,None,"d",""),("Betrag",12,EURO,"d",""),("Hinweis",14,None,"f",'=IF(D{r}>100,"Belegprüfung","-")')],
 rows=sp,
 extras=[lab("H13","Kostenart",True),lab("I13","Summe",True),lab("J13","Anteil",True)]+
   [lab(f"H{14+i}",k) for i,k in enumerate(["Fahrt","Bewirtung","Material","Unterkunft"])]+
   [tgt(f"I{14+i}",f"=SUMIF($B$14:$B$29,H{14+i},$D$14:$D$29)",EURO,f"Summe {k}") for i,k in enumerate(["Fahrt","Bewirtung","Material","Unterkunft"])]+
   [tgt(f"J{14+i}",f"=I{14+i}/$I$18",PROZ,f"Anteil {k}") for i,k in enumerate(["Fahrt","Bewirtung","Material","Unterkunft"])]+
   [lab("H18","Summe",True),tgt("I18","=SUM(I14:I17)",EURO,"Gesamtsumme Kostenarten"),lab("H20","Abteilung",True),lab("I20","Summe",True)]+
   [lab(f"H{21+i}",k) for i,k in enumerate(["Verkauf","Einkauf","Verwaltung"])]+
   [tgt(f"I{21+i}",f"=SUMIF($C$14:$C$29,H{21+i},$D$14:$D$29)",EURO,f"Summe {k}") for i,k in enumerate(["Verkauf","Einkauf","Verwaltung"])]+
   [lab("H24","Summe",True),tgt("I24","=SUM(I21:I23)",EURO,"Gesamtsumme Abteilungen"),lab("H25","Kontrolle"),
    tgt("I25",'=IF(AND(ABS(I18-SUM(D14:D29))<0.005,ABS(I24-SUM(D14:D29))<0.005),"stimmt","stimmt nicht")',None,"Kontrolle")],
 widths={"H":16,"I":14,"J":10},
 checks=["Die Anteile in J14 bis J17 ergeben zusammen 100,00 %."]),
dict(name="Aufgabe 3 Anspruchsvoll", niveau="plus", min=20, kurz="Aufträge mit Bedingungen",
 titel="Lernsituation: Auftragsumsätze mit Bedingungen",
 sit="Der Vertrieb der BüroComfort GmbH wertet 14 Aufträge aus. Stornierte Aufträge dürfen nicht zum Umsatz zählen. Der Schwellenwert für große Aufträge steht in Zelle I15 und kann geändert werden.",
 auftr=["Berechnen Sie in I17 den Umsatz der erledigten Aufträge und in I18 den Umsatz aller Aufträge, die nicht storniert sind.","Berechnen Sie in I19 die Summe der Aufträge mit einem Nettowert über dem Schwellenwert aus I15. Das Kriterium muss auf die Zelle verweisen.",
        "Berechnen Sie in I20 den Umsatz der Kunden mit „bau“ im Namen und in I21 den Umsatz der Kunden, deren Name mit „Muster“ beginnt.","Berechnen Sie in I22 den Anteil der stornierten Aufträge am Gesamtwert aller Aufträge.",
        "Prüfen Sie: I17 plus der Umsatz der offenen Aufträge ergibt I18. Ermitteln Sie den Umsatz der offenen Aufträge in I23."],
 cols=[("Auftrag",10,None,"d",""),("Kunde",30,None,"d",""),("Status",12,None,"d",""),("Nettowert",14,EURO,"d","")],
 rows=au,
 extras=[lab("H14","Auswertung",True),lab("H15","Schwellenwert (Nettowert)",True),inp("I15",1000,EURO),
   lab("H17","Umsatz erledigt"),tgt("I17",'=SUMIF($C$14:$C$27,"erledigt",$D$14:$D$27)',EURO,"Umsatz erledigt"),
   lab("H18","Umsatz ohne Stornos"),tgt("I18",'=SUMIF($C$14:$C$27,"<>storniert",$D$14:$D$27)',EURO,"Umsatz ohne Stornos"),
   lab("H19","Aufträge über Schwellenwert"),tgt("I19",'=SUMIF($D$14:$D$27,">"&$I$15)',EURO,"Aufträge über Schwellenwert"),
   lab("H20","Kunden mit „bau“"),tgt("I20",'=SUMIF($B$14:$B$27,"*bau*",$D$14:$D$27)',EURO,"Kunden mit „bau“"),
   lab("H21","Kunden „Muster …“"),tgt("I21",'=SUMIF($B$14:$B$27,"Muster*",$D$14:$D$27)',EURO,"Kunden „Muster …“"),
   lab("H22","Anteil Stornos"),tgt("I22",'=SUMIF($C$14:$C$27,"storniert",$D$14:$D$27)/SUM($D$14:$D$27)',PROZ,"Anteil Stornos"),
   lab("H23","Umsatz offen"),tgt("I23",'=SUMIF($C$14:$C$27,"offen",$D$14:$D$27)',EURO,"Umsatz offen")],
 widths={"H":32,"I":14},
 checks=["I17 + I23 ergibt I18.","Bei I19 ohne Summenbereich wird der geprüfte Bereich selbst summiert."]),
dict(name="Aufgabe 4 Transfer", niveau="plus", min=20, kurz="Provision je Mitarbeit berechnen",
 titel="Lernsituation: Provisionsabrechnung",
 sit="Im Verkauf der BüroComfort GmbH arbeiten vier Personen (V1 bis V4). Aus der Verkaufsliste soll die Provision berechnet werden. Der Provisionssatz richtet sich nach dem Gesamtumsatz der Person: ab 8.000 € 3 %, ab 5.000 € 2 %, darunter 1 %.",
 auftr=["Berechnen Sie in Spalte D für jeden Verkauf den Anteil am Gesamtumsatz aller Verkäufe.","Berechnen Sie in I14 bis I17 den Umsatz je Person mit SUMMEWENN. Der Suchbegriff steht jeweils in Spalte H.","Ermitteln Sie in J14 bis J17 den Provisionssatz mit einer verschachtelten WENN-Funktion und in K14 bis K17 die Provision.",
        "Berechnen Sie in Zeile 18 die Summen. Prüfen Sie in I19, ob die Summe der Personenumsätze mit der Summe der Spalte D übereinstimmt (Text „stimmt“ oder „stimmt nicht“)."],
 cols=[("Verkauf",10,None,"d",""),("Person",10,None,"d",""),("Umsatz",12,EURO,"d",""),("Anteil",10,PROZ,"f","=C{r}/SUM($C$14:$C$31)")],
 rows=vk,
 extras=[lab("H13","Person",True),lab("I13","Umsatz",True),lab("J13","Satz",True),lab("K13","Provision",True)]+
   [lab(f"H{14+i}",p) for i,p in enumerate(["V1","V2","V3","V4"])]+
   [tgt(f"I{14+i}",f"=SUMIF($B$14:$B$31,H{14+i},$C$14:$C$31)",EURO,f"Umsatz {p}") for i,p in enumerate(["V1","V2","V3","V4"])]+
   [tgt(f"J{14+i}",f"=IF(I{14+i}>=8000,0.03,IF(I{14+i}>=5000,0.02,0.01))",PROZ,f"Satz {p}") for i,p in enumerate(["V1","V2","V3","V4"])]+
   [tgt(f"K{14+i}",f"=I{14+i}*J{14+i}",EURO,f"Provision {p}") for i,p in enumerate(["V1","V2","V3","V4"])]+
   [lab("H18","Summe",True),tgt("I18","=SUM(I14:I17)",EURO,"Summe Umsätze"),tgt("K18","=SUM(K14:K17)",EURO,"Summe Provisionen"),
    lab("H19","Kontrolle"),tgt("I19",'=IF(I18=SUM(C14:C31),"stimmt","stimmt nicht")',None,"Kontrolle")],
 widths={"H":12,"I":14,"J":10,"K":14},
 checks=["Nur V2 erreicht die höchste Stufe."]),
]
