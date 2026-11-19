from ext_framework import *
from ext_s2 import lab, tgt, inp, ans, CD
import random
SPECS = {}
rm13 = [("B-5101","K-4201","Ja","DE",620.0),("B-5102","K-4202","Nein","AT",95.5),("B-5103","K-4203","Ja","DE",1250.0),("B-5104","K-4204","Nein","CH",480.0),("B-5105","K-4205","Ja","DK",340.0),("B-5106","K-4206","Ja","DE",500.0),
        ("B-5107","K-4207","Nein","DE",1000.0),("B-5108","K-4208","Ja","AT",299.99),("B-5109","K-4209","Ja","DE",300.0),("B-5110","K-4210","Nein","AT",760.4),("B-5111","K-4211","Ja","DE",100.0),("B-5112","K-4212","Ja","CH",1890.0)]
SPECS["rm"] = [dict(name="Erweiterung KW 13", niveau="plus", min=20, kurz="Kalkulation mit Express-Versand", nr="Erweiterung",
 titel="Erweiterung: Bestellungen der KW 13 mit Express-Versand",
 sit="Für die KW 13 gelten dieselben Regeln wie in der KW 12 (Rabattstaffel, Zusatzbonus). Neu ist die Versandart „Express“: Bonuskunden aus Deutschland (DE) mit einem Bestellwert ab 500,00 € erhalten Expressversand. Alle anderen Bestellungen aus DE oder AT gehen als „Standard“, alle übrigen als „Ausland“.",
 auftr=["Ermitteln Sie in Spalte F den Rabattsatz (5 % ab 100,00 €, 10 % ab 500,00 €) und in Spalte G den Rabatt in Euro.","Berechnen Sie in Spalte H den Zusatzbonus (2 % des Bestellwerts für Bonuskunden ab 300,00 €).",
        "Legen Sie in Spalte I die Versandart fest (Express, Standard oder Ausland).","Berechnen Sie in Spalte J den Zahlbetrag.","Ermitteln Sie in den Zellen B28 bis B31 die Summe der Zahlbeträge, den Durchschnitt, den größten Zahlbetrag und die Anzahl der Expresssendungen.", CD],
 cols=[("Bestell-Nr.",12,None,"d",""),("Kunden-Nr.",12,None,"d",""),("Bonuskunde",12,None,"d",""),("Land",8,None,"d",""),("Bestellwert",14,EURO,"d",""),
       ("Rabattsatz",11,PROZ,"f","=IF(E{r}>=500,0.1,IF(E{r}>=100,0.05,0))"),("Rabatt in €",13,EURO,"f","=E{r}*F{r}"),
       ("Zusatzbonus in €",14,EURO,"f",'=IF(AND(C{r}="Ja",E{r}>=300),E{r}*0.02,0)'),("Versandart",12,None,"f",'=IF(AND(D{r}="DE",C{r}="Ja",E{r}>=500),"Express",IF(OR(D{r}="DE",D{r}="AT"),"Standard","Ausland"))'),
       ("Zahlbetrag",14,EURO,"f","=E{r}-G{r}-H{r}")],
 rows=rm13,
 extras=[lab("A28","Summe Zahlbeträge",True),tgt("B28","=SUM(J14:J25)",EURO,"Summe Zahlbeträge"),lab("A29","Durchschnitt Zahlbetrag",True),tgt("B29","=AVERAGE(J14:J25)",EURO,"Durchschnitt Zahlbetrag"),
         lab("A30","Größter Zahlbetrag",True),tgt("B30","=MAX(J14:J25)",EURO,"Größter Zahlbetrag"),lab("A31","Expresssendungen",True),tgt("B31",'=COUNTIF(I14:I25,"Express")',ZAHL,"Expresssendungen")],
 widths={"A":24}, checks=["Grenzfälle: B-5106 (500,00 €, Bonuskunde, DE) wird Express. B-5109 (300,00 €) erhält den Zusatzbonus. B-5108 (299,99 €) nicht."])]

rnd = random.Random(11)
fil = ["Rendsburg","Kiel","Flensburg","Neumünster"]; za = ["Rechnung","Lastschrift","Vorkasse"]
zm_rows = []
for i in range(24):
    st = rnd.choices(["versendet","offen","storniert"], [70,15,15])[0]
    zm_rows.append((f"B-8{101+i}", rnd.choice(fil), rnd.choice(za), st, round(rnd.uniform(40,1500),2)))
zm_rows[7] = zm_rows[7][:3] + ("offen", None)
SPECS["zm"] = [dict(name="Erweiterung Zahlungsarten", niveau="plus", min=25, kurz="Zahlungsarten und Storno-Quote", nr="Erweiterung",
 titel="Erweiterung: Zahlungsarten und Storno-Quote je Filiale",
 sit="Die Geschäftsleitung möchte zusätzlich wissen, wie sich der Umsatz auf die Zahlungsarten verteilt und in welcher Filiale besonders viele Bestellungen storniert werden. Bei einer Bestellung fehlt der Bestellwert. Stornierte Bestellungen gehören nicht zum Umsatz.",
 auftr=["Ermitteln Sie in Spalte F den Umsatz ohne Storno (bei stornierten Bestellungen 0) und in Spalte G eine Hilfsspalte: 1 bei stornierten Bestellungen, sonst 0.",
        "Werten Sie in I14 bis I17 für jede Filiale die Anzahl der Bestellungen, in J14 bis J17 den Umsatz ohne Storno, in K14 bis K17 die Storno-Quote (Anteil stornierter Bestellungen) und in L14 bis L17 den Umsatzrang aus.",
        "Werten Sie in I21 bis I23 für jede Zahlungsart die Anzahl, in J21 bis J23 den Umsatz ohne Storno und in K21 bis K23 den Anteil am Gesamtumsatz aus.","Prüfen Sie in J24, ob die Umsätze der Zahlungsarten zusammen den Umsatz der Filialen ergeben (Text „stimmt“ oder „stimmt nicht“).", CD],
 cols=[("Bestell-Nr.",12,None,"d",""),("Filiale",13,None,"d",""),("Zahlungsart",13,None,"d",""),("Status",11,None,"d",""),("Bestellwert",13,EURO,"d",""),
       ("Umsatz ohne Storno",15,EURO,"f",'=IF(D{r}="storniert",0,E{r})'),("Storno (1/0)",10,ZAHL,"f",'=IF(D{r}="storniert",1,0)')],
 rows=zm_rows,
 extras=[lab("H13","Filiale",True),lab("I13","Anzahl",True),lab("J13","Umsatz",True),lab("K13","Storno-Quote",True),lab("L13","Rang",True)]+
   [lab(f"H{14+i}",f) for i,f in enumerate(fil)]+
   [tgt(f"I{14+i}",f"=COUNTIF($B$14:$B$37,H{14+i})",ZAHL,f"Anzahl {f}") for i,f in enumerate(fil)]+
   [tgt(f"J{14+i}",f"=SUMIF($B$14:$B$37,H{14+i},$F$14:$F$37)",EURO,f"Umsatz {f}") for i,f in enumerate(fil)]+
   [tgt(f"K{14+i}",f"=SUMIF($B$14:$B$37,H{14+i},$G$14:$G$37)/I{14+i}",PROZ,f"Storno-Quote {f}") for i,f in enumerate(fil)]+
   [tgt(f"L{14+i}",f"=RANK(J{14+i},$J$14:$J$17,0)",ZAHL,f"Rang {f}") for i,f in enumerate(fil)]+
   [lab("H20","Zahlungsart",True),lab("I20","Anzahl",True),lab("J20","Umsatz",True),lab("K20","Anteil",True)]+
   [lab(f"H{21+i}",z) for i,z in enumerate(za)]+
   [tgt(f"I{21+i}",f"=COUNTIF($C$14:$C$37,H{21+i})",ZAHL,f"Anzahl {z}") for i,z in enumerate(za)]+
   [tgt(f"J{21+i}",f"=SUMIF($C$14:$C$37,H{21+i},$F$14:$F$37)",EURO,f"Umsatz {z}") for i,z in enumerate(za)]+
   [tgt(f"K{21+i}",f"=J{21+i}/SUM($F$14:$F$37)",PROZ,f"Anteil {z}") for i,z in enumerate(za)]+
   [lab("H24","Kontrolle"),tgt("J24",'=IF(ABS(SUM(J21:J23)-SUM(J14:J17))<0.005,"stimmt","stimmt nicht")',None,"Kontrolle")],
 widths={"H":14,"I":10,"J":15,"K":13,"L":8}, checks=["Eine Bestellung hat keinen Bestellwert (leere Zelle). Sie zählt bei der Anzahl mit, bringt aber 0 € Umsatz."])]

nmp = [("A-1002",2),("A-1008",25),("A-1005",310),("A-1012",40),("A-1003",1),("A-1011",50),("A-1007",15)]
SPECS["nm"] = [dict(name="Erweiterung Skonto", niveau="plus", min=25, kurz="Zweite Rechnung mit Skonto", nr="Erweiterung",
 titel="Erweiterung: Zweite Rechnung mit Skonto und Fristen",
 sit="Für den Kunden K-4108 soll eine weitere Rechnung erstellt werden. Es gelten dieselben Stammdaten (Blatt „Stammdaten“) wie in der Hauptaufgabe. Neu: Wer innerhalb von 10 Tagen nach dem Rechnungsdatum zahlt, darf 2 % Skonto vom Rechnungsbetrag abziehen. Preise und Steuer werden kaufmännisch auf zwei Nachkommastellen gerundet.",
 auftr=["Holen Sie für jede Position Bezeichnung (Spalte C) und Einzelpreis (Spalte D) aus dem Artikelstamm und berechnen Sie in Spalte E den Gesamtpreis (gerundet).","Holen Sie in H16 die Kundengruppe des Kunden aus H15 und in H17 das Zahlungsziel dieser Gruppe. Berechnen Sie in H18 das Fälligkeitsdatum und in H19 das Ende der Skontofrist.",
        "Berechnen Sie in H20 bis H27 die Summe netto, den Rabattsatz (ungefähre Suche in der Staffel), den Rabatt, den Nettobetrag nach Rabatt, die Umsatzsteuer (19 %), den Rechnungsbetrag, den Skontobetrag und den Zahlbetrag bei Skonto.","Formatieren Sie die Rechnung nach dem Corporate Design. Datumsangaben im Format TT.MM.JJJJ."],
 cols=[("Art.-Nr.",12,None,"d",""),("Menge",9,ZAHL,"d",""),("Bezeichnung",30,None,"f","=VLOOKUP(A{r},Stammdaten!$A$4:$D$15,2,FALSE)"),("Einzelpreis netto",15,EURO,"f","=VLOOKUP(A{r},Stammdaten!$A$4:$D$15,3,FALSE)"),("Gesamtpreis netto",16,EURO,"f","=ROUND(B{r}*D{r},2)")],
 rows=nmp,
 extras=[lab("G14","Rechnungsdatum",True),inp("H14",date(2026,4,20),DATUM),lab("G15","Kunden-Nr.",True),inp("H15","K-4108",None),
   lab("G16","Kundengruppe",True),tgt("H16","=VLOOKUP(H15,Stammdaten!$F$4:$I$11,4,FALSE)",None,"Kundengruppe"),
   lab("G17","Zahlungsziel in Tagen",True),tgt("H17","=VLOOKUP(H16,Stammdaten!$N$4:$O$6,2,FALSE)",ZAHL,"Zahlungsziel in Tagen"),
   lab("G18","Fällig am",True),tgt("H18","=H14+H17",DATUM,"Fällig am"),lab("G19","Skontofrist bis",True),tgt("H19","=H14+10",DATUM,"Skontofrist bis"),
   lab("G20","Summe netto",True),tgt("H20","=SUM(E14:E20)",EURO,"Summe netto"),lab("G21","Rabattsatz",True),tgt("H21","=VLOOKUP(H20,Stammdaten!$K$4:$L$7,2,TRUE)",PROZ,"Rabattsatz"),
   lab("G22","Rabatt",True),tgt("H22","=ROUND(H20*H21,2)",EURO,"Rabatt"),lab("G23","Nettobetrag nach Rabatt",True),tgt("H23","=H20-H22",EURO,"Nettobetrag nach Rabatt"),
   lab("G24","Umsatzsteuer 19 %",True),tgt("H24","=ROUND(H23*0.19,2)",EURO,"Umsatzsteuer"),lab("G25","Rechnungsbetrag",True),tgt("H25","=H23+H24",EURO,"Rechnungsbetrag"),
   lab("G26","Skonto 2 %",True),tgt("H26","=ROUND(H25*0.02,2)",EURO,"Skonto"),lab("G27","Zahlbetrag bei Skonto",True),tgt("H27","=H25-H26",EURO,"Zahlbetrag bei Skonto")],
 widths={"G":26,"H":16}, checks=["Der Kunde gehört zur Gruppe A und hat 30 Tage Zahlungsziel."])]
