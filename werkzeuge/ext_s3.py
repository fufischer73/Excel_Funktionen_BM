from ext_framework import *
from ext_s2 import lab, tgt, inp, ans, CD
from collections import Counter
SPECS = {}
EURO4 = '#,##0.0000 "€"'

# ================================================================ n1
st1 = [("A-101","Kugelschreiber, 50er Pack","Pack",18.90),("A-102","Kopierpapier, 5 x 500 Blatt","Karton",21.50),("A-103","Ordner, 10er Pack","Pack",19.40),("A-104","Haftnotizen, 12er Pack","Pack",9.95),
       ("A-105","Tacker","Stück",7.60),("A-106","Bürostuhl Standard","Stück",129.00),("A-107","Schreibtischlampe LED","Stück",34.90),("A-108","Aktenvernichter","Stück",89.00)]
ang1 = [("A-106",4),("A-101",10),("A-104",25),("A-102",6),("A-108",1),("A-103",12),("A-107",5),("A-105",20)]
st2 = [("N-301","Briefumschläge C6, 500 Stück",24.80,5,2),("N-302","Versandkartons klein",0.95,100,4),("N-303","Versandkartons groß",1.60,100,4),("N-304","Klebeband, 6er Pack",11.90,2,2),
       ("N-305","Etiketten, 1000 Stück",8.40,10,3),("N-306","Toner schwarz",69.00,1,5),("N-307","Toner farbig, 4er Set",219.00,1,6),("N-308","USB-Stick 32 GB",6.90,10,3),
       ("N-309","Whiteboard-Marker, 4er",4.95,6,2),("N-310","Flipchartpapier, 5er",17.90,2,3)]
ang2 = [("N-302",250),("N-306",2),("N-301",3),("N-308",20),("N-310",1),("N-305",15),("N-304",4),("N-307",1)]
st3 = [("P-205","Briefpapier A4, 500 Blatt","Pack",14.90),("P-201","Kugelschreiber, 10er Pack","Pack",6.90),("P-208","Locher","Stück",12.50),("P-203","Heftklammern, 1000 Stück","Schachtel",1.80),
       ("P-206","Ablagekorb","Stück",8.90),("P-202","Textmarker, 4er Set","Set",4.60),("P-207","Schere","Stück",5.40),("P-204","Klarsichthüllen, 100 Stück","Pack",9.25)]
ang3 = [("P-203",30),("P-207",6),("P-201",20),("P-208",4),("P-202",12),("P-206",5),("P-205",8),("P-204",10)]
def nv(rows, idx, key): return {r[0]: r[idx] for r in rows}

SPECS["n1"] = [
dict(name="Aufgabe 2 Standard", niveau="basis", min=15, kurz="Angebot aus dem Artikelstamm",
 titel="Lernsituation: Angebot mit Artikelstamm erstellen",
 sit="Der Vertrieb der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, erstellt ein Angebot. Bezeichnung, Einheit und Preis stehen im Artikelstamm rechts neben der Tabelle (H13:K21). Im Angebot sind nur Artikelnummer und Menge erfasst.",
 auftr=["Holen Sie mit SVERWEIS die Bezeichnung (Spalte C), die Einheit (Spalte D) und den Einzelpreis (Spalte E) aus dem Artikelstamm. Setzen Sie die Matrix fest.","Berechnen Sie in Spalte F den Gesamtpreis jeder Position.",
        "Berechnen Sie in B24 die Angebotssumme netto und in B25 den Gesamtpreis der teuersten Position.","Ändern Sie den Preis von A-106 im Stamm auf 135,00 €. Prüfen Sie, dass sich das Angebot automatisch anpasst. Setzen Sie anschließend den alten Preis wieder ein.", CD],
 cols=[("Artikel-Nr.",12,None,"d",""),("Menge",10,ZAHL,"d",""),("Bezeichnung",28,None,"f","=VLOOKUP(A{r},$H$14:$K$21,2,FALSE)"),("Einheit",10,None,"f","=VLOOKUP(A{r},$H$14:$K$21,3,FALSE)"),
       ("Einzelpreis",14,EURO,"f","=VLOOKUP(A{r},$H$14:$K$21,4,FALSE)"),("Gesamtpreis",14,EURO,"f","=B{r}*E{r}")],
 rows=ang1,
 extras=[("table",13,8,["Artikel-Nr.","Bezeichnung","Einheit","Preis"],st1,[None,None,None,EURO]),
   lab("A24","Angebotssumme netto",True),tgt("B24","=SUM(F14:F21)",EURO,"Angebotssumme netto"),lab("A25","Teuerste Position",True),tgt("B25","=MAX(F14:F21)",EURO,"Teuerste Position")],
 widths={"A":22,"H":12,"I":28,"J":10,"K":12}, checks=["Bei A-106 steht in Zeile 14 die Bezeichnung „Bürostuhl Standard“."]),
dict(name="Aufgabe 3 Anspruchsvoll", niveau="plus", min=20, kurz="Mindestmengen und Lieferzeiten",
 titel="Lernsituation: Angebot mit Mindestmengen und Lieferzeiten",
 sit="Der Vertrieb der BüroComfort GmbH erstellt ein Angebot für Verpackungsmaterial und Verbrauchsmaterial. Der Artikelstamm steht im Blatt „Artikelstamm_3“. Er enthält Preis, Mindestmenge und Lieferzeit. Unter der Mindestmenge wird nicht geliefert.",
 auftr=["Holen Sie mit SVERWEIS aus dem Blatt „Artikelstamm_3“ die Bezeichnung (C), den Preis (D), die Mindestmenge (F) und die Lieferzeit (H).","Berechnen Sie in Spalte E den Gesamtpreis.","Zeigen Sie in Spalte G „ok“ an, wenn die Menge die Mindestmenge erreicht, sonst „unter Mindestmenge“.",
        "Berechnen Sie in B24 die Summe der Positionen, die „ok“ sind, und ermitteln Sie in B25 die längste Lieferzeit aller Positionen."],
 extra_sheets=[dict(name="Artikelstamm_3", titel="BüroComfort GmbH – Artikelstamm Verbrauchsmaterial", headers=["Artikel-Nr.","Bezeichnung","Preis","Mindestmenge","Lieferzeit (Tage)"],
                     rows=st2, fmts=[None,None,EURO,ZAHL,ZAHL], widths=[14,32,12,16,18])],
 cols=[("Artikel-Nr.",12,None,"d",""),("Menge",10,ZAHL,"d",""),("Bezeichnung",30,None,"f","=VLOOKUP(A{r},Artikelstamm_3!$A$4:$E$13,2,FALSE)"),("Preis",12,EURO,"f","=VLOOKUP(A{r},Artikelstamm_3!$A$4:$E$13,3,FALSE)"),
       ("Gesamtpreis",14,EURO,"f","=B{r}*D{r}"),("Mindestmenge",14,ZAHL,"f","=VLOOKUP(A{r},Artikelstamm_3!$A$4:$E$13,4,FALSE)"),("Prüfung",20,None,"f",'=IF(B{r}>=F{r},"ok","unter Mindestmenge")'),
       ("Lieferzeit (Tage)",16,ZAHL,"f","=VLOOKUP(A{r},Artikelstamm_3!$A$4:$E$13,5,FALSE)")],
 rows=ang2,
 extras=[lab("A24","Summe der Positionen mit „ok“",True),tgt("B24",'=SUMIF(G14:G21,"ok",E14:E21)',EURO,"Summe der Positionen mit „ok“"),lab("A25","Längste Lieferzeit",True),tgt("B25","=MAX(H14:H21)",ZAHL,"Längste Lieferzeit")],
 widths={"A":28}, checks=["Zwei Positionen liegen unter der Mindestmenge: N-301 und N-310."]),
dict(name="Aufgabe 4 Transfer", niveau="plus", min=20, kurz="Fehlersuche im Angebot",
 titel="Lernsituation: Fehlerhafte Angebotstabelle reparieren",
 sit="Eine Auszubildende hat ein Angebot mit SVERWEIS erstellt (Artikelstamm im Blatt „Artikelstamm_4“). Die Tabelle zeigt Fehlermeldungen und falsche Werte. Die Fehler liegen in den Formeln, nicht in den Daten.",
 auftr=["Untersuchen Sie die Formeln in den Spalten C bis F. Jede Spalte enthält einen anderen Fehler.","Korrigieren Sie die Formeln in Zeile 14 und kopieren Sie sie in die Zeilen darunter.","Notieren Sie in B24 in einem Satz, welche vier Fehlerarten Sie gefunden haben.","Berechnen Sie in B25 die Angebotssumme netto."],
 extra_sheets=[dict(name="Artikelstamm_4", titel="BüroComfort GmbH – Artikelstamm Büroartikel", headers=["Artikel-Nr.","Bezeichnung","Einheit","Preis"],rows=st3,fmts=[None,None,None,EURO],widths=[14,32,12,12])],
 cols=[("Artikel-Nr.",12,None,"d",""),("Menge",10,ZAHL,"d",""),
       ("Bezeichnung",28,None,"x","=VLOOKUP(A{r},Artikelstamm_4!$A$4:$D$11,2,FALSE)",lambda r: f"=VLOOKUP(A{r},Artikelstamm_4!A{r-10}:D{r-3},2,FALSE)","Die Matrix ist nicht mit Dollarzeichen festgesetzt und wandert beim Kopieren nach unten mit. Folge: Artikel am Anfang des Stamms werden nicht mehr gefunden (#NV)."),
       ("Einheit",12,None,"x","=VLOOKUP(A{r},Artikelstamm_4!$A$4:$D$11,3,FALSE)",lambda r: f"=VLOOKUP(A{r},Artikelstamm_4!$A$4:$D$11,5,FALSE)","Der Spaltenindex 5 ist größer als die Zahl der Matrixspalten (4). Folge: #BEZUG!. Die Einheit steht in der dritten Spalte."),
       ("Einzelpreis",14,EURO,"x","=VLOOKUP(A{r},Artikelstamm_4!$A$4:$D$11,4,FALSE)",lambda r: f"=VLOOKUP(A{r},Artikelstamm_4!$A$4:$D$11,4)","Das letzte Argument fehlt. Excel sucht dann ungefähr und erwartet einen aufsteigend sortierten Stamm. Der Stamm ist nicht sortiert, deshalb entstehen falsche Preise oder #NV. Richtig ist FALSCH für einen genauen Treffer."),
       ("Gesamtpreis",14,EURO,"x","=B{r}*E{r}",lambda r: f"=B{r}*D{r}","Die Formel multipliziert mit der Spalte D (Einheit, Text) statt mit dem Einzelpreis in Spalte E. Folge: #WERT!.")],
 rows=ang3,
 extras=[lab("A24","Gefundene Fehler",True),ans("B24","Matrix nicht festgesetzt, Spaltenindex zu groß, Bereich_Verweis (FALSCH) fehlt, falsche Spalte im Produkt.",None,"B24:F24"),lab("A25","Angebotssumme netto",True),tgt("B25","=SUM(F14:F21)",EURO,"Angebotssumme netto")],
 widths={"A":22}, checks=["Nach der Korrektur steht in keiner Zelle der Spalten C bis F eine Fehlermeldung."]),
]

# ================================================================ n2
ks = [("K-501","Beispielbau GmbH","Rendsburg","A"),("K-502","Muster-Elektro GmbH","Kiel","B"),("K-503","Test-Gastro e. K.","Eckernförde","C"),("K-504","Demo-Möbel GmbH","Rendsburg","A"),
      ("K-505","Probe-Druck OHG","Büdelsdorf","B"),("K-506","Beispiel-Spedition AG","Kiel","A"),("K-507","Modell-Handel KG","Flensburg","C"),("K-508","Muster-Verlag GmbH","Rendsburg","B")]
bes2 = [("B-7001","K-504",640.0),("B-7002","K-502",1280.5),("B-7003","K-501",95.9),("B-7004","K-507",310.0),("B-7005","K-506",2250.0),("B-7006","K-503",49.9),("B-7007","K-508",780.0),("B-7008","K-505",1520.0),("B-7009","K-501",430.0),("B-7010","K-502",205.4)]
bes2b = [("B-8001","A",120.0),("B-8002","B",499.99),("B-8003","C",500.0),("B-8004","A",999.99),("B-8005","B",1000.0),("B-8006","C",2500.0),("B-8007","A",3200.0),("B-8008","C",65.5),("B-8009","B",1750.0),("B-8010","A",640.0)]
bes2c = [("B-9001","Beispielbau GmbH",24768,150.0),("B-9002","Muster-Elektro GmbH",24340,149.9),("B-9003","Test-Gastro e. K.",24103,89.5),("B-9004","Demo-Möbel GmbH",25746,310.0),("B-9005","Probe-Druck OHG",20095,64.9),
         ("B-9006","Beispiel-Spedition AG",24937,1200.0),("B-9007","Modell-Handel KG",23552,45.0),("B-9008","Muster-Verlag GmbH",24534,175.0),("B-9009","Beispiel-Bau & Garten e. K.",26121,98.0),("B-9010","Test-Reisen GmbH",10115,210.0)]
SPECS["n2"] = [
dict(name="Aufgabe 2 Standard", niveau="basis", min=15, kurz="Kundengruppe und Rabatt holen",
 titel="Lernsituation: Bestellungen mit Kundenstamm und Gruppenrabatt",
 sit="Die Verwaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, soll zu jeder Bestellung Firma, Ort und Kundengruppe aus dem Kundenstamm (K13:N21) holen. Der Rabatt und das Zahlungsziel hängen von der Kundengruppe ab (Gruppentabelle in P13:R16).",
 auftr=["Holen Sie in den Spalten D, E und F Firma, Ort und Kundengruppe aus dem Kundenstamm.","Holen Sie in Spalte G den Rabattsatz und in Spalte H das Zahlungsziel aus der Gruppentabelle. Das Suchkriterium ist die Gruppe aus Spalte F.",
        "Berechnen Sie in Spalte I den Rechnungsbetrag nach Abzug des Rabatts.","Berechnen Sie in B24 die Summe der Rechnungsbeträge und ermitteln Sie in B25 das längste Zahlungsziel.", CD],
 cols=[("Bestell-Nr.",12,None,"d",""),("Kunden-Nr.",12,None,"d",""),("Warenwert",14,EURO,"d",""),("Firma",24,None,"f","=VLOOKUP(B{r},$K$14:$N$21,2,FALSE)"),("Ort",14,None,"f","=VLOOKUP(B{r},$K$14:$N$21,3,FALSE)"),
       ("Gruppe",9,None,"f","=VLOOKUP(B{r},$K$14:$N$21,4,FALSE)"),("Rabattsatz",11,PROZ,"f","=VLOOKUP(F{r},$P$14:$R$16,2,FALSE)"),("Zahlungsziel (Tage)",13,ZAHL,"f","=VLOOKUP(F{r},$P$14:$R$16,3,FALSE)"),("Rechnungsbetrag",15,EURO,"f","=C{r}*(1-G{r})")],
 rows=bes2,
 extras=[("table",13,11,["Kunden-Nr.","Firma","Ort","Gruppe"],ks,None),("table",13,16,["Gruppe","Rabattsatz","Zahlungsziel (Tage)"],[("A",0.10,30),("B",0.05,20),("C",0.0,10)],[None,PROZ,ZAHL]),
         lab("A24","Summe Rechnungsbeträge",True),tgt("B24","=SUM(I14:I23)",EURO,"Summe Rechnungsbeträge"),lab("A25","Längstes Zahlungsziel",True),tgt("B25","=MAX(H14:H23)",ZAHL,"Längstes Zahlungsziel")],
 widths={"K":12,"L":24,"M":14,"N":9,"O":3,"P":9,"Q":12,"R":14,"A":24}, checks=["Kunde K-503 (Gruppe C) erhält keinen Rabatt und 10 Tage Zahlungsziel."]),
dict(name="Aufgabe 3 Anspruchsvoll", niveau="plus", min=20, kurz="Rabattstaffel mit ungefährer Suche",
 titel="Lernsituation: Mengenrabatt und Gruppenrabatt kombinieren",
 sit="Der Rabatt setzt sich aus zwei Teilen zusammen: dem Mengenrabatt nach Warenwert (Staffel in J13:K17) und einem Gruppenrabatt nach Kundengruppe (Tabelle in M13:N16). Zusammen darf der Rabatt den Höchstrabatt in K19 nicht überschreiten.",
 auftr=["Ermitteln Sie in Spalte D den Mengenrabatt mit SVERWEIS und ungefährer Suche (WAHR). Der Warenwert ist das Suchkriterium.","Holen Sie in Spalte E den Gruppenrabatt mit SVERWEIS und genauer Suche (FALSCH).",
        "Berechnen Sie in Spalte F den Gesamtrabatt: Summe beider Teile, höchstens der Höchstrabatt aus K19.","Berechnen Sie in Spalte G den Rabatt in Euro und in Spalte H den Rechnungsbetrag.",
        "Prüfen Sie die Grenzfälle 499,99 €, 500,00 €, 999,99 € und 1.000,00 €. Begründen Sie in B25, warum die ungefähre Suche eine aufsteigend sortierte Staffel braucht.", CD],
 cols=[("Bestell-Nr.",12,None,"d",""),("Gruppe",9,None,"d",""),("Warenwert",14,EURO,"d",""),("Mengenrabatt",13,PROZ,"f","=VLOOKUP(C{r},$J$14:$K$17,2,TRUE)"),("Gruppenrabatt",13,PROZ,"f","=VLOOKUP(B{r},$M$14:$N$16,2,FALSE)"),
       ("Gesamtrabatt",13,PROZ,"f","=MIN(D{r}+E{r},$K$19)"),("Rabatt in €",13,EURO,"f","=C{r}*F{r}"),("Rechnungsbetrag",15,EURO,"f","=C{r}-G{r}")],
 rows=bes2b,
 extras=[("table",13,10,["Warenwert ab","Mengenrabatt"],[(0,0.0),(500,0.02),(1000,0.04),(2500,0.06)],[EURO,PROZ]),("table",13,13,["Gruppe","Gruppenrabatt"],[("A",0.04),("B",0.02),("C",0.0)],[None,PROZ]),
         lab("J19","Höchstrabatt",True),inp("K19",0.08,PROZ),
         lab("A25","Begründung",True),ans("B25","Bei ungefährer Suche liefert SVERWEIS den größten Wert, der kleiner oder gleich dem Suchkriterium ist. Das funktioniert nur, wenn die erste Spalte aufsteigend sortiert ist.",None,"B25:H25")],
 widths={"J":14,"K":14,"L":3,"M":9,"N":14,"A":14}, checks=["500,00 € gehört schon zur Stufe ab 500 €, 499,99 € noch nicht.","Bei B-8007 (Gruppe A, 3.200 €) ergibt die Summe 10 %, der Höchstrabatt begrenzt auf 8 %."]),
dict(name="Aufgabe 4 Transfer", niveau="plus", min=20, kurz="Versandzonen nach Postleitzahl",
 titel="Lernsituation: Versandkosten nach Postleitzahl-Zone",
 sit="Die BüroComfort GmbH berechnet Versandkosten nach Zonen. Die Zone hängt von der Postleitzahl ab (Tabelle in J13:K18), die Kosten hängen von der Zone ab (Tabelle in M13:N16). Ab einem Warenwert von 150,00 € (K20) ist der Versand frei. Die Zonen sind fiktiv.",
 auftr=["Ermitteln Sie in Spalte E die Zone mit einer ungefähren Suche (WAHR). Die Postleitzahl ist das Suchkriterium.","Holen Sie in Spalte F die Versandkosten der Zone mit einer genauen Suche (FALSCH).",
        "Berechnen Sie in Spalte G die berechneten Versandkosten: 0,00 € bei einem Warenwert ab der Freigrenze, sonst die Kosten aus Spalte F.","Berechnen Sie in Spalte H die Rechnungssumme.",
        "Berechnen Sie in B24 die Versandkosten insgesamt und zählen Sie in B25 die Bestellungen in Zone 3."],
 cols=[("Bestell-Nr.",12,None,"d",""),("Kunde",28,None,"d",""),("PLZ",9,"0","d",""),("Warenwert",13,EURO,"d",""),("Zone",8,ZAHL,"f","=VLOOKUP(C{r},$J$14:$K$18,2,TRUE)"),
       ("Versandkosten Zone",14,EURO,"f","=VLOOKUP(E{r},$M$14:$N$16,2,FALSE)"),("Berechnete Versandkosten",16,EURO,"f","=IF(D{r}>=$K$20,0,F{r})"),("Rechnungssumme",15,EURO,"f","=D{r}+G{r}")],
 rows=bes2c,
 extras=[("table",13,10,["PLZ ab","Zone"],[(0,3),(20000,2),(24000,1),(25000,2),(26000,3)],["0",ZAHL]),("table",13,13,["Zone","Versandkosten"],[(1,4.90),(2,6.90),(3,9.90)],[ZAHL,EURO]),
         lab("J20","Frei ab Warenwert",True),inp("K20",150,EURO),
         lab("A24","Versandkosten gesamt",True),tgt("B24","=SUM(G14:G23)",EURO,"Versandkosten gesamt"),lab("A25","Bestellungen in Zone 3",True),tgt("B25","=COUNTIF(E14:E23,3)",ZAHL,"Bestellungen in Zone 3")],
 widths={"J":10,"K":14,"L":3,"M":8,"N":14,"A":24}, checks=["PLZ 25746 liegt in Zone 2, PLZ 24937 in Zone 1."]),
]

# ================================================================ n3
n3a = [("Kugelschreiber",7,3.5),("Kopierpapier",12,4.2),("Haftnotizen",9,2.81),("Ordner",15,1.91),("Tacker",4,6.37),("Heftklammern",25,0.9),("Textmarker",18,1.16),("Locher",3,9.5)]
ust = sum(round(b*m*c*0.19,2) for _,b,c in [(a,m,c) for a,m,c in n3a for b in [m]] ) if False else None
rnd_sum = round(sum(round(m*c*0.19,2) for _,m,c in n3a),2); rnd_tot = round(round(sum(m*c for _,m,c in n3a),2)*0.19,2)
n3b = [("Briefumschläge",1275,250),("Versandkartons",830,40),("Etiketten",4600,1000),("Klebeband",95,6),("Ordner",230,12),("Kopierpapier",1960,40),("Toner",19,4),("Haftnotizen",730,24)]
vk = [("Kugelschreiber",1.62),("Tacker",3.98),("Ordner",1.74),("Bürostuhl",84.30),("Schreibtischlampe",18.85),("Aktenvernichter",52.10)]
SPECS["n3"] = [
dict(name="Aufgabe 2 Standard", niveau="basis", min=15, kurz="Umsatzsteuer und Rundungsdifferenz",
 titel="Lernsituation: Umsatzsteuer berechnen und runden",
 sit="Die Buchhaltung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, berechnet auf einer Rechnung die Umsatzsteuer. Der Steuersatz steht in Zelle I15. Bei einer Rechnung mit mehreren Positionen kann es einen Unterschied machen, ob je Position oder auf die Gesamtsumme gerundet wird.",
 auftr=["Berechnen Sie in Spalte D den Nettobetrag jeder Position.","Berechnen Sie in Spalte E die Umsatzsteuer ungerundet und in Spalte F die Umsatzsteuer auf zwei Nachkommastellen gerundet. Spalte G enthält den Bruttobetrag mit der gerundeten Steuer.",
        "Bilden Sie in Zeile 22 die Summen der Spalten D bis G.","Berechnen Sie in B23 die Umsatzsteuer aus der Summe der Nettobeträge, gerundet auf zwei Nachkommastellen, und in B24 die Differenz zur Summe der Einzelrundungen (F22).",
        "Beantworten Sie in B25 in einem Satz, warum auf einer Rechnung nur die Rundung auf Cent zulässig ist, obwohl intern mit mehr Nachkommastellen gerechnet wird.", CD],
 cols=[("Artikel",18,None,"d",""),("Menge",9,ZAHL,"d",""),("Nettopreis",12,EURO,"d",""),("Nettobetrag",13,EURO,"f","=B{r}*C{r}"),("USt ungerundet",14,EURO4,"f","=D{r}*$I$15"),
       ("USt gerundet",13,EURO,"f","=ROUND(E{r},2)"),("Bruttobetrag",14,EURO,"f","=D{r}+F{r}")],
 rows=n3a,
 extras=[lab("H14","Steuersatz",True),inp("I15",0.19,PROZ),lab("A22","Summe",True)]+
   [tgt(f"{c}22",f"=SUM({c}14:{c}21)",f,f"Summe {n}",True) for c,f,n in (("D",EURO,"Netto"),("E",EURO4,"USt ungerundet"),("F",EURO,"USt gerundet"),("G",EURO,"Brutto"))]+
   [lab("A23","USt aus Gesamtnetto",True),tgt("B23","=ROUND(D22*$I$15,2)",EURO,"USt aus Gesamtnetto"),lab("A24","Differenz",True),tgt("B24","=F22-B23",EURO,"Differenz zur Summe der Einzelrundungen"),
    lab("A25","Begründung",True),ans("B25","Beträge in Euro können nur in ganzen Cent gezahlt werden. Deshalb muss jede ausgewiesene Zahl auf zwei Nachkommastellen gerundet sein.",None,"B25:G25")],
 widths={"H":14,"I":10,"A":22}, checks=["Die Rundung je Position und die Rundung der Gesamtsumme führen zu unterschiedlichen Steuerbeträgen. Die Differenz in B24 ist nicht null."]),
dict(name="Aufgabe 3 Transfer", niveau="plus", min=20, kurz="Kartons und Verkaufspreise",
 titel="Lernsituation: Kartons berechnen und Verkaufspreise festlegen",
 sit="Das Lager der BüroComfort GmbH verpackt Bestellungen in Kartons, die nur voll gefüllt oder angebrochen verschickt werden. Außerdem sollen Verkaufspreise festgelegt werden, die immer auf ,99 enden.",
 auftr=["Berechnen Sie in Spalte D die benötigten Kartons. Angebrochene Kartons zählen als ganze Kartons.","Berechnen Sie in Spalte E die Zahl der voll gefüllten Kartons (angebrochene werden nicht gezählt) und in Spalte F die freien Plätze im letzten Karton.",
        "Ermitteln Sie in B23 die Gesamtzahl der benötigten Kartons.","Berechnen Sie in J14 bis J19 den Verkaufspreis: Einkaufspreis mal Aufschlagfaktor (I20), auf den nächsten vollen Euro aufgerundet, abzüglich 0,01 €."],
 cols=[("Artikel",18,None,"d",""),("Menge (Stück)",14,ZAHL,"d",""),("Stück je Karton",14,ZAHL,"d",""),("Kartons",10,ZAHL,"f","=ROUNDUP(B{r}/C{r},0)"),("Volle Kartons",13,ZAHL,"f","=ROUNDDOWN(B{r}/C{r},0)"),("Freie Plätze",12,ZAHL,"f","=D{r}*C{r}-B{r}")],
 rows=n3b,
 extras=[lab("A23","Kartons insgesamt",True),tgt("B23","=SUM(D14:D21)",ZAHL,"Kartons insgesamt"),
   ("table",13,8,["Artikel","Einkaufspreis"],vk,[None,EURO]),lab("J13","Verkaufspreis",True),lab("H20","Aufschlagfaktor",True),inp("I20",1.45,"0.00")]+
   [tgt(f"J{14+i}",f"=ROUNDUP(I{14+i}*$I$20,0)-0.01",EURO,f"Verkaufspreis {a}") for i,(a,_) in enumerate(vk)],
 widths={"H":20,"I":14,"J":14,"A":22}, checks=["Alle Verkaufspreise enden auf ,99.","Ändern Sie den Faktor in I20 auf 1,50: Die Preise passen sich an."]),
]

# ================================================================ n4
t4 = [("Verkaufsteam 1",18200,19850),("Verkaufsteam 2",21400,18400),("Verkaufsteam 3",16750,17900),("Verkaufsteam 4",19900,21600),("Verkaufsteam 5",22800,22150),
      ("Verkaufsteam 6",17300,18400),("Verkaufsteam 7",20100,19200),("Verkaufsteam 8",15600,16800)]
def rk(v, vs, asc=False):
    return 1 + sum(1 for x in vs if (x < v if asc else x > v))
r_a = [rk(a, [t[2] for t in t4]) for _,_,a in t4]; r_v = [rk(v, [t[1] for t in t4]) for _,v,_ in t4]
besser = sum(1 for a,b in zip(r_v,r_a) if a-b>0)
zt = [("Mitarbeit 1",12),("Mitarbeit 2",18),("Mitarbeit 3",15),("Mitarbeit 4",18),("Mitarbeit 5",25),("Mitarbeit 6",9),("Mitarbeit 7",21),("Mitarbeit 8",30),("Mitarbeit 9",15),("Mitarbeit 10",27)]
SPECS["n4"] = [
dict(name="Aufgabe 2 Standard", niveau="basis", min=12, kurz="Teams im Vergleich",
 titel="Lernsituation: Ranking der Verkaufsteams",
 sit="Die Geschäftsleitung der BüroComfort GmbH, Buntstraße 11, 24768 Rendsburg, vergleicht die Umsätze der acht Verkaufsteams im Vormonat und im aktuellen Monat. Der höchste Umsatz erhält Rang 1.",
 auftr=["Ermitteln Sie in Spalte D den Rang im Vormonat und in Spalte E den Rang im aktuellen Monat.","Berechnen Sie in Spalte F die Platzveränderung (positiv = Verbesserung) und in Spalte G die Umsatzentwicklung in Prozent.",
        "Zählen Sie in B23, wie viele Teams sich im Rang verbessert haben.","Beantworten Sie in B24: Welche Teams teilen sich einen Rang, und welcher Rang entfällt dadurch?", CD],
 cols=[("Team",18,None,"d",""),("Umsatz Vormonat",16,EURO,"d",""),("Umsatz aktuell",16,EURO,"d",""),("Rang Vormonat",14,ZAHL,"f","=RANK(B{r},$B$14:$B$21,0)"),("Rang aktuell",13,ZAHL,"f","=RANK(C{r},$C$14:$C$21,0)"),
       ("Platzveränderung",15,ZAHL,"f","=D{r}-E{r}"),("Entwicklung",12,PROZ,"f","=C{r}/B{r}-1")],
 rows=t4,
 extras=[lab("A23","Verbesserte Teams",True),tgt("B23",'=COUNTIF(F14:F21,">0")',ZAHL,"Verbesserte Teams"),lab("A24","Gleicher Rang",True),
         ans("B24","Verkaufsteam 2 und Verkaufsteam 6 haben beide 18.400 € und damit Rang 5. Rang 6 entfällt.",None,"B24:G24")],
 widths={"A":22}, checks=["Bei Gleichstand erhalten beide Teams denselben Rang, der nächste Rang wird übersprungen."]),
dict(name="Aufgabe 3 Transfer", niveau="plus", min=15, kurz="Bearbeitungszeiten und Prämie",
 titel="Lernsituation: Schnellste Bearbeitung belohnen",
 sit="Im Kundenservice der BüroComfort GmbH wird die durchschnittliche Bearbeitungszeit je Anfrage in Minuten gemessen. Die schnellsten drei Ränge erhalten eine Prämie: Rang 1 erhält 100 €, die Ränge 2 und 3 erhalten je 50 €, alle anderen keine Prämie. Bei Gleichstand erhalten alle Beteiligten die Prämie des Rangs.",
 auftr=["Ermitteln Sie in Spalte C den Rang. Die kürzeste Zeit ist Rang 1 (aufsteigende Reihenfolge).","Zeigen Sie in Spalte D „Top 3“ an, wenn der Rang höchstens 3 ist, sonst „-“.","Berechnen Sie in Spalte E die Prämie mit zwei verschachtelten WENN-Funktionen.",
        "Berechnen Sie in B25 die Summe aller Prämien. Beantworten Sie in B26, wie viele Personen zu den Top 3 gehören und warum es mehr als drei sind."],
 cols=[("Mitarbeit",16,None,"d",""),("Zeit (Minuten)",14,ZAHL,"d",""),("Rang",8,ZAHL,"f","=RANK(B{r},$B$14:$B$23,1)"),("Top 3?",10,None,"f",'=IF(C{r}<=3,"Top 3","-")'),("Prämie",12,EURO,"f","=IF(C{r}=1,100,IF(C{r}<=3,50,0))")],
 rows=zt,
 extras=[lab("A25","Summe der Prämien",True),tgt("B25","=SUM(E14:E23)",EURO,"Summe der Prämien"),lab("A26","Anzahl Top 3",True),
         ans("B26","4 Personen, weil Mitarbeit 3 und Mitarbeit 9 beide 15 Minuten brauchen und gemeinsam Rang 3 teilen.",None,"B26:E26")],
 widths={"A":20}, checks=["Rang 4 wird nicht vergeben."]),
]

# ================================================================ n5
n5a = [("Projektstart", date(2026,3,2), 14), ("Konzept", date(2026,3,16), 21), ("Beschaffung", date(2026,4,6), 28), ("Aufbau", date(2026,5,4), 19),
       ("Testlauf", date(2026,5,26), 12), ("Schulung", date(2026,6,8), 17), ("Abnahme", date(2026,6,25), 8), ("Nachbetreuung", date(2026,7,6), 24)]
n5b = [("Montag",time(7,30),time(16,0)),("Dienstag",time(8,0),time(18,0)),("Mittwoch",time(8,0),time(13,30)),("Donnerstag",time(7,45),time(15,45)),("Freitag",time(9,0),time(15,0))]
SPECS["n5"] = [
dict(name="Aufgabe 6 Anspruchsvoll", niveau="plus", min=20, kurz="Projektplan mit Stichtag", nr=6,
 titel="Lernsituation: Projektplan mit Fristen überwachen",
 sit="Die Abteilung Verwaltung der BüroComfort GmbH plant die Einführung einer neuen Lagersoftware in acht Abschnitten. Zu jedem Abschnitt sind Start und Dauer in Kalendertagen bekannt. Der Stichtag der Auswertung steht in Zelle K15.",
 auftr=["Berechnen Sie in Spalte D das Ende jedes Abschnitts (Start plus Dauer).","Berechnen Sie in Spalte E, wie viele Tage das Ende vom Stichtag entfernt ist (negativ = bereits abgelaufen).",
        "Zeigen Sie in Spalte F den Status: „abgelaufen“ bei negativem Wert, „läuft bald ab“ bis einschließlich 7 Tage, sonst „im Plan“.","Ermitteln Sie in Spalte G mit einer 30-Tage-Methode die Laufzeit zwischen Start und Ende in Tagen.",
        "Zerlegen Sie in Spalte H den Monat des Endes. Zählen Sie in B24 mit ZÄHLENWENN die abgelaufenen Abschnitte."],
 cols=[("Abschnitt",18,None,"d",""),("Start",13,DATUM,"d",""),("Dauer (Tage)",12,ZAHL,"d",""),("Ende",13,DATUM,"f","=B{r}+C{r}"),("Tage bis Stichtag",14,ZAHL,"f","=D{r}-$K$15"),
       ("Status",16,None,"f",'=IF(E{r}<0,"abgelaufen",IF(E{r}<=7,"läuft bald ab","im Plan"))'),("Laufzeit (30-Tage-Methode)",16,ZAHL,"f","=DAYS360(B{r},D{r})"),("Monat des Endes",12,ZAHL,"f","=MONTH(D{r})")],
 rows=n5a,
 extras=[lab("J14","Stichtag",True),inp("K15",date(2026,5,16),DATUM),lab("A24","Abgelaufene Abschnitte",True),tgt("B24",'=COUNTIF(F14:F21,"abgelaufen")',ZAHL,"Abgelaufene Abschnitte")],
 widths={"A":24,"J":14,"K":14}, checks=["Der Stichtag ist fest eingetragen. Ändern Sie ihn auf ein anderes Datum, damit der Status neu berechnet wird."]),
dict(name="Aufgabe 7 Transfer", niveau="plus", min=20, kurz="Arbeitszeit mit Pausenregel", nr=7,
 titel="Lernsituation: Arbeitszeit mit gesetzlicher Pausenregel",
 sit="In der Zeiterfassung der BüroComfort GmbH wird die Pause nach der Länge des Arbeitstages berechnet: Bei mehr als 6 Stunden Brutto-Arbeitszeit 30 Minuten Pause, bei mehr als 9 Stunden 45 Minuten Pause, sonst keine. Der Stundenlohn steht in Zelle K15.",
 auftr=["Berechnen Sie in Spalte D die Brutto-Arbeitszeit (Ende minus Beginn).","Bestimmen Sie in Spalte E die Pause nach der Regel. Rechnen Sie mit Bruchteilen eines Tages, zum Beispiel 30/1440 für 30 Minuten.",
        "Berechnen Sie in Spalte F die Netto-Arbeitszeit und in Spalte G die Netto-Arbeitszeit in Dezimalstunden (ganze Stunden plus Minuten geteilt durch 60).","Berechnen Sie in Spalte H den Lohn des Tages.",
        "Berechnen Sie in Zeile 19 die Summen der Dezimalstunden und des Lohns. Prüfen Sie den Freitag mit genau 6 Stunden: Hier gibt es noch keine Pause."],
 cols=[("Tag",14,None,"d",""),("Beginn",10,UHR,"d",""),("Ende",10,UHR,"d",""),("Brutto",10,UHR,"f","=C{r}-B{r}"),("Pause",10,UHR,"f","=IF(D{r}>9/24,45/1440,IF(D{r}>6/24,30/1440,0))"),("Netto",10,UHR,"f","=D{r}-E{r}"),
       ("Dezimalstunden",14,"0.00","f","=HOUR(F{r})+MINUTE(F{r})/60"),("Lohn",12,EURO,"f","=G{r}*$K$15")],
 rows=n5b,
 extras=[lab("J14","Stundenlohn",True),inp("K15",16.0,EURO),lab("A19","Summe",True),tgt("G19","=SUM(G14:G18)","0.00","Summe Dezimalstunden",True),tgt("H19","=SUM(H14:H18)",EURO,"Summe Lohn",True)],
 widths={"J":14,"K":12}, checks=["Am Freitag (genau 6:00 Stunden brutto) wird keine Pause abgezogen."]),
]
