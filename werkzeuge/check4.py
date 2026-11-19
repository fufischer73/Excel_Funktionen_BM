import sys
sys.path.insert(0,"/tmp/claude-0/-home-claude/1e3cd395-f72e-5b45-8541-e91303c7c408/scratchpad")
sys.argv=["x","/tmp/o1","/tmp/o2"]
from decimal import Decimal as D, ROUND_HALF_UP
from datetime import date
import gen_stufe4 as g
from openpyxl import load_workbook
def r2(x): return D(str(x)).quantize(D("0.01"),ROUND_HALF_UP)
L="/home/claude/excel-funktionen/lehrkraft/loesungen/"
# p1
st={a:(n,D(str(p))) for a,n,p,_ in g.STAMM}
sumn=D(0)
rows=[]
for a,m in g.P1_POS:
    rab=D("0")
    for ab,s in g.MSTAFFEL:
        if m>=ab: rab=D(str(s))
    v=r2(D(m)*st[a][1]*(1-rab)); rows.append(v); sumn+=v
vers=D(0) if sumn>=200 else D("7.90")
net=sumn+vers; ust=r2(net*D("0.19")); print("p1 exp",sumn,vers,net,ust,net+ust)
ws=load_workbook(L+"p1-angebot-loesung.xlsx",data_only=True)["Angebot"]
print("p1 got",[ws[f"G{r}"].value for r in (19,20,21,22,23)],[ws[f"G{r}"].value for r in range(12,18)],[str(x) for x in rows])
# p2
ws=load_workbook(L+"p2-auswertung-loesung.xlsx",data_only=True)["Verkäufe"]
from collections import defaultdict
fs=defaultdict(D); fc=defaultdict(int); wgs=defaultdict(D)
for _,f,w,u in g.P2_DATEN: fs[f]+=D(str(u)); fc[f]+=1; wgs[w]+=D(str(u))
print("p2 exp",dict(fs),dict(fc),dict(wgs))
for r in range(4,9): print([ws.cell(row=r,column=c).value for c in range(7,14)])
for r in range(11,17): print(ws[f"G{r}"].value, ws[f"H{r}"].value)
for r in range(24,30): print([ws.cell(row=r,column=c).value for c in range(7,11)])
# p3
ws=load_workbook(L+"p3-mahnwesen-loesung.xlsx",data_only=True)["Offene Posten"]
from datetime import timedelta
st3=g.P3_STICHTAG
for (rn,k,d,z,b),r in zip(g.P3_RECH,range(7,19)):
    f=d+timedelta(days=z); t=max((st3-f).days,0)
    s="offen" if t==0 else "Zahlungserinnerung" if t<=14 else "Mahnung"
    v=(st3.year-f.year)*360+(st3.month-f.month)*30+(min(st3.day,30)-min(f.day,30)) if t>0 else 0
    z_=r2(D(str(b))*D("0.05")*v/360)
    got=[ws.cell(row=r,column=c).value for c in (6,7,8,9,10)]
    print(rn,f,t,s,v,z_,"|",got)
print([ws[f"C{r}"].value for r in range(21,28)])
t=load_workbook(L+"p3-mahnwesen-loesung.xlsx",data_only=True)["Telefonate"]
print([t[f"E{r}"].value for r in range(4,12)],t["E13"].value,t["E14"].value,t["E15"].value)
# pm
wb=load_workbook(L+"pm-gesamtaufgabe-loesung.xlsx",data_only=True)
b=wb["Bestellungen"]; a=wb["Auswertung"]
kg={k:(n,o,gr) for k,n,o,gr in g.KUNDEN}
tot=D(0); grp=defaultdict(D); ov=0; ovs=D(0); n=0
zi={"A":30,"B":20,"C":14}
for (nr,d,k,art,m),r in zip(g.PM_DATEN,range(5,29)):
    nw=r2(D(m)*st[art][1]); rab=D("0")
    for ab,s in g.RSTAFFEL:
        if nw>=ab: rab=D(str(s))
    l=r2(nw*rab); mb=nw-l; u=r2(mb*D("0.19")); br=mb+u
    gr=kg[k][2]; f=d+timedelta(days=zi[gr]); st_="überfällig" if f<g.PM_STICHTAG else "offen"
    tot+=br; grp[gr]+=br
    if st_=="überfällig": ov+=1; ovs+=br
    got=(b[f"O{r}"].value,b[f"Q{r}"].value)
    if abs(D(str(got[0]))-br)>D("0.001") or got[1]!=st_: print("DIFF",nr,br,st_,got)
print("pm exp",tot,dict(grp),ov,ovs)
print([a[f"B{r}"].value for r in range(11,18)],[a[f"C{r}"].value for r in range(4,8)],[a[f"E{r}"].value for r in range(4,7)],a["B27"].value)
