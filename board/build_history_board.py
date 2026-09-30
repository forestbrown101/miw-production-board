import build_production_json as B
import csv, io, json, glob, os, re, datetime
from collections import defaultdict

DR=os.environ.get("DR","/sessions/charming-tender-noether/mnt/Daily Reports")
DATA=os.environ.get("DATA","/sessions/charming-tender-noether/mnt/MIW Projects/PowerFab_API/data")
CUT_CSV=os.environ.get("CUT_CSV",os.path.join(DATA,"cut_by_saw_history.csv"))
PAINT_CSV=os.environ.get("PAINT_CSV",os.path.join(DATA,"paint_by_plant_history.csv"))
# THE CANONICAL FEED: the same two files the WEEK ## SCHEDULES workbook reads.
FEED_FIT=os.environ.get("FEED_FIT",""); FEED_CUT=os.environ.get("FEED_CUT","")
NAMES=os.environ.get("NAMES","")
OUT=os.environ.get("OUT","/sessions/charming-tender-noether/mnt/outputs")
today=datetime.date.today()
d0=datetime.date.fromisoformat

def load_fit():
    rows=[]
    for r in csv.DictReader(open(f"{DATA}/qc_history.csv",encoding="utf-8")):
        if r.get("Test")!="VT - Final Visual Weld Inspection": continue
        if int(float(r.get("Failed") or 0))!=0: continue
        try: dt=d0(r["Date"][:10])
        except: continue
        p=B.FIT_LOC_TO_PLANT.get(r.get("Plant"))
        if not p: continue
        j=None
        m=re.search(r"(\d{3,5})",str(r.get("Job") or ""))
        if m: j=int(m.group(1))
        rows.append({"dt":dt,"p":p,"t":float(r.get("Tons") or 0),"j":j,
                     "pc":int(float(r.get("Quantity") or 0))})
    return rows

def load_hours():
    rows=[]
    for fn in glob.glob(os.path.join(DR,"QuickBooks Time","*.csv")):
        try: raw=open(fn,encoding="utf-8",errors="replace").read().replace("\x00","")
        except: continue
        try: rdr=csv.DictReader(io.StringIO(raw))
        except: continue
        for row in rdr:
            try: dt=d0(row["local_date"])
            except: continue
            grp=row.get("group"); j=None
            m=re.search(r"(\d{3,5})",str(row.get("jobcode_2") or ""))
            if m: j=int(m.group(1))
            if j is None:
                code=((row.get("jobcode_1") or "")+" "+(row.get("jobcode_2") or "")).lower()
                if any(k in code for k in B.NONWORK_KEYS): continue
            rows.append({"dt":dt,"p":B.tile_plant(grp),"po":B.overall_plant(grp),
                         "h":B.numf(row.get("hours")),"j":j})
    return rows

def load_cut_total():
    cd=defaultdict(float)
    for r in csv.DictReader(open(f"{DATA}/station_history.csv",encoding="utf-8")):
        if r.get("Station")=="Cut":
            try: cd[d0(r["StartDate"][:10])]+=float(r.get("Tons") or 0)
            except: pass
    return cd

def load_paint_total():
    pd=defaultdict(lambda:[0.0,0])
    for r in csv.DictReader(open(f"{DATA}/station_history.csv",encoding="utf-8")):
        if r.get("Station")=="Paint":
            try: k=d0(r["StartDate"][:10])
            except: continue
            pd[k][0]+=float(r.get("Tons") or 0); pd[k][1]+=int(float(r.get("Quantity") or 0))
    return pd
# Prefer the canonical feed. load_fit()/read_saw_detail_api remain as the fallback.
# The feed already resolves Plant (C-Beams and SHOP -> P1, Shop 3 -> P3) and already
# drops P3 when SHOW_P3 is False, so never re-derive plant from Test Location here.
if FEED_FIT and os.path.exists(FEED_FIT):
    fit=B.load_feed_fit(FEED_FIT); print("fit source: feed_fit.csv")
else:
    fit=load_fit(); print("fit source: qc_history.csv (feed not found)")
if FEED_CUT and os.path.exists(FEED_CUT):
    detail=B.read_cut_feed(FEED_CUT); print("cut source: feed_cut.csv")
else:
    detail=B.read_saw_detail_api(CUT_CSV); print("cut source: cut_by_saw_history.csv (feed not found)")
hrs=load_hours(); cutday=load_cut_total(); paintday=load_paint_total()
# Paint by plant (Plant 2 paint shop, 2026-09-15). Dates in this file are split P1/P2;
# dates missing from it fall back to the station total above as Plant 1.
paint_bp=B.read_paint_by_plant(PAINT_CSV)
print("paint split source:", PAINT_CSV if paint_bp else "(none yet, all Plant 1)", "| split dates", len(paint_bp))
names=B.load_names(NAMES)
alldates=sorted({r["dt"] for r in fit})
print("fit dates",alldates[0],"to",alldates[-1],"| saw detail dates",len(detail))

def period_cut(ps,pe):
    """Returns (P1, P4, unattributed, dates).

    The unattributed bucket is cut tonnage whose operator maps to no saw. It is REAL and
    belongs in the headline total, but it cannot be assigned to a plant. Never default it
    to P1: all-history that is ~26% of cut, which would be a very large lie about Plant 1.
    On current periods it runs 2-5% (Aug 94.6% attributed, Sep 97.7%)."""
    p1=p4=un=0.0; ds=set()
    for iso,per in detail.items():
        dd=d0(iso)
        if ps<=dd<=pe:
            for sid,rec in per.items():
                pl=B.SAW_PLANT_OF.get(sid)      # None for the unattributed bucket
                if   pl=="P1": p1+=rec["tons"]
                elif pl=="P4": p4+=rec["tons"]
                else:          un+=rec["tons"]
            ds.add(dd)
    return p1,p4,un,ds

def build(ps,pe,buckets,labels):
    cur = ps<=today<=pe
    tday = today if cur else pe+datetime.timedelta(days=1)   # past -> fully elapsed
    cs=period_cut(ps,pe)
    # PLANT_GOALS are WEEKLY goals. For a month, scale them to that month's Mon-Fri workdays
    # (same rule as build_production_json: weekly goal x workdays/5, Saturday is bonus), so
    # the monthly board shows a monthly goal instead of the weekly one (bug seen 2026-09-08).
    goals=B.PLANT_GOALS
    if labels.get("scope")=="month":
        _sc=B.workdays_between(ps,pe)/5
        goals={k:{"fit":round(v["fit"]*_sc),"cut":round(v["cut"]*_sc)} for k,v in B.PLANT_GOALS.items()}
    data=B.aggregate(fit,hrs,names,ps,pe,buckets,tday,goals,cs,detail,labels)
    if not cs[3]:   # no saw data this period -> cut as single total, no P1/P4 split
        ct=round(sum(cutday.get(ps+datetime.timedelta(days=i),0) for i in range((pe-ps).days+1)),1)
        data["top"]["cut_wtd"]=ct
        for k in ("p1","p4"):
            data[k]["cut_wtd"]=None; data[k]["cut_pct"]=None; data[k]["cut_proj"]=None
        data["saws"]=[]; data["cut_split"]=False
    else:
        data["cut_split"]=True
    data["paint"]=B.paint_rows({d:(v[0],v[1]) for d,v in paintday.items()}, paint_bp, buckets, ps, pe)
    return data

# ---- weeks (Mon-Sun; Sunday added 2026-09-29) ----
weeks=[]
w=B.monday_of(alldates[0])
end=B.monday_of(today)
while w<=end:
    days=[w+datetime.timedelta(days=i) for i in range(7)]
    if any(w<=r["dt"]<=days[-1] for r in fit) or w==end:
        bk=[(d.strftime("%a"),{d}) for d in days]
        wk=w.isocalendar()[1]; sun=w+datetime.timedelta(days=6)
        rng=f"{w.strftime('%b')} {w.day}–{sun.day}" if w.month==sun.month else f"{w.strftime('%b')} {w.day} – {sun.strftime('%b')} {sun.day}"
        weeks.append({"key":w.isoformat(),"label":f"Week {wk} · {rng}",
            "data":build(w,days[-1],bk,{"scope":"week","title":f"Week {wk} · {rng}","period_label":"week",
                "chart_label":"","tile_big":"day","tile_small":"wk"})})
    w+=datetime.timedelta(days=7)

# ---- months ----
months=[]
m=alldates[0].replace(day=1)
mend=today.replace(day=1)
while m<=mend:
    m_end=(m.replace(day=28)+datetime.timedelta(days=4)).replace(day=1)-datetime.timedelta(days=1)
    wk=B.monday_of(m); bk=[]
    while wk<=m_end:
        dd={wk+datetime.timedelta(days=i) for i in range(7)}; dd={x for x in dd if m<=x<=m_end}   # Mon-Sun
        if dd: md=min(dd); bk.append((f"Wk {md.month}/{md.day}",dd))
        wk+=datetime.timedelta(days=7)
    months.append({"key":m.isoformat(),"label":m.strftime("%B %Y"),
        "data":build(m,m_end,bk,{"scope":"month","title":m.strftime("%B %Y"),"period_label":"month",
            "chart_label":"","tile_big":"wk","tile_small":"MTD"})})
    m=(m.replace(day=28)+datetime.timedelta(days=4)).replace(day=1)

hist={"generated_at":datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
      "weeks":weeks,"months":months,
      "current_week":B.monday_of(today).isoformat(),"current_month":today.replace(day=1).isoformat()}
B.atomic_write(os.path.join(OUT,"history_data.json"), json.dumps(hist))
_tpl=os.path.join(OUT,"explorer_template.html")
if os.path.exists(_tpl):
    B.atomic_write(os.path.join(OUT,"explorer.html"), open(_tpl,encoding="utf-8").read().replace("__HISTORY_JSON__", json.dumps(hist)))
print(f"weeks {len(weeks)} | months {len(months)}")
cw=[x for x in weeks if x["key"]==hist["current_week"]]
if cw:
    d=cw[0]["data"]; print("current week:",cw[0]["label"],"| fit P1",d["p1"]["fit_wtd"],"P4",d["p4"]["fit_wtd"],"| cut top",d["top"]["cut_wtd"],"| split",d.get("cut_split"))
pw=weeks[-4] if len(weeks)>=4 else weeks[0]
d=pw["data"]; print("3 weeks ago:",pw["label"],"| fit P1",d["p1"]["fit_wtd"],"P4",d["p4"]["fit_wtd"],"| cut top",d["top"]["cut_wtd"],"| split",d.get("cut_split"))
old=weeks[10] if len(weeks)>10 else weeks[0]
d=old["data"]; print("old (2025):",old["label"],"| fit P1",d["p1"]["fit_wtd"],"| cut top",d["top"]["cut_wtd"],"| split",d.get("cut_split"),"| saws",len(d["saws"]))
