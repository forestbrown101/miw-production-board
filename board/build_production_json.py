#!/usr/bin/env python3
"""
MIW Production Board - data builder
Writes production_data.json (weekly, Mon-Sat) and production_data_month.json (month-to-date).

DATA SOURCE (auto-detected from first arg):
  - A FOLDER (the "Daily Reports" drop folder) -> reads raw PowerFab/QuickBooks files directly. LIVE MODE.
      Needs job names: set NAMES env var (or 3rd arg) to the backlog workbook path.
  - An .xlsx file -> reads the workbook's combined tabs (QCFinal/CutDaily/QuickBooks). LEGACY MODE.

Usage:  python build_production_json.py <SOURCE> <OUT_DIR> [NAMES_XLSX]
Run every 5 min via Task Scheduler. Both TVs update from one run.
"""
import json, datetime, sys, os, re, io, csv, glob
import openpyxl

SOURCE   = sys.argv[1] if len(sys.argv) > 1 else "source.xlsx"
OUT_DIR  = sys.argv[2] if len(sys.argv) > 2 else "."
NAMES_XLSX = sys.argv[3] if len(sys.argv) > 3 else os.environ.get("NAMES","")
WEEK_GOAL   = 600
CUT_WEEK_GOAL = 650
WORK_DAYS_PER_WEEK = 7   # Mon-Sun columns (Forest 2026-09-29: Sunday cut/QC was falling off the board). Pace stays Mon-Fri.
TOP_JOBS = 6
TODAY_OVERRIDE = os.environ.get("TODAY")
STATION_CSV = os.environ.get("STATION_CSV","")
PAINT_CSV = os.environ.get("PAINT_CSV","")      # paint_by_plant_history.csv (P1/P2 split)
QC_CSV = os.environ.get("QC_CSV","")
# THE CANONICAL FEED. Same two files the WEEK ## SCHEDULES workbook reads, so the
# board and the workbook cannot drift. Falls back to QC_CSV / the saw folders if unset.
FEED_FIT = os.environ.get("FEED_FIT","")
FEED_CUT = os.environ.get("FEED_CUT","")
# Forest 2026-09-08: P3 is COLLECTED in the feed but ignored on the board. It charges
# ~10% of shop hours for ~1% of finaled tons, which is a scan-capture gap, not a
# performance number, and it does not belong on a plant-floor TV next to P1 and P4.
# Dropping it here removes its HOURS too, so MH/ton stays honest.
SHOW_P3 = False

FIT_LOC_TO_PLANT = {"East Bay":"P1","West Bay":"P1","Paint Bay":"P1","Plant 3":"P3","Plant 4":"P4"}
def tile_plant(loc):   # FABRICATION hours only, per Forest 2026-09-10: P1 = the bays, P4 = Fab + Fab night shift
    # Receiving, Detail, Shipping, Paint and QC are deliberately stripped out of the plant
    # tiles so MH/ton is fabrication labour over fabricated tons. The P4 night shift IS a
    # fabrication crew (its output lands in P4 fit tons), so "P4 - Fab (Nightshift)" counts;
    # the old "P4 - Fab"-only filter dropped it and understated P4 hours.
    loc = loc or ""
    if loc in ("P1 - East Bay","P1 - West Bay"): return "P1"
    if loc.startswith("P3"): return "P3" if SHOW_P3 else None
    if loc.startswith("P4 - Fab"): return "P4"          # "P4 - Fab" and "P4 - Fab (Nightshift)"
    return None
def overall_plant(loc):  # STRUCTURAL overall: every department in P1/P3/P4
    loc = loc or ""
    if loc.startswith("P1"): return "P1"
    if loc.startswith("P3"): return "P3" if SHOW_P3 else None
    if loc.startswith("P4"): return "P4"
    return None
PLANTS = ([("P1","Plant 1"),("P3","Plant 3"),("P4","Plant 4")] if SHOW_P3
          else [("P1","Plant 1"),("P4","Plant 4")])
SAW_PLANT = {"Saw 1":"P1","Saw 2":"P1","Saw 3":"P1","Saw 4":"P1","Saw 5 (DAY SHIFT)":"P4","Saw 5 (NIGHT SHIFT)":"P4"}
PLANT_GOALS = {"P1":{"fit":600,"cut":650},"P4":{"fit":100,"cut":300}}
# non-worked / leave codes (in jobcode_1) that are NOT actual man-hours - excluded to match the Weekly HR Report
NONWORK_KEYS = ("lunch","vacation","time off","no call","no show","tardy","leave early","holiday","bereavement","jury","sick","pto")
SAW_ID_OF = {"Saw 1":"S1","Saw 2":"S2","Saw 3":"S3","Saw 4":"S4","Saw 5 (DAY SHIFT)":"S5","Saw 5 (NIGHT SHIFT)":"S5"}
SAW_PLANT_OF = {"S1":"P1","S2":"P1","S3":"P1","S4":"P1","S5":"P4"}
SAW_DEFS = [("S1","Saw 1","P1"),("S2","Saw 2","P1"),("S3","Saw 3","P1"),("S4","Saw 4","P1"),("S5","Saw 5","P4"),
            ("SP","Plate","")]   # Forest 2026-09-29: Afilyaw = plate. No plant yet, so SAW_PLANT_OF has no SP and it rides the unattributed P1/P4 bucket.

def d0(d):
    if isinstance(d, datetime.datetime): return d.date()
    if isinstance(d, datetime.date): return d
    return None
def monday_of(d): return d - datetime.timedelta(days=d.weekday())
def numf(v):
    if v is None: return 0.0
    if isinstance(v,(int,float)): return float(v)
    try: return float(str(v).replace(",","").replace("#","").replace("$","").strip())
    except: return 0.0
def file_date(fn):
    m=re.match(r"(\d{8})", os.path.basename(fn))
    return datetime.date(int(m.group(1)[:4]),int(m.group(1)[4:6]),int(m.group(1)[6:8])) if m else None
def loadx(p):
    try: return openpyxl.load_workbook(io.BytesIO(open(p,"rb").read()), data_only=True)
    except Exception: return None
def find_header(ws):
    for r in range(1, ws.max_row+1):
        if ws.cell(r,1).value=="Job #": return r
    return None
def is_jobnum(a): return isinstance(a,(int,float)) or (isinstance(a,str) and str(a).strip().isdigit())

def _find_cols(ws, hrow):
    cols={}
    for c in range(1, ws.max_column+1):
        v=ws.cell(hrow,c).value
        if isinstance(v,str) and v.strip() in ("Job #","Comp","Comp Wt"): cols[v.strip()]=c
    return cols

def read_saw_detail(root):
    """Per-date, per-saw cut detail from DATED per-saw files. S5 = day+night merged. Undated legacy files skipped."""
    detail={}
    cutdir=os.path.join(root,"Cut Daily")
    for saw in SAW_ID_OF:
        sid=SAW_ID_OF[saw]
        folder=os.path.join(cutdir,saw)
        if not os.path.isdir(folder): continue
        for fn in glob.glob(os.path.join(folder,"*")):
            base=os.path.basename(fn)
            if base.startswith("~") or base.startswith("."): continue
            m=re.search(r"(\d{8})", base)
            if not m: continue
            g=m.group(1); iso=f"{g[:4]}-{g[4:6]}-{g[6:8]}"
            wb=loadx(fn)
            if wb is None: continue
            ws=wb.active; h=find_header(ws)
            if not h: continue
            cols=_find_cols(ws,h); wcol=cols.get("Comp Wt"); ccol=cols.get("Comp"); jcol=cols.get("Job #",1)
            if not wcol: continue
            rec=detail.setdefault(iso,{}).setdefault(sid,{"tons":0.0,"cuts":0.0,"jobs":{}})
            for r in range(h+1, ws.max_row+1):
                a=ws.cell(r,jcol).value
                if not is_jobnum(a): continue
                tw=numf(ws.cell(r,wcol).value)/2000.0
                rec["tons"]+=tw
                rec["cuts"]+=numf(ws.cell(r,ccol).value) if ccol else 1
                j=int(str(a).strip()); rec["jobs"][j]=rec["jobs"].get(j,0)+tw
    return detail

def read_saw_detail_api(csv_path):
    """Per-date, per-saw cut detail straight from the API cut_by_saw history CSV
    (columns Date,Saw,Plant,Job,Pounds,Tons,Pieces). Returns the SAME structure as
    read_saw_detail so the board consumes it unchanged. This replaces the Daily Reports
    saw .xls exports: cut now rides the PowerFab API pull. The Saw column is
    'Saw 1'..'Saw 5' (day+night already merged into 'Saw 5' by pull_cut_by_saw.py)."""
    name2sid={nm:sid for sid,nm,_ in SAW_DEFS}   # 'Saw 1'->S1 ... 'Saw 5'->S5
    detail={}
    if not csv_path or not os.path.exists(csv_path): return detail
    with open(csv_path,newline="",encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            iso=(r.get("Date") or "")[:10]
            if len(iso)!=10: continue
            sid=name2sid.get((r.get("Saw") or "").strip())
            if not sid: continue
            rec=detail.setdefault(iso,{}).setdefault(sid,{"tons":0.0,"cuts":0.0,"jobs":{}})
            t=numf(r.get("Tons")); rec["tons"]+=t
            rec["cuts"]+=numf(r.get("Pieces"))
            m=re.search(r"(\d{3,5})",str(r.get("Job") or ""))
            if m:
                j=int(m.group(1)); rec["jobs"][j]=rec["jobs"].get(j,0.0)+t
    for per in detail.values():
        for rec in per.values(): rec["cuts"]=int(round(rec["cuts"]))
    return detail

UNATTRIBUTED_SID = "SU"

def load_feed_fit(path):
    """Fit rows from feed_fit.csv, the canonical feed.

    Plant is ALREADY RESOLVED in the feed (C-Beams and SHOP fold into P1 and report as
    West Bay, Shop 3 is Plant 3). Do NOT re-derive it from Test Location here: the old
    FIT_LOC_TO_PLANT map has no entry for C-Beams and silently dropped 4,552 tons over
    fifteen months. One mapping, in build_feed.py, and nowhere else."""
    rows=[]
    if not path or not os.path.exists(path): return rows
    try: rdr=csv.DictReader(open(path,encoding="utf-8-sig"))
    except Exception: return rows
    for r in rdr:
        pl=(r.get("Plant") or "").strip()
        if not pl: continue
        if pl=="P3" and not SHOW_P3: continue
        try: dt=datetime.date.fromisoformat((r.get("Report Date") or "")[:10])
        except Exception: continue
        j=None; m=re.search(r"(\d{3,5})", str(r.get("Job #") or ""))
        if m: j=int(m.group(1))
        rows.append({"dt":dt,"p":pl,"t":numf(r.get("Weight"))/2000.0,"j":j,
                     "pc":int(numf(r.get("Total")))})
    return rows

def read_cut_feed(csv_path):
    """Per-date, per-saw cut detail from feed_cut.csv.

    Cut has no plant dimension in PowerFab; attribution is operator -> saw, and unmapped
    operators cannot be placed. That remainder arrives as Saw = '(unattributed)' and is
    carried under SU, never dropped and never spread across the saws. Recent months run
    95-98% attributed; the all-history figure is ~74% because the saw logins are new."""
    name2sid={nm:sid for sid,nm,_ in SAW_DEFS}
    detail={}
    if not csv_path or not os.path.exists(csv_path): return detail
    with open(csv_path,newline="",encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            iso=(r.get("Date") or "")[:10]
            if len(iso)!=10: continue
            saw=(r.get("Saw") or "").strip()
            sid=name2sid.get(saw) or (UNATTRIBUTED_SID if saw else None)
            if not sid: continue
            rec=detail.setdefault(iso,{}).setdefault(sid,{"tons":0.0,"cuts":0.0,"jobs":{}})
            t=numf(r.get("Comp Wt"))/2000.0
            rec["tons"]+=t; rec["cuts"]+=numf(r.get("Comp"))
            m=re.search(r"(\d{3,5})",str(r.get("Job #") or ""))
            if m:
                j=int(m.group(1)); rec["jobs"][j]=rec["jobs"].get(j,0.0)+t
    for per in detail.values():
        for rec in per.values(): rec["cuts"]=int(round(rec["cuts"]))
    return detail

def detail_to_plant(detail):
    out={}
    for iso,perssaw in detail.items():
        agg=out.setdefault(iso,{"P1":0.0,"P4":0.0,"UN":0.0})
        for sid,rec in perssaw.items():
            # never default an unknown saw to P1; that is how a quarter of cut tonnage
            # would silently land in Plant 1.
            agg[SAW_PLANT_OF.get(sid) or "UN"]+=rec["tons"]
    return out

def merge_saw_history(out_dir, detail_now):
    """SAFEGUARD: persist the HIGHEST reading per saw per day. Cut only climbs during a day,
    so an empty/partial export (contention, mid-write, sync) can never drop a saw's number -
    the next good pull tops it off. Returns the max-merged detail used for all cut figures."""
    path=os.path.join(out_dir,"saw_history.json")
    store={}
    if os.path.exists(path):
        try: store=json.load(open(path,encoding="utf-8"))
        except Exception: store={}
    for iso,saws in detail_now.items():
        day=store.setdefault(iso,{})
        for sid,rec in saws.items():
            prev=day.get(sid)
            if prev is None or rec["tons"]>=prev.get("tons",0):   # keep the higher reading
                day[sid]={"tons":round(rec["tons"],4),"cuts":int(round(rec["cuts"])),
                          "jobs":{str(k):round(v,4) for k,v in rec["jobs"].items()}}
    atomic_write(path, json.dumps(store,indent=2))
    out={}
    for iso,saws in store.items():
        out[iso]={sid:{"tons":r["tons"],"cuts":r.get("cuts",0),
                       "jobs":{int(k):v for k,v in r.get("jobs",{}).items()}} for sid,r in saws.items()}
    return out

def merge_cut_history(out_dir, file_hist):
    """Persist per-date cut so it survives even if dated files are later cleaned up. Files win for dates they cover."""
    path=os.path.join(out_dir,"cut_saw_history.json")
    hist={}
    if os.path.exists(path):
        try: hist=json.load(open(path,encoding="utf-8"))
        except Exception: hist={}
    for iso,vals in file_hist.items():
        hist[iso]={"P1":round(vals.get("P1",0.0),3),"P4":round(vals.get("P4",0.0),3)}
    atomic_write(path, json.dumps(hist, indent=2))
    return hist

# ---------- RAW FOLDER READERS (validated against workbook tabs) ----------
def _row_with(ws, label):
    for r in range(1, min(ws.max_row,20)+1):
        v=ws.cell(r,1).value
        if isinstance(v,str) and label.lower() in v.lower(): return r
    return None
def _row_value(ws, r):
    for c in range(2, ws.max_column+1):
        v=ws.cell(r,c).value
        if v not in (None,""): return v
    return None
def _parse_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    if isinstance(v, str):
        for fmt in ("%m/%d/%Y","%m/%d/%y","%Y-%m-%d"):
            try: return datetime.datetime.strptime(v.strip(), fmt).date()
            except: pass
    return None
def cut_workdate(ws, fd):
    r=_row_with(ws, "Date Station Completed")
    if r:
        d=_parse_date(_row_value(ws, r))
        if d: return d
    return fd-datetime.timedelta(days=1)
def qc_workdate(ws, fd):
    r=_row_with(ws, "Test Date")
    if r:
        v=_row_value(ws, r)
        d=_parse_date(v)
        if d: return d                      # explicit date set in the filter
        if isinstance(v,str) and "current date" in v.lower():
            m=re.search(r"-\s*(\d+)", v)
            off=int(m.group(1)) if m else 0  # "Current Date" -> 0 (today); "- 1 day" -> 1
            return fd-datetime.timedelta(days=off)
    return fd-datetime.timedelta(days=1)

def read_raw(root):
    fit_rows=[]; cut_rows=[]; hr_rows=[]
    # cut now comes from the per-saw folders (read_saw_cut); combined report intentionally ignored
    for fn in glob.glob(os.path.join(root,"QC Final Report","*")):
        fd=file_date(fn)
        if not fd: continue
        _wb=loadx(fn)
        if _wb is None: continue
        ws=_wb.active; h=find_header(ws)
        wd=qc_workdate(ws, fd)
        if not h: continue
        totcol=next((c for c in range(1,ws.max_column+1) if ws.cell(h,c).value=="Total"), 7)  # piece count
        for r in range(h+1, ws.max_row+1):
            a=ws.cell(r,1).value
            if not is_jobnum(a): continue
            p=FIT_LOC_TO_PLANT.get(ws.cell(r,4).value)
            if not p: continue
            fit_rows.append({"dt":wd,"p":p,"t":numf(ws.cell(r,8).value)/2000.0,"j":int(str(a).strip()),
                             "pc":numf(ws.cell(r,totcol).value)})
    for fn in glob.glob(os.path.join(root,"QuickBooks Time","*.csv")):
        try:
            raw=open(fn, encoding="utf-8", errors="replace").read().replace("\x00","")  # tolerate NUL/partial writes
        except Exception:
            continue
        try:
            reader=csv.DictReader(io.StringIO(raw))
            for row in reader:
                try: d=datetime.date.fromisoformat(row["local_date"])
                except: continue
                grp=row.get("group")
                j=None
                m=re.search(r"(\d{3,5})", str(row.get("jobcode_2") or ""))
                if m: j=int(m.group(1))
                if j is None:
                    code=((row.get("jobcode_1") or "")+" "+(row.get("jobcode_2") or "")).lower()
                    if any(k in code for k in NONWORK_KEYS): continue   # drop leave/absence codes
                hr_rows.append({"dt":d,"p":tile_plant(grp),"po":overall_plant(grp),"h":numf(row.get("hours")),"j":j})
        except Exception:
            continue  # one corrupt CSV must never crash the whole build
    return fit_rows, cut_rows, hr_rows

# ---------- LEGACY: workbook combined tabs ----------
def load_rows(ws, cols):
    hdr={ws.cell(1,c).value:c for c in range(1,ws.max_column+1)}
    idx={k:hdr[k] for k in cols}
    return [{k:ws.cell(r,idx[k]).value for k in cols} for r in range(2,ws.max_row+1)]
def read_workbook(path):
    wb=openpyxl.load_workbook(path, data_only=True)
    fit_rows=[]
    for row in load_rows(wb["QCFinal"],["Report Date","Test Location","Weight","Job #"]):
        dt=d0(row["Report Date"]); p=FIT_LOC_TO_PLANT.get(row["Test Location"])
        if dt and p: fit_rows.append({"dt":dt,"p":p,"t":(row["Weight"] or 0)/2000.0,"j":row["Job #"]})
    cut_rows=[]
    for row in load_rows(wb["CutDaily"],["Date","Comp Wt"]):
        dt=d0(row["Date"])
        if dt: cut_rows.append({"dt":dt,"t":(row["Comp Wt"] or 0)/2000.0})
    hr_rows=[]
    for row in load_rows(wb["QuickBooks"],["Location","Date","Hours","Job #"]):
        dt=d0(row["Date"]); loc=row["Location"]
        j=None
        if row["Job #"]:
            mm=re.search(r"(\d{3,5})", str(row["Job #"]))
            if mm: j=int(mm.group(1))
        if dt: hr_rows.append({"dt":dt,"p":tile_plant(loc),"po":overall_plant(loc),"h":(row["Hours"] or 0),"j":j})
    return fit_rows, cut_rows, hr_rows

def load_names(path):
    names={}
    if not path or not os.path.exists(path): return names
    wb=openpyxl.load_workbook(path, data_only=True)
    if "2026 WEEKLY HR REPORT" in wb.sheetnames:
        hr=wb["2026 WEEKLY HR REPORT"]
        for r in range(9,200):
            a=hr.cell(r,1).value; j=hr.cell(r,10).value
            if isinstance(a,(int,float)) and j: names[int(a)]=str(j)
    return names

def atomic_write(path, text):
    tmp=path+".tmp"
    with open(tmp,"w",encoding="utf-8") as f: f.write(text)
    os.replace(tmp, path)

SHIFT_START_H=6.0; SHIFT_END_H=16.5   # day-shift window for intraday pace ramp
def _day_fraction(dtnow):
    h=dtnow.hour+dtnow.minute/60.0
    if h<=SHIFT_START_H: return 0.0
    if h>=SHIFT_END_H: return 1.0
    return (h-SHIFT_START_H)/(SHIFT_END_H-SHIFT_START_H)

def workdays_between(start, end):
    n=0; d=start
    while d<=end:
        if d.weekday()<5: n+=1   # Mon-Fri only; Saturday is optional/bonus
        d+=datetime.timedelta(days=1)
    return n

def aggregate(fit_rows, hr_rows, names, period_start, period_end, buckets, today,
              pgoals, cut_split, detail, labels):
    pset=set()
    for _,ds in buckets: pset|=ds
    bidx={}
    for i,(_,ds) in enumerate(buckets):
        for d in ds: bidx[d]=i
    nb=len(buckets)
    fit_b={p:[0.0]*nb for p,_ in PLANTS}
    fit_pc_b={p:[0]*nb for p,_ in PLANTS}
    job_tons={}; job_plant_tons={}
    for row in fit_rows:
        dt=row["dt"]; p=row["p"]
        if dt in bidx and p: fit_b[p][bidx[dt]]+=row["t"]; fit_pc_b[p][bidx[dt]]+=row.get("pc",0)
        if dt in pset and p:
            j=row["j"]
            if j is None: continue
            j=int(j)
            job_tons[j]=job_tons.get(j,0)+row["t"]
            job_plant_tons.setdefault(j,{}).setdefault(p,0); job_plant_tons[j][p]+=row["t"]
    hours={p:0.0 for p,_ in PLANTS}; overall_hours=0.0; job_hours={}
    hours_b={p:[0.0]*nb for p,_ in PLANTS}          # hours per bucket, so a tile's MH/ton can match its tons period
    for row in hr_rows:
        dt=row["dt"]
        if dt in pset:
            if row.get("p"): hours[row["p"]]+=row["h"]; hours_b[row["p"]][bidx[dt]]+=row["h"]
            if row.get("po"): overall_hours+=row["h"]
            if row.get("j") is not None: job_hours[row["j"]]=job_hours.get(row["j"],0)+row["h"]
    bucket_totals=[sum(fit_b[p][i] for p,_ in PLANTS) for i in range(nb)]
    last_idx=max([i for i,v in enumerate(bucket_totals) if v>0], default=0)
    plants_out=[]
    for key,name in PLANTS:
        wtd=sum(fit_b[key])
        ft=fit_b[key][last_idx]; ht=hours_b[key][last_idx]
        plants_out.append({"key":key,"name":name,"fit_today":round(ft,1),
            "fit_wtd":round(wtd,1),"hours_wtd":round(hours[key],1),
            "mhton":round(hours[key]/wtd,1) if wtd>0 else None,            # whole-period rate (wk on weekly, month on monthly)
            # Same-period rate for the tile: the tile's big number is the LAST bucket's tons
            # (today on the weekly board, this week on the monthly board), so its MH/ton must be
            # that bucket's hours over that bucket's tons, not the period rate (Forest 2026-09-10).
            "hours_today":round(ht,1),
            "mhton_today":round(ht/ft,1) if ft>0 else None,
            "daily_fit":[round(x,1) for x in fit_b[key]],
            "daily_fit_pc":[int(round(x)) for x in fit_pc_b[key]],"fit_pc_wtd":int(round(sum(fit_pc_b[key])))})
    fitw={p["key"]:p["fit_wtd"] for p in plants_out}
    fit_total=sum(p["fit_wtd"] for p in plants_out)
    hours_total=overall_hours
    # Accept both shapes: (p1,p4,dates) from older callers, (p1,p4,unattributed,dates) now.
    if len(cut_split)==4: cut_p1,cut_p4,cut_un,cut_dates=cut_split
    else:                 cut_p1,cut_p4,cut_dates=cut_split; cut_un=0.0
    active=[]
    for j,t in sorted(job_tons.items(), key=lambda x:-x[1])[:TOP_JOBS]:
        pp=job_plant_tons[j]
        active.append({"job":int(j),"tons":round(t,1),
                       "p1":round(pp.get("P1",0.0),1),"p3":round(pp.get("P3",0.0),1),"p4":round(pp.get("P4",0.0),1),
                       "hours":round(job_hours.get(j,0),1),
                       "name":names.get(int(j), f"Job {int(j)}")})
    data_dates=set()
    for row in fit_rows:
        if row["dt"] in pset and row["p"]: data_dates.add(row["dt"])
    for row in hr_rows:
        if row["dt"] in pset: data_dates.add(row["dt"])
    data_dates|=set(cut_dates)
    last_day=max([d for d in data_dates if d<=today], default=today)
    total_wd=workdays_between(period_start, period_end)   # Mon-Fri (Sat optional)
    now=datetime.datetime.now()
    if today<period_start:
        elapsed_frac=0.0; day_disp=0
    elif today>period_end:
        elapsed_frac=float(total_wd); day_disp=total_wd
    else:
        completed_before=workdays_between(period_start, today-datetime.timedelta(days=1))
        if today.weekday()<5:                             # Mon-Fri: ramp through the shift
            elapsed_frac=min(total_wd, completed_before+_day_fraction(now)); day_disp=min(total_wd, completed_before+1)
        else:                                             # Saturday: M-F already fully counted
            elapsed_frac=float(total_wd); day_disp=total_wd
    proj_elapsed=max(elapsed_frac,0.4)                    # floor so early-day projection doesn't blow up
    cut_elapsed=max(1,len(cut_dates))
    def blk(key,cutw):
        g=pgoals[key]; fw=fitw.get(key,0.0)
        exp_fit=g["fit"]*elapsed_frac/total_wd if total_wd else 0
        return {"fit_wtd":round(fw,1),"fit_goal":g["fit"],
            "fit_pct":round(fw/g["fit"]*100) if g["fit"] else 0,
            "fit_proj":round(fw/proj_elapsed*total_wd,0) if total_wd else 0,
            "fit_expected":round(exp_fit,1),"fit_on_pace":fw>=exp_fit,
            "cut_wtd":round(cutw,1),"cut_goal":g["cut"],
            "cut_pct":round(cutw/g["cut"]*100) if g["cut"] else 0,
            "cut_proj":round(cutw/cut_elapsed*total_wd,0)}
    # Forest 2026-09-29: unattributed cut gets its own "Unassigned" row so the saw table's day
    # columns add up to the headline and the workbook (Mon 9/21 read 68.5 vs 107.8 without it).
    # The row only appears when it has tons.
    # Forest 2026-09-29: no Other row. Unmapped logins are mostly the angle line, which is not
    # tracked on the saw table. Their tons stay in the headline Tons Cut (so it still matches the
    # workbook) but get no row. Only Saw 1-5 and Plate are listed.
    SAW_ROWS=SAW_DEFS
    saw_tons={s[0]:0.0 for s in SAW_ROWS}; saw_cuts={s[0]:0.0 for s in SAW_ROWS}; saw_today={s[0]:0.0 for s in SAW_ROWS}
    saw_daily={s[0]:[0.0]*nb for s in SAW_ROWS}; saw_daily_pc={s[0]:[0]*nb for s in SAW_ROWS}
    # Forest 2026-09-29: job breakdown behind every saw cell, for the hover tooltip.
    saw_daily_jobs={s[0]:[{} for _ in range(nb)] for s in SAW_ROWS}
    cutjob_tons={}; cutjob_plant={}
    for iso,perssaw in detail.items():
        d=datetime.date.fromisoformat(iso)
        if not (period_start<=d<=period_end and d<=today): continue
        bi=bidx.get(d)
        for sid,rec in perssaw.items():
            if sid in saw_tons:
                saw_tons[sid]+=rec["tons"]; saw_cuts[sid]+=rec["cuts"]
                if bi is not None:
                    saw_daily[sid][bi]+=rec["tons"]; saw_daily_pc[sid][bi]+=rec["cuts"]
                    bj=saw_daily_jobs[sid][bi]
                    for job,tt in rec["jobs"].items(): bj[job]=bj.get(job,0.0)+tt
                if d==today: saw_today[sid]+=rec["tons"]
            pl=SAW_PLANT_OF.get(sid)          # None for the unattributed bucket
            for job,tt in rec["jobs"].items():
                cutjob_tons[job]=cutjob_tons.get(job,0)+tt   # job TOTAL is always complete
                if pl:
                    cutjob_plant.setdefault(job,{}).setdefault(pl,0); cutjob_plant[job][pl]+=tt
    saws_out=[{"id":sid,"name":nm,"plant":pl,"tons_wtd":round(saw_tons[sid],1),
        "tons_today":round(saw_today[sid],1),"pieces_wtd":int(round(saw_cuts[sid])),"util_pct":None,
        "daily":[round(x,1) for x in saw_daily[sid]],"daily_pc":[int(round(x)) for x in saw_daily_pc[sid]],
        # [[job, tons], ...] per bucket, largest first; slivers under 0.05 t dropped. Kept at 2 dp
        # so a tooltip that adds buckets (Total column, footer) matches the cell it points at.
        "daily_jobs":[[[int(j),round(t,2)] for j,t in sorted(bj.items(),key=lambda x:-x[1]) if t>=0.05]
                      for bj in saw_daily_jobs[sid]]} for sid,nm,pl in SAW_ROWS
        if sid!=UNATTRIBUTED_SID or saw_tons[sid]>0.05]
    # .get(j,{}) not [j]: a job whose cut is entirely unattributed has a total but no
    # plant entry at all, and indexing it raised KeyError.
    cut_jobs=[{"job":int(j),"tons":round(tt,1),
        "p1":round(cutjob_plant.get(j,{}).get("P1",0.0),1),"p4":round(cutjob_plant.get(j,{}).get("P4",0.0),1),
        "name":names.get(int(j), f"Job {int(j)}")} for j,tt in sorted(cutjob_tons.items(), key=lambda x:-x[1])[:TOP_JOBS]]
    p1=blk("P1",cut_p1); p4=blk("P4",cut_p4)
    # Headline cut is the COMPLETE figure, attributed plus unattributed, so it matches
    # the workbook. The per-plant cut tiles below stay attributed-only, so on any day with
    # unattributed tonnage top.cut_wtd is deliberately larger than p1.cut + p4.cut.
    top={"fit_wtd":round(fitw.get("P1",0)+fitw.get("P4",0),1),
         "cut_wtd":round(cut_p1+cut_p4+cut_un,1),
         "cut_unattributed":round(cut_un,1)}
    coverage=round(top["cut_wtd"]/top["fit_wtd"]*100) if top["fit_wtd"]>0 else None
    return {"generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "scope":labels["scope"],"week_label":labels["title"],"period_label":labels["period_label"],
        "chart_label":labels["chart_label"],"days":[{"label":b[0],"date":f"{min(b[1]).month}/{min(b[1]).day}"} for b in buckets],
        "today_index":last_idx,"today_label":f"{last_day.month}/{last_day.day}",
        "tile_big":labels["tile_big"],"tile_small":labels["tile_small"],"goal_tons":pgoals["P1"]["fit"],
        "totals":{"hours_wtd":round(hours_total,1),
            "mhton":round(hours_total/fit_total,1) if fit_total>0 else None,
            "elapsed_days":day_disp,"total_days":total_wd},
        "top":top,"p1":p1,"p4":p4,
        "plants":plants_out,"active_jobs":active,
        "saws":saws_out,"cut_jobs":cut_jobs,"coverage":coverage,
        "daily_total_fit":[round(x,1) for x in bucket_totals],
        # names for every job in the saw tooltips (backlog workbook), so the box can say
        # "2319  ALIGNED CHM04" instead of a bare number. Missing names fall back to the number.
        "job_names":{str(j):names[j] for j in sorted({p[0] for s in saws_out for b in s["daily_jobs"] for p in b})
                     if j in names}}

PAINT_PLANTS=[("P1","Plant 1"),("P2","Plant 2")]
P2_PAINT_START=datetime.date(2026,9,15)     # Plant 2 paint shop first day; row shows from here on

def read_paint_by_plant(path):
    """paint_by_plant_history.csv -> {date: {plant: [tons, pieces]}}. Empty dict if missing."""
    out={}
    if not path or not os.path.exists(path): return out
    try: rdr=csv.DictReader(open(path,encoding="utf-8"))
    except Exception: return out
    for r in rdr:
        try: d=datetime.date.fromisoformat((r.get("Date") or "")[:10])
        except Exception: continue
        p=r.get("Plant") or "P1"
        rec=out.setdefault(d,{}).setdefault(p,[0.0,0])
        rec[0]+=float(r.get("Tons") or 0); rec[1]+=int(float(r.get("Pieces") or 0))
    return out

def paint_rows(paintday, byplant, buckets, ps, pe):
    """Board rows for the Paint table. paintday = {date:(tons,pieces)} total from the station
    history; byplant = read_paint_by_plant(). A date present in byplant is split by plant;
    a date missing from it (before the split feed existed) counts as Plant 1. Plant 2 shows
    once the period reaches its start date, so a zero day is visible rather than hidden."""
    bidx={}
    for i,(_,ds) in enumerate(buckets):
        for x in ds: bidx[x]=i
    nb=len(buckets)
    acc={p:([0.0]*nb,[0]*nb) for p,_ in PAINT_PLANTS}
    for d,i in bidx.items():
        if d in byplant:
            for p,(t,q) in byplant[d].items():
                if p not in acc: acc[p]=([0.0]*nb,[0]*nb)
                acc[p][0][i]+=t; acc[p][1][i]+=q
        else:
            t,q=paintday.get(d,(0.0,0))
            acc["P1"][0][i]+=t; acc["P1"][1][i]+=q
    rows=[]
    for p,name in PAINT_PLANTS:
        td,pc=acc[p]
        show = sum(td)>0 or (p=="P1") or (p=="P2" and pe>=P2_PAINT_START)
        if not show: continue
        rows.append({"plant":p,"name":name,"daily":[round(x,1) for x in td],
                     "daily_pc":[int(x) for x in pc],"tons_wtd":round(sum(td),1),"pc_wtd":int(sum(pc))})
    return rows if any(r["tons_wtd"]>0 for r in rows) else []

def load_paint(path, buckets, byplant=None, ps=None, pe=None):
    """Paint tons + pieces by bucket and plant, from the API Paint station."""
    paintday={}
    if path and os.path.exists(path):
        try:
            for r in csv.DictReader(open(path,encoding="utf-8")):
                if r.get("Station")!="Paint": continue
                try: d=datetime.date.fromisoformat((r.get("StartDate") or "")[:10])
                except Exception: continue
                t,q=paintday.get(d,(0.0,0))
                paintday[d]=(t+float(r.get("Tons") or 0), q+int(float(r.get("Quantity") or 0)))
        except Exception: pass
    alld=[x for _,ds in buckets for x in ds]
    if ps is None: ps=min(alld)
    if pe is None: pe=max(alld)
    return paint_rows(paintday, byplant or {}, buckets, ps, pe)

def load_fit_csv(path):
    """Fit rows from the API qc_history/qc_by_plant CSV (VT-Final, passed)."""
    rows=[]
    try: rdr=csv.DictReader(open(path,encoding="utf-8"))
    except Exception: return rows
    for r in rdr:
        if r.get("Test")!="VT - Final Visual Weld Inspection": continue
        if int(float(r.get("Failed") or 0))!=0: continue
        p=FIT_LOC_TO_PLANT.get(r.get("Plant"))
        if not p: continue
        try: dt=datetime.date.fromisoformat((r.get("Date") or "")[:10])
        except Exception: continue
        j=None; m=re.search(r"(\d{3,5})", str(r.get("Job") or ""))
        if m: j=int(m.group(1))
        rows.append({"dt":dt,"p":p,"t":float(r.get("Tons") or 0),"j":j,"pc":int(float(r.get("Quantity") or 0))})
    return rows

def main():
    today=datetime.date.fromisoformat(TODAY_OVERRIDE) if TODAY_OVERRIDE else datetime.date.today()
    using_feed=False
    if os.path.isdir(SOURCE):
        fit_rows,cut_rows,hr_rows=read_raw(SOURCE); names=load_names(NAMES_XLSX)
        detail_now=read_saw_detail(SOURCE)
        # Preference order: canonical feed > raw API history > the .xls export folders.
        if FEED_FIT and os.path.exists(FEED_FIT):
            fit_rows=load_feed_fit(FEED_FIT); using_feed=True
        elif QC_CSV and os.path.exists(QC_CSV):
            fit_rows=load_fit_csv(QC_CSV)
        if FEED_CUT and os.path.exists(FEED_CUT):
            detail_now=read_cut_feed(FEED_CUT); using_feed=True
    else:
        fit_rows,cut_rows,hr_rows=read_workbook(SOURCE); names=load_names(SOURCE)
        detail_now={}
    if using_feed:
        # NO max-merge on the feed. That safeguard exists because the .xls saw exports
        # intermittently came back empty and a zero must never erase a real number. The
        # API is a query, not an export: it cannot come back partially empty, and
        # max-merging it would pin a day to a stale high forever.
        detail=detail_now
    else:
        detail=merge_saw_history(OUT_DIR, detail_now)   # SAFEGUARD: daily max per saw
    cut_hist=detail_to_plant(detail)
    def period_cut(ps,pe):
        p1=p4=un=0.0; ds=set()
        for k,v in cut_hist.items():
            try: d=datetime.date.fromisoformat(k)
            except Exception: continue
            if ps<=d<=pe and d<=today:
                p1+=v.get("P1",0); p4+=v.get("P4",0); un+=v.get("UN",0); ds.add(d)
        return p1,p4,un,ds

    ws_start=monday_of(today)
    wdays=[ws_start+datetime.timedelta(days=i) for i in range(WORK_DAYS_PER_WEEK)]
    wbuckets=[(d.strftime("%a"),{d}) for d in wdays]
    wk_num=ws_start.isocalendar()[1]; wk_sun=ws_start+datetime.timedelta(days=6)
    if ws_start.month==wk_sun.month: wk_rng=f"{ws_start.strftime('%b')} {ws_start.day}–{wk_sun.day}"
    else: wk_rng=f"{ws_start.strftime('%b')} {ws_start.day} – {wk_sun.strftime('%b')} {wk_sun.day}"
    week=aggregate(fit_rows,hr_rows,names,wdays[0],wdays[-1],wbuckets,today,PLANT_GOALS,period_cut(wdays[0],wdays[-1]),detail,
        {"scope":"week","title":f"Week {wk_num} · {wk_rng}","period_label":"week to date",
         "chart_label":"Fit tons by day · Mon–Sun","tile_big":"today","tile_small":"wtd"})
    paint_bp=read_paint_by_plant(PAINT_CSV)
    week["paint"]=load_paint(STATION_CSV, wbuckets, paint_bp, wdays[0], wdays[-1])
    atomic_write(os.path.join(OUT_DIR,"production_data.json"), json.dumps(week, indent=2))

    m_start=today.replace(day=1)
    m_end=(m_start.replace(day=28)+datetime.timedelta(days=4)).replace(day=1)-datetime.timedelta(days=1)
    wk=monday_of(m_start); mbuckets=[]
    while wk<=m_end:
        days={wk+datetime.timedelta(days=i) for i in range(7)}; days={d for d in days if m_start<=d<=m_end}   # Mon-Sun
        if days:
            md=min(days); mbuckets.append((f"Wk {md.month}/{md.day}", days))
        wk+=datetime.timedelta(days=7)
    wd_month=workdays_between(m_start,m_end)
    _sc=wd_month/5
    mg={k:{"fit":round(v["fit"]*_sc),"cut":round(v["cut"]*_sc)} for k,v in PLANT_GOALS.items()}
    month=aggregate(fit_rows,hr_rows,names,m_start,m_end,mbuckets,today,mg,period_cut(m_start,m_end),detail,
        {"scope":"month","title":today.strftime("%B %Y"),"period_label":"month to date",
         "chart_label":"Fit tons by week","tile_big":"this wk","tile_small":"MTD"})
    month["paint"]=load_paint(STATION_CSV, mbuckets, paint_bp, m_start, m_end)
    atomic_write(os.path.join(OUT_DIR,"production_data_month.json"), json.dumps(month, indent=2))

    # also rebuild self-contained boards so opening the .html file shows current data
    tpl_path=os.path.join(OUT_DIR,"board_template.html")
    if os.path.exists(tpl_path):
        tpl=open(tpl_path,encoding="utf-8").read()
        atomic_write(os.path.join(OUT_DIR,"index.html"),
            tpl.replace("__DATA_FILE__","production_data.json").replace("__FALLBACK_JSON__",json.dumps(week)))
        atomic_write(os.path.join(OUT_DIR,"index_month.html"),
            tpl.replace("__DATA_FILE__","production_data_month.json").replace("__FALLBACK_JSON__",json.dumps(month)))

    print("MODE:", "RAW FOLDER" if os.path.isdir(SOURCE) else "WORKBOOK")
    print("WEEK :",week["week_label"],"as of",week["today_label"],
          "| P1 fit",week["p1"]["fit_wtd"],"/",week["p1"]["fit_goal"],"cut",week["p1"]["cut_wtd"],"/",week["p1"]["cut_goal"],
          "| P4 fit",week["p4"]["fit_wtd"],"/",week["p4"]["fit_goal"],"cut",week["p4"]["cut_wtd"],"/",week["p4"]["cut_goal"],
          "| TOP fit",week["top"]["fit_wtd"],"cut",week["top"]["cut_wtd"])
    print("MONTH:",month["week_label"],"TOP fit",month["top"]["fit_wtd"],"cut",month["top"]["cut_wtd"])

if __name__=="__main__":
    main()
