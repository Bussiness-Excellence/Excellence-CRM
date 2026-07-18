#!/usr/bin/env python3
"""
pulpoplus_auto_upload.py
─────────────────────────
Watches two folders on your machine and automatically uploads data to
Supabase whenever Excel files are added or replaced:

  E:\periods\last_month\   → period label "Last Month"
  E:\periods\recent\       → period label "Recent"

SETUP:
    pip install pandas openpyxl requests watchdog

SET ENV VAR:
    $env:SUPABASE_SERVICE_ROLE_KEY = "sb_secret_..."

USAGE:
    python pulpoplus_auto_upload.py

    # Or specify custom folder paths:
    python pulpoplus_auto_upload.py --last-month "E:\\periods\\last_month" --recent "E:\\periods\\recent"

    # One-time upload without watching:
    python pulpoplus_auto_upload.py --once
"""

import os, sys, time, argparse, math, requests
from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime

import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://xxbfwvlqixnmonxytdxq.supabase.co")
SERVICE_KEY  = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
BATCH_SIZE   = 500

ACC_TYPE_LABELS_LOWER = {
    "clinic","hospital","am center","poly clinics","office work",
    "pharmacy","distributors","activities","events","activity","event",
}
COACHING_TOLERANCE = 90  # minutes


def headers():
    return {
        "apikey": SERVICE_KEY,
        "Authorization": f"Bearer {SERVICE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }

def sb_delete(table, batch):
    requests.delete(f"{SUPABASE_URL}/rest/v1/{table}", headers=headers(),
                    params={"upload_batch": f"eq.{batch}"})

def sb_insert(table, rows):
    if not rows: return 0
    inserted = 0
    for i in range(0, len(rows), BATCH_SIZE):
        r = requests.post(f"{SUPABASE_URL}/rest/v1/{table}",
                          headers=headers(), json=rows[i:i+BATCH_SIZE])
        if r.ok:
            inserted += min(BATCH_SIZE, len(rows)-i)
        else:
            print(f"    ⚠️  Insert {table} chunk {i}: {r.status_code} {r.text[:100]}")
    return inserted

def safe(v):
    if v is None: return None
    if isinstance(v, float) and math.isnan(v): return None
    s = str(v).strip()
    return None if s.lower() in ("nan","none","") else s

def fmt_hm(mins):
    if not mins: return "0:00"
    return f"{int(mins//60)}:{int(mins%60):02d}"

def mins_between(t1, t2):
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            a = datetime.strptime(t1.strip(), fmt)
            b = datetime.strptime(t2.strip(), fmt)
            return abs((b-a).total_seconds()/60)
        except: continue
    return 999


# ── Load all xlsx files from a folder ────────────────────────────────────────

def load_folder(folder):
    records = []
    files = list(Path(folder).glob("*.xlsx"))
    if not files:
        print(f"  ⚠️  No .xlsx files found in {folder}")
        return records
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name="Raw Data", engine="openpyxl")
            for _, row in df.iterrows():
                user = safe(row.get("user"))
                if not user: continue
                records.append({
                    "team":               safe(row.get("team")) or "",
                    "user":               user,
                    "employee_code":      safe(row.get("employee_code")) or "",
                    "territory":          safe(row.get("territory")) or "",
                    "date":               safe(row.get("date")) or "",
                    "time":               safe(row.get("time")) or "",
                    "acc_type_category":  safe(row.get("acc_type_category")) or "",
                    "shift":              safe(row.get("shift")) or "",
                    "visit_type_category":safe(row.get("visit_type_category")) or "",
                    "acc_id":             safe(row.get("acc_id")) or "",
                    "acc_name":           safe(row.get("acc_name")) or "",
                    "doctor_key":         safe(row.get("doctor_key")) or "",
                    "doctor_name":        safe(row.get("doctor_name")) or "",
                    "specialty":          safe(row.get("specialty")) or "",
                    "classification":     safe(row.get("classification")) or "",
                    "products":           safe(row.get("products")) or "",
                    "members":            [user],
                })
            print(f"  ✓ {len(df)} rows from {f.name}")
        except Exception as e:
            print(f"  ⚠️  Could not read {f.name}: {e}")
    return records


# ── KPI computation (same as uploader script) ─────────────────────────────────

def backfill_territory(records):
    date_terr = {}
    user_counts = defaultdict(Counter)
    for r in records:
        t = r["territory"]
        if t and t.lower() not in ACC_TYPE_LABELS_LOWER:
            user_counts[r["user"]][t] += 1
            k = (r["user"], r["date"])
            if k not in date_terr: date_terr[k] = t
    for r in records:
        if r["territory"] and r["territory"].lower() not in ACC_TYPE_LABELS_LOWER: continue
        r["territory"] = date_terr.get((r["user"],r["date"])) or \
            (user_counts[r["user"]].most_common(1)[0][0] if user_counts[r["user"]] else "")

def identify_managers(records):
    terrs = defaultdict(set)
    for r in records:
        t = r["territory"]
        if t and t.lower() not in ACC_TYPE_LABELS_LOWER:
            terrs[r["user"]].add(t)
    return {u for u,ts in terrs.items() if len(ts)>1}

def compute_all(records, batch, period):
    backfill_territory(records)
    managers = identify_managers(records)

    # Also fetch manager roles from hierarchy
    try:
        r = requests.get(f"{SUPABASE_URL}/rest/v1/hierarchy",
                         headers=headers(),
                         params={"select":"employee_name,role","limit":10000})
        if r.ok:
            for row in r.json():
                if row.get("role","") in ("Supervisor","Area Manager","BLM"):
                    managers.add((row.get("employee_name") or "").strip())
    except: pass

    # Coaching days
    matched = defaultdict(set)
    by_key = defaultdict(list)
    for r in records:
        if r["shift"] not in ("AM","PM") or not r["date"] or not r["acc_id"]: continue
        by_key[(r["date"],r["shift"],r["acc_id"])].append((r["user"],r["time"],r["employee_code"]))
    for (date,shift,acc_id), visits in by_key.items():
        mgrs = [(u,t,c) for u,t,c in visits if u in managers]
        reps = [(u,t,c) for u,t,c in visits if u not in managers]
        for mu,mt,mc in mgrs:
            for ru,rt,rc in reps:
                if mt and rt and mins_between(mt,rt) > COACHING_TOLERANCE: continue
                matched[(mu,mc,ru,rc,date)].add(shift)

    coaching_by_mgr = defaultdict(int)
    coaching_rows = []
    for (mgr,mc,rep,rc,date),shifts in matched.items():
        if {"AM","PM"}.issubset(shifts):
            coaching_by_mgr[mgr]+=1
            coaching_rows.append({
                "manager_name":mgr,"manager_code":mc or None,
                "rep_name":rep,"rep_code":rc or None,
                "coaching_date":date,"upload_batch":batch,
            })

    # Summaries
    by_user = defaultdict(list)
    for r in records: by_user[r["user"]].append(r)

    code_by_user = {r["user"]:r["employee_code"] for r in records if r["employee_code"]}
    team_by_user = {r["user"]:r["team"] for r in records}

    summary_rows = []
    for user, urecs in sorted(by_user.items()):
        code = urecs[0]["employee_code"] or code_by_user.get(user,"")
        team = urecs[0]["team"]
        am_recs = [r for r in urecs if r["shift"]=="AM"]
        pm_recs = [r for r in urecs if r["shift"]=="PM"]
        ow_recs = [r for r in urecs if r["acc_type_category"]=="Office Work"]
        ph_recs = [r for r in urecs if r["acc_type_category"]=="Pharmacy"]

        am_dates = set(r["date"] for r in am_recs if r["date"])
        pm_dates = set(r["date"] for r in pm_recs if r["date"])
        ow_dates = set(r["date"] for r in ow_recs if r["date"])
        working_days = len(set(r["date"] for r in urecs if r["date"]))
        complete_field_days = round((len(am_dates|ow_dates)+len(pm_dates))/2,1)

        doc_by_cat = defaultdict(set)
        for r in am_recs+pm_recs:
            if r["doctor_key"] and r["acc_type_category"] in ("Clinic","Poly Clinics","Hospital","AM Center"):
                doc_by_cat[r["acc_type_category"]].add(r["doctor_key"])

        am_calls = len(am_recs); pm_calls = len(pm_recs)
        am_days  = len(am_dates); pm_days  = len(pm_dates)

        def avg_dur(recs):
            bd = defaultdict(list)
            for r in recs:
                if r["date"] and r["time"]: bd[r["date"]].append(r["time"])
            durs = []
            for ts in bd.values():
                ts = sorted(ts)
                if len(ts)>=2: durs.append(mins_between(ts[0],ts[-1]))
            return sum(durs)/len(durs) if durs else 0.0

        am_start_mins = []
        for d in am_dates:
            times = sorted(r["time"] for r in am_recs if r["date"]==d and r["time"])
            if times:
                for fmt in ("%H:%M:%S","%H:%M"):
                    try:
                        dt = datetime.strptime(times[0].strip(),fmt)
                        am_start_mins.append(dt.hour*60+dt.minute); break
                    except: continue
        avg_start = ""
        if am_start_mins:
            avg = sum(am_start_mins)/len(am_start_mins)
            avg_start = f"{int(avg//60):02d}:{int(avg%60):02d}"

        prod_counts = defaultdict(int)
        for r in urecs:
            if r["shift"] in ("AM","PM") and r["products"]:
                for p in r["products"].split(","):
                    p=p.strip()
                    if p: prod_counts[p]+=1

        ph_ids = set()
        for r in ph_recs: ph_ids.add(r["acc_id"] if r["acc_id"] else r["acc_name"])

        summary_rows.append({
            "employee_code":       code or None, "user_name":user,
            "period":period, "team":team,
            "territory":"; ".join(sorted(set(r["territory"] for r in urecs
                if r["territory"] and r["territory"].lower() not in ACC_TYPE_LABELS_LOWER))) or None,
            "is_manager":user in managers,
            "working_days":working_days, "complete_field_days":complete_field_days,
            "office_work_days":len(ow_dates),
            "no_activities":sum(1 for r in urecs if r["acc_type_category"]=="Activities"),
            "no_events":sum(1 for r in urecs if r["acc_type_category"]=="Events"),
            "double_visit_days":len(set(r["date"] for r in urecs if r["visit_type_category"]=="Double" and r["date"])),
            "coaching_days":coaching_by_mgr.get(user,0),
            "am_shift_days":am_days, "pm_shift_days":pm_days,
            "am_calls":am_calls, "pm_calls":pm_calls,
            "am_call_rate":round(am_calls/am_days,2) if am_days else 0,
            "pm_call_rate":round(pm_calls/pm_days,2) if pm_days else 0,
            "avg_am_start_time":avg_start or None,
            "am_unique_doctors":len(set(r["doctor_key"] for r in am_recs if r["doctor_key"])),
            "pm_unique_doctors":len(set(r["doctor_key"] for r in pm_recs if r["doctor_key"])),
            "total_am_covered":len(doc_by_cat.get("AM Center",set())|doc_by_cat.get("Hospital",set())),
            "total_pm_covered":len(doc_by_cat.get("Clinic",set())|doc_by_cat.get("Poly Clinics",set())),
            "clinic_covered":len(doc_by_cat.get("Clinic",set())),
            "polyclinic_covered":len(doc_by_cat.get("Poly Clinics",set())),
            "amcenter_covered":len(doc_by_cat.get("AM Center",set())),
            "hospital_covered":len(doc_by_cat.get("Hospital",set())),
            "pharmacies_visited":len(ph_recs), "pharmacies_covered":len(ph_ids-{""}),
            "total_product_calls":sum(prod_counts.values()),
            "distinct_products":len(prod_counts),
            "product_calls_detail":", ".join(f"{p}({c})" for p,c in sorted(prod_counts.items(),key=lambda x:-x[1])) or None,
            "avg_am_shift_hm":fmt_hm(avg_dur(am_recs)),
            "avg_pm_shift_hm":fmt_hm(avg_dur(pm_recs)),
            "avg_field_overall_hm":fmt_hm(avg_dur([r for r in urecs if r["shift"] in ("AM","PM")])),
            "total_visits":len(urecs), "upload_batch":batch,
        })

    # Specialty
    spec_groups = defaultdict(lambda:{"calls":0,"doctors":set()})
    for r in records:
        if r["shift"] not in ("AM","PM") or not r["specialty"] or not r["classification"]: continue
        k=(r["employee_code"] or r["user"],r["user"],r["specialty"],r["classification"],r["shift"])
        spec_groups[k]["calls"]+=1
        if r["doctor_key"]: spec_groups[k]["doctors"].add(r["doctor_key"])
    spec_rows=[{"employee_code":c or None,"user_name":u,"period":period,
                "specialty":sp,"classification":cl,"shift":sh,
                "call_count":agg["calls"],"unique_doctors":len(agg["doctors"]),"upload_batch":batch}
               for (c,u,sp,cl,sh),agg in sorted(spec_groups.items())]

    # Products
    prod_groups = defaultdict(lambda:{"calls":0,"doctors":set()})
    for r in records:
        if r["shift"] not in ("AM","PM") or not r["products"]: continue
        spec = r["specialty"] or "Unknown"
        for p in r["products"].split(","):
            p=p.strip()
            if not p: continue
            k=(r["employee_code"] or r["user"],r["user"],p,r["shift"],spec)
            prod_groups[k]["calls"]+=1
            if r["doctor_key"]: prod_groups[k]["doctors"].add(r["doctor_key"])
    prod_rows=[{"employee_code":c or None,"user_name":u,"period":period,
                "product":prod,"shift":sh,"specialty":sp,
                "call_count":agg["calls"],"unique_doctors":len(agg["doctors"]),"upload_batch":batch}
               for (c,u,prod,sh,sp),agg in sorted(prod_groups.items())]

    return summary_rows, coaching_rows, spec_rows, prod_rows


# ── Upload one folder ─────────────────────────────────────────────────────────

def upload_folder(folder, period, batch):
    print(f"\n{'='*60}")
    print(f"📤 Uploading: {folder}")
    print(f"   Period: {period} | Batch: {batch}")

    records = load_folder(folder)
    if not records:
        print("   ❌ No records found — skipping")
        return

    print(f"   Total records: {len(records)}")
    summary, coaching, spec, prod = compute_all(records, batch, period)

    print(f"   🗑️  Clearing old data for batch '{batch}'...")
    for table in ("summaries","coaching_days","specialty_classification","product_calls"):
        sb_delete(table, batch)

    print(f"   💾 Uploading...")
    print(f"      summaries:               {sb_insert('summaries', summary)} rows")
    print(f"      coaching_days:           {sb_insert('coaching_days', coaching)} rows")
    print(f"      specialty_classification:{sb_insert('specialty_classification', spec)} rows")
    print(f"      product_calls:           {sb_insert('product_calls', prod)} rows")
    print(f"   ✅ Done!")


# ── File watcher ──────────────────────────────────────────────────────────────

class FolderHandler(FileSystemEventHandler):
    def __init__(self, folder, period, batch):
        self.folder = folder
        self.period = period
        self.batch  = batch
        self._pending = False

    def on_any_event(self, event):
        if event.is_directory: return
        if not str(event.src_path).endswith(".xlsx"): return
        if not self._pending:
            self._pending = True
            # Debounce 5s — wait for all files to finish copying
            import threading
            def run():
                time.sleep(5)
                upload_folder(self.folder, self.period, self.batch)
                self._pending = False
            threading.Thread(target=run, daemon=True).start()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--last-month", default=r"E:\periods\last_month",
                        help="Folder for last month data")
    parser.add_argument("--recent",     default=r"E:\periods\recent",
                        help="Folder for recent (1-15) data")
    parser.add_argument("--once", action="store_true",
                        help="Upload once and exit (no watching)")
    args = parser.parse_args()

    if not SERVICE_KEY:
        print("❌ Set SUPABASE_SERVICE_ROLE_KEY environment variable.")
        sys.exit(1)

    folders = [
        (args.last_month, "Last Month", "last_month"),
        (args.recent,     "Recent",     "recent"),
    ]

    # Create folders if they don't exist
    for folder, _, _ in folders:
        Path(folder).mkdir(parents=True, exist_ok=True)
        print(f"✓ Folder ready: {folder}")

    # Always do an initial upload
    for folder, period, batch in folders:
        upload_folder(folder, period, batch)

    if args.once:
        print("\n✅ One-time upload complete.")
        return

    # Start watching
    print(f"\n👁️  Watching folders for new files... (Ctrl+C to stop)")
    observer = Observer()
    for folder, period, batch in folders:
        observer.schedule(FolderHandler(folder, period, batch), folder, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n⏹️  Stopped.")
    observer.join()


if __name__ == "__main__":
    main()
