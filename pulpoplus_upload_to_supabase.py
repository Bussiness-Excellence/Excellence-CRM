#!/usr/bin/env python3
"""
pulpoplus_upload_to_supabase.py
────────────────────────────────
Reads one or more rebuilt Excel files (each with a "Raw Data" sheet),
computes all KPIs from scratch, and pushes everything to Supabase:

  • summaries            — one row per user per period
  • coaching_days        — one row per manager/rep/date coaching day
  • specialty_classification — calls per user/specialty/class/shift
  • product_calls        — calls per user/product/shift/specialty

Safe to re-run: deletes existing rows for the same upload_batch first,
then inserts fresh data — so you can re-upload after a fix without
leaving duplicate rows.

SETUP:
    pip install pandas openpyxl requests
    set SUPABASE_URL=https://xxbfwvlqixnmonxytdxq.supabase.co
    set SUPABASE_SERVICE_ROLE_KEY=your-service-role-key

USAGE:
    python pulpoplus_upload_to_supabase.py file1.xlsx file2.xlsx --period "Jun 2026"
    python pulpoplus_upload_to_supabase.py *.xlsx --period "Jun 2026" --dry-run
"""

import os, sys, argparse, math, requests
from collections import defaultdict, Counter
from datetime import datetime

import pandas as pd

# ── Supabase config ────────────────────────────────────────────────────────────
SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://xxbfwvlqixnmonxytdxq.supabase.co")
SERVICE_KEY  = os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4Mjc1NjE2NSwiZXhwIjoyMDk4MzMyMTY1fQ.PSk6RyFmg_OFTcCtYO74AeJj6wT4FGZS2K2JT9GEJ_A)"
BATCH_SIZE   = 500   # rows per POST to Supabase REST API

# ── Domain constants (must match rebuild script) ──────────────────────────────
ACC_TYPE_LABELS_LOWER = {
    "clinic","hospital","am center","poly clinics","office work",
    "pharmacy","distributors","activities","events","activity","event",
    "p","h","c","am","pc","d","ow","a","e",
}
SHIFT_MAP = {"Clinic":"PM","Poly Clinics":"PM","Hospital":"AM","AM Center":"AM"}
COACHING_TIME_TOLERANCE_MINUTES = 90


# ── helpers ───────────────────────────────────────────────────────────────────

def _headers():
    return {
        "apikey": SERVICE_KEY,
        "Authorization": f"Bearer {SERVICE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }

def sb_delete(table, batch_col, batch_val):
    r = requests.delete(
        f"{SUPABASE_URL}/rest/v1/{table}",
        headers=_headers(),
        params={batch_col: f"eq.{batch_val}"}
    )
    if not r.ok:
        print(f"  ⚠️  Delete from {table} failed: {r.status_code} {r.text[:200]}")

def sb_insert(table, rows, dry_run=False):
    if not rows: return
    if dry_run:
        print(f"  [dry-run] would insert {len(rows)} rows into {table}")
        return
    for i in range(0, len(rows), BATCH_SIZE):
        chunk = rows[i:i+BATCH_SIZE]
        r = requests.post(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers=_headers(),
            json=chunk,
        )
        if not r.ok:
            print(f"  ⚠️  Insert into {table} failed at chunk {i}: {r.status_code} {r.text[:300]}")
            return
    print(f"  ✓ {table}: {len(rows)} rows inserted")

def safe(v):
    if v is None: return None
    if isinstance(v, float) and math.isnan(v): return None
    s = str(v).strip()
    return None if s.lower() in ("nan","none","") else s

def parse_members(raw):
    if not raw: return []
    return [n.strip() for n in str(raw).split(",") if n.strip() and n.strip().lower() not in ("nan","none")]

def minutes_between(t1, t2):
    for fmt in ("%H:%M:%S","%H:%M","%I:%M %p"):
        try:
            a = datetime.strptime(t1.strip(), fmt)
            b = datetime.strptime(t2.strip(), fmt)
            return abs((b-a).total_seconds()/60)
        except: continue
    return 999

def fmt_hm(total_min):
    if not total_min: return "0:00"
    h = int(total_min//60); m = int(total_min%60)
    return f"{h}:{m:02d}"


# ── load Raw Data ─────────────────────────────────────────────────────────────

def load_records(paths):
    all_records = []
    for path in paths:
        try:
            df = pd.read_excel(path, sheet_name="Raw Data", engine="openpyxl")
        except Exception as e:
            print(f"  ⚠️  Cannot read {path}: {e}")
            continue

        # Build employee_code lookup from hierarchy if visits table has it
        for _, row in df.iterrows():
            user = safe(row.get("user"))
            if not user: continue
            rec = {
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
                # members: rebuild from user column (single name per row in export)
                "members":            [user],
            }
            all_records.append(rec)
        print(f"  ✓ {len(df)} rows from {path}")

    # Build employee_code map from visits table in Supabase
    print("  📋 Fetching employee_code map from Supabase visits...")
    code_map = {}
    try:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/visits",
            headers=_headers(),
            params={"select":"user,employee_code", "limit":50000}
        )
        if r.ok:
            for row in r.json():
                n = (row.get("user") or "").strip()
                c = (row.get("employee_code") or "").strip()
                if n and c and n not in code_map:
                    code_map[n] = c
            print(f"  ✓ Code map: {len(code_map)} names resolved")
    except Exception as e:
        print(f"  ⚠️  Could not fetch code map: {e}")

    # Also fetch from hierarchy as fallback
    try:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/hierarchy",
            headers=_headers(),
            params={"select":"employee_name,employee_code", "limit":10000}
        )
        if r.ok:
            for row in r.json():
                n = (row.get("employee_name") or "").strip()
                c = (row.get("employee_code") or "").strip()
                if n and c and n not in code_map:
                    code_map[n] = c
    except: pass

    # Patch employee_codes into records
    for rec in all_records:
        if not rec["employee_code"] and rec["user"] in code_map:
            rec["employee_code"] = code_map[rec["user"]]

    return all_records


# ── territory backfill ────────────────────────────────────────────────────────

def backfill_territory(records):
    user_date_terr = {}
    user_terr_counts = defaultdict(Counter)
    for r in records:
        t = r["territory"]
        if t and t.lower() not in ACC_TYPE_LABELS_LOWER:
            user_terr_counts[r["user"]][t] += 1
            key = (r["user"], r["date"])
            if key not in user_date_terr:
                user_date_terr[key] = t
    for r in records:
        if r["territory"] and r["territory"].lower() not in ACC_TYPE_LABELS_LOWER:
            continue
        r["territory"] = ""
        key = (r["user"], r["date"])
        if key in user_date_terr:
            r["territory"] = user_date_terr[key]
        elif user_terr_counts[r["user"]]:
            r["territory"] = user_terr_counts[r["user"]].most_common(1)[0][0]


# ── manager identification ─────────────────────────────────────────────────────

def identify_managers(records):
    terr_by_user = defaultdict(set)
    for r in records:
        t = r["territory"]
        if t and t.lower() not in ACC_TYPE_LABELS_LOWER:
            terr_by_user[r["user"]].add(t)
    managers = {u for u, ts in terr_by_user.items() if len(ts) > 1}

    # Also check hierarchy roles
    try:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/hierarchy",
            headers=_headers(),
            params={"select":"employee_name,role","limit":10000}
        )
        if r.ok:
            for row in r.json():
                if row.get("role","").strip() in ("Supervisor","Area Manager","BLM"):
                    managers.add((row.get("employee_name") or "").strip())
    except: pass
    return managers


# ── coaching days ─────────────────────────────────────────────────────────────

def compute_coaching_days(records, managers, batch, team_map):
    matched = defaultdict(set)
    visits_by_key = defaultdict(list)
    for r in records:
        if r["shift"] not in ("AM","PM") or not r["date"] or not r["acc_id"]: continue
        visits_by_key[(r["date"], r["shift"], r["acc_id"])].append((r["user"], r["time"], r["employee_code"]))

    for (date, shift, acc_id), visits in visits_by_key.items():
        mgr_v   = [(u,t,c) for u,t,c in visits if u in managers]
        rep_v   = [(u,t,c) for u,t,c in visits if u not in managers]
        for mu,mt,mc in mgr_v:
            for ru,rt,rc in rep_v:
                if mt and rt and minutes_between(mt,rt) > COACHING_TIME_TOLERANCE_MINUTES: continue
                matched[(mu,mc,ru,rc,date)].add(shift)

    for r in records:
        if r["shift"] not in ("AM","PM") or not r["date"]: continue
        members = r["members"]
        if len(members) < 2: continue
        mgrs = [m for m in members if m in managers]
        reps = [m for m in members if m not in managers]
        for m in mgrs:
            for rep in reps:
                mc = team_map.get(m,{}).get("code","")
                rc = team_map.get(rep,{}).get("code","")
                matched[(m,mc,rep,rc,r["date"])].add(r["shift"])

    coaching_by_manager = defaultdict(int)
    rows = []
    for (mgr,mc,rep,rc,date), shifts in matched.items():
        if {"AM","PM"}.issubset(shifts):
            coaching_by_manager[mgr] += 1
            rows.append({
                "manager_name":  mgr,
                "manager_code":  mc or None,
                "rep_name":      rep,
                "rep_code":      rc or None,
                "coaching_date": date,
                "team":          team_map.get(mgr,{}).get("team",""),
                "upload_batch":  batch,
            })

    rows.sort(key=lambda x: (x["manager_name"], x["coaching_date"], x["rep_name"]))
    return dict(coaching_by_manager), rows


# ── summary computation ───────────────────────────────────────────────────────

def compute_summaries(records, managers, coaching_by_manager, batch, period, team_map):
    by_user = defaultdict(list)
    for r in records:
        by_user[r["user"]].append(r)

    rows = []
    for user, urecs in sorted(by_user.items()):
        code = urecs[0]["employee_code"] or team_map.get(user,{}).get("code","")
        team = urecs[0]["team"] or team_map.get(user,{}).get("team","")

        territories = sorted(set(
            r["territory"] for r in urecs
            if r["territory"] and r["territory"].lower() not in ACC_TYPE_LABELS_LOWER
        ))
        territory_str = "; ".join(territories)

        am_recs = [r for r in urecs if r["shift"]=="AM"]
        pm_recs = [r for r in urecs if r["shift"]=="PM"]
        ow_recs = [r for r in urecs if r["acc_type_category"]=="Office Work"]
        ph_recs = [r for r in urecs if r["acc_type_category"]=="Pharmacy"]

        am_dates = set(r["date"] for r in am_recs if r["date"])
        pm_dates = set(r["date"] for r in pm_recs if r["date"])
        ow_dates = set(r["date"] for r in ow_recs if r["date"])
        slot_dates = am_dates | ow_dates
        working_days = len(set(r["date"] for r in urecs if r["date"]))
        complete_field_days = round((len(slot_dates)+len(pm_dates))/2, 1)

        dv_dates = set(r["date"] for r in urecs if r["date"] and r["visit_type_category"]=="Double")
        coaching_days = coaching_by_manager.get(user, 0)
        office_work_days = len(ow_dates)

        # Doctor coverage
        doc_by_cat = defaultdict(set)
        for r in am_recs+pm_recs:
            if r["doctor_key"] and r["acc_type_category"] in ("Clinic","Poly Clinics","Hospital","AM Center"):
                doc_by_cat[r["acc_type_category"]].add(r["doctor_key"])

        am_calls = len(am_recs)
        pm_calls = len(pm_recs)
        am_shift_days = len(am_dates)
        pm_shift_days = len(pm_dates)
        pm_call_rate = round(pm_calls/pm_shift_days,2) if pm_shift_days else 0.0
        am_call_rate = round(am_calls/am_shift_days,2) if am_shift_days else 0.0

        # Pharmacy
        ph_visited = len(ph_recs)
        ph_unique = set()
        for r in ph_recs:
            ph_unique.add(r["acc_id"] if r["acc_id"] else r["acc_name"])
        ph_covered = len(ph_unique - {""})

        # Shift durations
        def avg_dur(recs):
            by_day = defaultdict(list)
            for r in recs:
                if r["date"] and r["time"]: by_day[r["date"]].append(r["time"])
            durs = []
            for times in by_day.values():
                ts = sorted(times)
                if len(ts)>=2: durs.append(minutes_between(ts[0],ts[-1]))
                else: durs.append(0.0)
            return sum(durs)/len(durs) if durs else 0.0

        am_dur = avg_dur(am_recs)
        pm_dur = avg_dur(pm_recs)
        overall_dur = avg_dur([r for r in urecs if r["shift"] in ("AM","PM")])

        # AM avg start time
        am_start_mins = []
        for d in am_dates:
            day_times = sorted(r["time"] for r in am_recs if r["date"]==d and r["time"])
            if day_times:
                for fmt in ("%H:%M:%S","%H:%M"):
                    try:
                        dt = datetime.strptime(day_times[0].strip(), fmt)
                        am_start_mins.append(dt.hour*60+dt.minute)
                        break
                    except: continue
        avg_start = ""
        if am_start_mins:
            avg = sum(am_start_mins)/len(am_start_mins)
            avg_start = f"{int(avg//60):02d}:{int(avg%60):02d}"

        # Products
        prod_counts = defaultdict(int)
        for r in urecs:
            if r["shift"] in ("AM","PM") and r["products"]:
                for p in r["products"].split(","):
                    p = p.strip()
                    if p: prod_counts[p]+=1
        prod_detail = ", ".join(f"{p}({c})" for p,c in sorted(prod_counts.items(), key=lambda x:-x[1]))

        rows.append({
            "employee_code":       code or None,
            "user_name":           user,
            "period":              period,
            "team":                team,
            "territory":           territory_str or None,
            "is_manager":          user in managers,
            "working_days":        working_days,
            "complete_field_days": complete_field_days,
            "office_work_days":    office_work_days,
            "no_activities":       sum(1 for r in urecs if r["acc_type_category"]=="Activities"),
            "no_events":           sum(1 for r in urecs if r["acc_type_category"]=="Events"),
            "double_visit_days":   len(dv_dates),
            "coaching_days":       coaching_days,
            "am_shift_days":       am_shift_days,
            "pm_shift_days":       pm_shift_days,
            "am_calls":            am_calls,
            "pm_calls":            pm_calls,
            "am_call_rate":        am_call_rate,
            "pm_call_rate":        pm_call_rate,
            "avg_am_start_time":   avg_start or None,
            "am_unique_doctors":   len(set(r["doctor_key"] for r in am_recs if r["doctor_key"])),
            "pm_unique_doctors":   len(set(r["doctor_key"] for r in pm_recs if r["doctor_key"])),
            "total_am_covered":    len(doc_by_cat.get("AM Center",set())|doc_by_cat.get("Hospital",set())),
            "total_pm_covered":    len(doc_by_cat.get("Clinic",set())|doc_by_cat.get("Poly Clinics",set())),
            "clinic_covered":      len(doc_by_cat.get("Clinic",set())),
            "polyclinic_covered":  len(doc_by_cat.get("Poly Clinics",set())),
            "amcenter_covered":    len(doc_by_cat.get("AM Center",set())),
            "hospital_covered":    len(doc_by_cat.get("Hospital",set())),
            "pharmacies_visited":  ph_visited,
            "pharmacies_covered":  ph_covered,
            "total_product_calls": sum(prod_counts.values()),
            "distinct_products":   len(prod_counts),
            "product_calls_detail":prod_detail or None,
            "avg_am_shift_hm":     fmt_hm(am_dur),
            "avg_pm_shift_hm":     fmt_hm(pm_dur),
            "avg_field_overall_hm":fmt_hm(overall_dur),
            "avg_am_duration_min": int(am_dur),
            "avg_pm_duration_min": int(pm_dur),
            "total_visits":        len(urecs),
            "upload_batch":        batch,
        })

    return rows


# ── specialty × classification ────────────────────────────────────────────────

def compute_specialty(records, batch, period, team_map):
    groups = defaultdict(lambda:{"calls":0,"doctors":set()})
    for r in records:
        if r["shift"] not in ("AM","PM") or not r["specialty"] or not r["classification"]: continue
        key = (r["employee_code"] or r["user"], r["user"], r["specialty"], r["classification"], r["shift"])
        groups[key]["calls"]+=1
        if r["doctor_key"]: groups[key]["doctors"].add(r["doctor_key"])
    rows = []
    for (code,uname,spec,cls,shift),agg in sorted(groups.items()):
        rows.append({
            "employee_code": code or None,
            "user_name":     uname,
            "period":        period,
            "specialty":     spec,
            "classification":cls,
            "shift":         shift,
            "call_count":    agg["calls"],
            "unique_doctors":len(agg["doctors"]),
            "upload_batch":  batch,
        })
    return rows


# ── product calls ─────────────────────────────────────────────────────────────

def compute_products(records, batch, period):
    groups = defaultdict(lambda:{"calls":0,"doctors":set()})
    for r in records:
        if r["shift"] not in ("AM","PM") or not r["products"]: continue
        spec = r["specialty"] or "Unknown"
        code = r["employee_code"] or r["user"]
        for p in r["products"].split(","):
            p = p.strip()
            if not p: continue
            key = (code, r["user"], p, r["shift"], spec)
            groups[key]["calls"]+=1
            if r["doctor_key"]: groups[key]["doctors"].add(r["doctor_key"])
    rows = []
    for (code,uname,product,shift,spec),agg in sorted(groups.items()):
        rows.append({
            "employee_code": code or None,
            "user_name":     uname,
            "period":        period,
            "product":       product,
            "shift":         shift,
            "specialty":     spec,
            "call_count":    agg["calls"],
            "unique_doctors":len(agg["doctors"]),
            "upload_batch":  batch,
        })
    return rows


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("inputs", nargs="+", help="Excel files with Raw Data sheet")
    parser.add_argument("--period", required=True,
                        help='Human-readable period label e.g. "Jun 2026" or "2026-06-01:2026-06-15"')
    parser.add_argument("--batch",  default=None,
                        help="Upload batch key (defaults to period, used for dedup)")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    batch = args.batch or args.period

    if not SERVICE_KEY:
        print("❌ Set SUPABASE_SERVICE_ROLE_KEY environment variable first.")
        sys.exit(1)

    print("\n" + "="*65)
    print("PulpoPlus → Supabase Uploader")
    print("="*65)
    print(f"  Period : {args.period}")
    print(f"  Batch  : {batch}")
    print(f"  Files  : {', '.join(args.inputs)}")
    if args.dry_run: print("  Mode   : DRY RUN (no changes)")
    print()

    print("📥 Loading Raw Data...")
    records = load_records(args.inputs)
    print(f"  Total records: {len(records)}")
    if not records:
        print("❌ No records found."); sys.exit(1)

    # Build name → {code, team} map from records
    team_map = {}
    for r in records:
        if r["user"] and r["user"] not in team_map:
            team_map[r["user"]] = {"code": r["employee_code"], "team": r["team"]}

    print("\n🌍 Backfilling territory...")
    backfill_territory(records)

    print("\n👔 Identifying managers...")
    managers = identify_managers(records)
    print(f"  {len(managers)} managers identified")

    print("\n🤝 Computing coaching days...")
    coaching_by_mgr, coaching_rows = compute_coaching_days(records, managers, batch, team_map)
    print(f"  {len(coaching_rows)} coaching day rows")
    for m,d in sorted(coaching_by_mgr.items()):
        print(f"    {m}: {d} day(s)")

    print("\n📊 Computing summaries...")
    summary_rows = compute_summaries(records, managers, coaching_by_mgr, batch, args.period, team_map)
    print(f"  {len(summary_rows)} user summaries")

    print("\n🔬 Computing specialty × classification...")
    spec_rows = compute_specialty(records, batch, args.period, team_map)
    print(f"  {len(spec_rows)} rows")

    print("\n💊 Computing product calls...")
    prod_rows = compute_products(records, batch, args.period)
    print(f"  {len(prod_rows)} rows")

    if args.dry_run:
        print("\n[dry-run] Skipping all database writes.")
        print("✅ Dry run complete.")
        return

    print(f"\n🗑️  Clearing existing data for batch '{batch}'...")
    for table in ("summaries","coaching_days","specialty_classification","product_calls"):
        sb_delete(table, "upload_batch", batch)

    print("\n💾 Uploading to Supabase...")
    sb_insert("summaries",               summary_rows)
    sb_insert("coaching_days",           coaching_rows)
    sb_insert("specialty_classification",spec_rows)
    sb_insert("product_calls",           prod_rows)

    print(f"\n✅ Done! Period '{args.period}' is now live in the app.")


if __name__ == "__main__":
    main()
