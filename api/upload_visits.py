import os
import math
import requests
import pandas as pd
from flask import Flask, request, jsonify
from collections import defaultdict, Counter
from datetime import datetime

app = Flask(__name__)

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://xxbfwvlqixnmonxytdxq.supabase.co")
SERVICE_KEY  = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4Mjc1NjE2NSwiZXhwIjoyMDk4MzMyMTY1fQ.PSk6RyFmg_OFTcCtYO74AeJj6wT4FGZS2K2JT9GEJ_A)")
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

    try:
        r = requests.get(f"{SUPABASE_URL}/rest/v1/hierarchy",
                         headers=headers(),
                         params={"select":"employee_name,role","limit":10000})
        if r.ok:
            for row in r.json():
                if row.get("role","") in ("Supervisor","Area Manager","BLM"):
                    managers.add((row.get("employee_name") or "").strip())
    except: pass

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

def provision_new_users(records):
    resp = requests.get(f"{SUPABASE_URL}/rest/v1/teams?select=id,name", headers=headers())
    if not resp.ok: return
    teams_dict = {t["name"].strip().upper(): t["id"] for t in resp.json()}

    resp = requests.get(f"{SUPABASE_URL}/rest/v1/app_users?select=employee_code", headers=headers())
    if not resp.ok: return
    existing_codes = {u["employee_code"].strip() for u in resp.json() if u.get("employee_code")}

    unique_users = {}
    for r in records:
        code = r["employee_code"]
        name = r["user"]
        team = r["team"]
        if code and name:
            unique_users[code.strip()] = {"name": name.strip(), "team": team.strip()}

    for code, info in unique_users.items():
        if code in existing_codes: continue
        
        team_name = info["team"] or "UNKNOWN"
        team_upper = team_name.upper()
        if team_upper not in teams_dict:
            t_resp = requests.post(f"{SUPABASE_URL}/rest/v1/teams", headers={**headers(), "Prefer": "return=representation"}, json={"name": team_name})
            if t_resp.ok and t_resp.json():
                team_id = t_resp.json()[0]["id"]
                teams_dict[team_upper] = team_id
            else: team_id = None
        else:
            team_id = teams_dict[team_upper]

        email = f"{code}@excellence-crm.internal"
        auth_resp = requests.post(f"{SUPABASE_URL}/auth/v1/admin/users",
            headers={"apikey": SERVICE_KEY, "Authorization": f"Bearer {SERVICE_KEY}", "Content-Type": "application/json"},
            json={"email": email, "password": code, "email_confirm": True})
        if not auth_resp.ok: continue
        
        auth_id = auth_resp.json().get("id")
        if not auth_id: continue

        h_resp = requests.post(f"{SUPABASE_URL}/rest/v1/hierarchy",
            headers={**headers(), "Prefer": "return=representation"},
            json={"employee_name": info["name"], "employee_code": code, "role": "MR", "team_id": team_id, "supervisor_name": None})
        hierarchy_id = h_resp.json()[0]["id"] if h_resp.ok and h_resp.json() else None

        requests.post(f"{SUPABASE_URL}/rest/v1/app_users", headers=headers(),
            json={"id": auth_id, "employee_code": code, "employee_name": info["name"], "role": "MR", "team_id": team_id, "hierarchy_id": hierarchy_id, "is_active": True, "is_default_password": True})

@app.route('/api/upload_visits', methods=['POST'])
def upload_visits():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    period = request.form.get('period', 'Unknown Period')
    batch = request.form.get('batch', 'unknown_batch')
    
    try:
        # Load the file from memory
        df = pd.read_excel(file.stream, sheet_name="Raw Data", engine="openpyxl")
        
        records = []
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
            
        if not records:
            return jsonify({"error": "No valid records found in Raw Data"}), 400
            
        provision_new_users(records)
        summary, coaching, spec, prod = compute_all(records, batch, period)

        for table in ("summaries","coaching_days","specialty_classification","product_calls"):
            sb_delete(table, batch)

        stats = {
            "summaries": sb_insert('summaries', summary),
            "coaching_days": sb_insert('coaching_days', coaching),
            "specialty_classification": sb_insert('specialty_classification', spec),
            "product_calls": sb_insert('product_calls', prod)
        }
        
        return jsonify({
            "message": "Upload successful",
            "period": period,
            "batch": batch,
            "rows_inserted": stats
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=3001)
