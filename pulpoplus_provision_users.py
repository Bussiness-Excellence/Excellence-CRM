"""
pulpoplus_provision_users.py

One-time (and re-runnable) script that creates real Supabase Auth login
accounts + app_users rows for everyone currently in the `hierarchy` table.

WHY THIS EXISTS: app_users.id is a foreign key to auth.users, so nobody can
log in to the web app until they have an actual Supabase Auth account. This
script bridges that gap using the Supabase Admin API.

LOGIN SCHEME (change the constants below if you want something different):
  - Username shown to the user: their Employee Code
  - Internally mapped to the fake email  "{code}@EMAIL_DOMAIN"
    (never needs to receive real mail — Supabase just needs email *format*)
  - Default password: the employee code itself
  - is_default_password = true, so the app can force a password-change
    screen on first login

SAFE TO RE-RUN: people who already have an app_users row are skipped, so
this won't reset anyone's password after they've changed it, and won't
create duplicate accounts.

SETUP:
    pip install requests --break-system-packages
    $env:SUPABASE_URL = "https://xxbfwvlqixnmonxytdxq.supabase.co"
    $env:SUPABASE_SERVICE_ROLE_KEY = "your-service-role-key"

USAGE:
    python pulpoplus_provision_users.py
    python pulpoplus_provision_users.py --dry-run   # preview without creating anything
"""

import os
import sys
import argparse
import requests

EMAIL_DOMAIN = "excellence-crm.internal"


def _config():
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        print("❌ Missing SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY environment variables.")
        sys.exit(1)
    return url.rstrip("/"), key


def _headers(key):
    return {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }


def fetch_all(url, key, table, select="*"):
    resp = requests.get(f"{url}/rest/v1/{table}",
                         headers=_headers(key), params={"select": select})
    resp.raise_for_status()
    return resp.json()


def create_auth_user(url, key, email, password):
    """Admin API: create a confirmed Auth user (no email verification needed
    since it's a synthetic address). Returns the new user's id, or None if
    something went wrong."""
    resp = requests.post(
        f"{url}/auth/v1/admin/users",
        headers=_headers(key),
        json={"email": email, "password": password, "email_confirm": True},
    )
    if resp.status_code in (200, 201):
        return resp.json().get("id")
    print(f"   ⚠️  Failed to create auth user for {email}: "
          f"{resp.status_code} {resp.text[:200]}")
    return None


def insert_app_user(url, key, row):
    resp = requests.post(
        f"{url}/rest/v1/app_users",
        headers={**_headers(key), "Prefer": "return=minimal"},
        json=row,
    )
    if not resp.ok:
        print(f"   ⚠️  Failed to insert app_users row for "
              f"{row['employee_code']}: {resp.status_code} {resp.text[:200]}")
        return False
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true",
                         help="Preview what would be created, without creating anything")
    args = parser.parse_args()

    url, key = _config()

    print("\n" + "=" * 65)
    print("PulpoPlus User Provisioning")
    print("=" * 65)

    print("\n📋 Loading hierarchy...")
    hierarchy_rows = fetch_all(url, key, "hierarchy",
                                select="id,team_id,employee_name,employee_code,role")
    print(f"   {len(hierarchy_rows)} hierarchy rows loaded")

    print("📋 Loading existing app_users...")
    existing = fetch_all(url, key, "app_users", select="employee_code")
    existing_codes = {r["employee_code"] for r in existing if r.get("employee_code")}
    print(f"   {len(existing_codes)} account(s) already exist — will be skipped")

    # De-dupe by employee_code (hierarchy can have the same person referenced
    # more than once, e.g. a BLM listed at the top of every team they manage)
    seen = set()
    to_create = []
    for r in hierarchy_rows:
        code = r.get("employee_code")
        if not code or code in seen or code in existing_codes:
            continue
        seen.add(code)
        to_create.append(r)

    print(f"\n👤 {len(to_create)} new account(s) to create")
    if args.dry_run:
        print("\n--dry-run: no changes will be made. Preview:")
        for r in to_create[:10]:
            print(f"   {r['employee_code']}  {r['employee_name']}  ({r['role']})"
                  f"  → login: {r['employee_code']}@{EMAIL_DOMAIN} / password: {r['employee_code']}")
        if len(to_create) > 10:
            print(f"   ... and {len(to_create) - 10} more")
        return

    created, failed = 0, 0
    for r in to_create:
        code = r["employee_code"]
        email = f"{code}@{EMAIL_DOMAIN}"
        auth_id = create_auth_user(url, key, email, password=code)
        if not auth_id:
            failed += 1
            continue

        ok = insert_app_user(url, key, {
            "id": auth_id,
            "employee_code": code,
            "employee_name": r["employee_name"],
            "role": r.get("role") or "MR",
            "team_id": r.get("team_id"),
            "hierarchy_id": r.get("id"),
            "is_active": True,
            "is_default_password": True,
        })
        if ok:
            created += 1
        else:
            failed += 1

    print(f"\n✅ Done — {created} account(s) created, {failed} failed.")
    print(f"\nEveryone logs in with their Employee Code as both username and "
          f"initial password (e.g. code 13512 → username 13512, password 13512), "
          f"and will be asked to set a new password on first login.")


if __name__ == "__main__":
    main()
