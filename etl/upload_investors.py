#!/usr/bin/env python3
"""
Jutem Investors Upload Script
==============================
Reads investors_upload.csv and creates investor records in Directus.

Usage:
    python3 etl/upload_investors.py                # Dry run
    python3 etl/upload_investors.py --execute      # Actually upload
"""

import os
import sys
import csv
import argparse
import time
import requests

# ─── Configuration ───────────────────────────────────────────────────
DIRECTUS_URL = "https://zedlenders.pickmesms.com"
COMPANY_ID = 22

CSV_PATH = os.path.join(os.path.dirname(__file__), "output", "investors_upload.csv")
MAPPING_PATH = os.path.join(os.path.dirname(__file__), "output", "investor_name_to_id.csv")


def get_token():
    token = os.environ.get("DIRECTUS_TOKEN")
    if not token:
        token = input("Enter Directus admin token: ").strip()
    return token


def load_csv():
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found.")
        print("Run jutem_etl.py --step investors first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} investors from CSV")
    return rows


def check_existing(token):
    """Check which investors already exist."""
    headers = {"Authorization": f"Bearer {token}"}
    existing = {}
    res = requests.get(
        f"{DIRECTUS_URL}/items/jutem_investors",
        headers=headers,
        params={
            "filter[company][_eq]": COMPANY_ID,
            "fields": "id,name",
            "limit": -1,
        },
    )
    res.raise_for_status()
    for item in res.json().get("data", []):
        existing[item["name"].upper().strip()] = item["id"]
    return existing


def main():
    parser = argparse.ArgumentParser(description="Upload Jutem investors to Directus")
    parser.add_argument("--execute", action="store_true", help="Actually upload (default is dry run)")
    args = parser.parse_args()

    rows = load_csv()
    token = get_token()

    existing = check_existing(token) if args.execute else {}
    skipped = 0
    valid = []

    for row in rows:
        name = row["name"].strip()
        if name.upper() in existing:
            skipped += 1
            continue
        valid.append(row)

    print(f"\nValidation results:")
    print(f"  Valid (new): {len(valid)}")
    print(f"  Skipped (existing): {skipped}")

    if not args.execute:
        print("\n--- DRY RUN ---")
        for v in valid:
            print(f"  {v['name']}")
        print("\nRun with --execute to actually upload.")
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    # Start with existing mappings
    name_to_id = dict(existing)
    successes = 0
    errors = []

    for i, row in enumerate(valid):
        payload = {
            "name": row["name"].strip(),
            "status": row.get("status", "Active"),
            "company": COMPANY_ID,
        }

        try:
            res = requests.post(
                f"{DIRECTUS_URL}/items/jutem_investors",
                headers=headers,
                json=payload,
            )
            res.raise_for_status()
            created = res.json()["data"]
            name_to_id[row["name"].upper().strip()] = created["id"]
            successes += 1
            print(f"  [{i+1}/{len(valid)}] Created: {row['name']} (ID: {created['id']})")
        except requests.RequestException as e:
            err_msg = str(e)
            try:
                err_msg = res.json().get("errors", [{}])[0].get("message", err_msg)
            except Exception:
                pass
            errors.append({"name": row["name"], "error": err_msg})
            print(f"  [{i+1}/{len(valid)}] FAILED: {row['name']} - {err_msg}")

        if (i + 1) % 10 == 0:
            time.sleep(0.5)

    # Also include previously existing ones in mapping
    for row in rows:
        key = row["name"].upper().strip()
        if key in existing and key not in name_to_id:
            name_to_id[key] = existing[key]

    # Write mapping CSV
    with open(MAPPING_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["name", "id"])
        writer.writeheader()
        for name_key, inv_id in sorted(name_to_id.items()):
            writer.writerow({"name": name_key, "id": inv_id})

    print(f"\n{'='*50}")
    print("Upload complete!")
    print(f"  Created: {successes}")
    print(f"  Skipped: {skipped}")
    print(f"  Failed: {len(errors)}")
    print(f"  Mapping saved to: {MAPPING_PATH}")


if __name__ == "__main__":
    main()
