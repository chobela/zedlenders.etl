#!/usr/bin/env python3
"""
Jutem Borrowers Upload Script
==============================
Reads borrowers_upload.csv and creates borrower users in Directus.

- Checks for duplicate NRC before creating
- Creates users with Borrower role
- Skips rows that already exist

Usage:
    python3 etl/upload_borrowers.py                    # Dry run (default)
    python3 etl/upload_borrowers.py --execute          # Actually upload
    python3 etl/upload_borrowers.py --execute --batch 10  # Upload in batches of 10
"""

import os
import sys
import csv
import argparse
import time
import requests

# ─── Configuration ───────────────────────────────────────────────────────────

DIRECTUS_URL = "https://zedlenders.pickmesms.com"
COMPANY_ID = 22
BORROWER_ROLE = "ab3c07d3-a225-4b68-8454-96e00116e307"
DEFAULT_PASSWORD = "123456"

CSV_PATH = os.path.join(os.path.dirname(__file__), "output", "borrowers_upload.csv")


def get_token():
    """Get token from environment or prompt."""
    token = os.environ.get("DIRECTUS_TOKEN")
    if not token:
        token = input("Enter Directus admin token: ").strip()
    return token


def check_existing_nrcs(token, nrcs):
    """Check which NRCs already exist in the system. Returns set of existing NRCs."""
    existing = set()
    headers = {"Authorization": f"Bearer {token}"}

    # Check in batches to avoid URL length limits
    batch_size = 20
    for i in range(0, len(nrcs), batch_size):
        batch = nrcs[i:i + batch_size]
        try:
            res = requests.get(
                f"{DIRECTUS_URL}/users",
                headers=headers,
                params={
                    "filter[nrc][_in]": ",".join(batch),
                    "fields": "nrc",
                    "limit": -1,
                },
            )
            res.raise_for_status()
            for user in res.json().get("data", []):
                if user.get("nrc"):
                    existing.add(user["nrc"])
        except requests.RequestException as e:
            print(f"  WARNING: Failed to check NRC batch: {e}")

    return existing


def load_csv():
    """Load and parse the borrowers CSV."""
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found. Run jutem_etl.py --step borrowers first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} borrowers from CSV")
    return rows


def build_payload(row):
    """Build Directus user payload from CSV row."""
    return {
        "first_name": row["first_name"],
        "last_name": row["last_name"],
        "email": row["email"],
        "phone": str(row["phone"]),
        "nrc": row["nrc"],
        "address": row.get("address", "Lusaka"),
        "province": row.get("province", "Lusaka"),
        "employment_type": row.get("employment_type", "self_employed"),
        "industry": "Business",
        "business": row.get("business", "N/A"),
        "department": row.get("department", "N/A"),
        "employee_number": row.get("reference_number", ""),
        "next_kin_name": row.get("next_kin_name", "N/A"),
        "next_kin_phone": row.get("next_kin_phone", "260977000000"),
        "next_kin_relationship": row.get("next_kin_relationship", "N/A"),
        "company": COMPANY_ID,
        "role": BORROWER_ROLE,
        "status": "active",
        "password": DEFAULT_PASSWORD,
        "email_notifications": True,
        "provider": "default",
        "credit_score": 650,
        "documents": [],
    }


def main():
    parser = argparse.ArgumentParser(description="Upload Jutem borrowers to Directus")
    parser.add_argument("--execute", action="store_true", help="Actually upload (default is dry run)")
    parser.add_argument("--batch", type=int, default=1, help="Batch size for uploads (default: 1, one at a time)")
    args = parser.parse_args()

    rows = load_csv()
    token = get_token()

    # Check which NRCs already exist
    all_nrcs = [r["nrc"] for r in rows]
    print("Checking for existing borrowers...")
    existing_nrcs = check_existing_nrcs(token, all_nrcs)
    print(f"  Found {len(existing_nrcs)} already in system")

    # Filter to new borrowers only
    new_rows = [r for r in rows if r["nrc"] not in existing_nrcs]
    skipped = len(rows) - len(new_rows)
    print(f"  Skipping {skipped} duplicates")
    print(f"  {len(new_rows)} new borrowers to upload")

    if not args.execute:
        print("\n--- DRY RUN ---")
        print("First 5 new borrowers:")
        for r in new_rows[:5]:
            print(f"  {r['first_name']} {r['last_name']} (NRC: {r['nrc']})")
        if len(new_rows) > 5:
            print(f"  ... and {len(new_rows) - 5} more")
        print("\nRun with --execute to actually upload.")
        return

    # Upload
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    successes = 0
    errors = []

    for i, row in enumerate(new_rows):
        payload = build_payload(row)
        name = f"{row['first_name']} {row['last_name']}"

        try:
            res = requests.post(
                f"{DIRECTUS_URL}/users",
                headers=headers,
                json=payload,
            )
            res.raise_for_status()
            successes += 1
            print(f"  [{i+1}/{len(new_rows)}] Created: {name} (NRC: {row['nrc']})")
        except requests.RequestException as e:
            error_msg = ""
            try:
                error_msg = res.json().get("errors", [{}])[0].get("message", str(e))
            except Exception:
                error_msg = str(e)
            errors.append({"name": name, "nrc": row["nrc"], "error": error_msg})
            print(f"  [{i+1}/{len(new_rows)}] FAILED: {name} - {error_msg}")

        # Small delay to avoid rate limiting
        if (i + 1) % 10 == 0:
            time.sleep(0.5)

    # Summary
    print(f"\n{'='*50}")
    print(f"Upload complete!")
    print(f"  Created: {successes}")
    print(f"  Skipped (existing): {skipped}")
    print(f"  Failed: {len(errors)}")

    if errors:
        errors_path = os.path.join(os.path.dirname(__file__), "output", "borrowers_upload_errors.csv")
        with open(errors_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["name", "nrc", "error"])
            writer.writeheader()
            writer.writerows(errors)
        print(f"  Error details saved to: {errors_path}")


if __name__ == "__main__":
    main()
