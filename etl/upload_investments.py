#!/usr/bin/env python3
"""
Jutem Investments Upload Script
================================
Reads investments_upload.csv and creates investment records in Directus.

Usage:
    python3 etl/upload_investments.py                # Dry run
    python3 etl/upload_investments.py --execute      # Actually upload
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

CSV_PATH = os.path.join(os.path.dirname(__file__), "output", "investments_upload.csv")
INVESTOR_MAPPING = os.path.join(os.path.dirname(__file__), "output", "investor_name_to_id.csv")
OUTPUT_MAPPING = os.path.join(os.path.dirname(__file__), "output", "investment_ref_to_id.csv")


def get_token():
    token = os.environ.get("DIRECTUS_TOKEN")
    if not token:
        token = input("Enter Directus admin token: ").strip()
    return token


def load_csv():
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found.")
        print("Run jutem_etl.py --step investments first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} investments from CSV")
    return rows


def load_investor_mapping():
    if not os.path.exists(INVESTOR_MAPPING):
        print(f"ERROR: {INVESTOR_MAPPING} not found.")
        print("Run upload_investors.py --execute first.")
        sys.exit(1)

    mapping = {}
    with open(INVESTOR_MAPPING, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            mapping[row["name"].upper().strip()] = int(row["id"])

    print(f"Loaded {len(mapping)} investor mappings")
    return mapping


def check_existing(token, reference_numbers):
    """Check which investment references already exist."""
    headers = {"Authorization": f"Bearer {token}"}
    existing = set()

    ref_list = list(reference_numbers)
    batch_size = 50
    for i in range(0, len(ref_list), batch_size):
        batch = ref_list[i : i + batch_size]
        refs_str = ",".join(batch)
        res = requests.get(
            f"{DIRECTUS_URL}/items/jutem_investments",
            headers=headers,
            params={
                "filter[reference_number][_in]": refs_str,
                "filter[company][_eq]": COMPANY_ID,
                "fields": "id,reference_number",
                "limit": -1,
            },
        )
        res.raise_for_status()
        for item in res.json().get("data", []):
            ref = item.get("reference_number")
            if ref:
                existing.add(ref)

    return existing


def delete_existing_investments(token):
    """Delete all jutem_investments for COMPANY_ID."""
    headers = {"Authorization": f"Bearer {token}"}
    print(f"\nDeleting existing jutem_investments for company {COMPANY_ID}...")

    # Fetch all IDs first
    res = requests.get(
        f"{DIRECTUS_URL}/items/jutem_investments",
        headers=headers,
        params={
            "filter[company][_eq]": COMPANY_ID,
            "fields": "id",
            "limit": -1,
        },
    )
    res.raise_for_status()
    items = res.json().get("data", [])

    if not items:
        print("  No existing investments found.")
        return

    ids = [item["id"] for item in items]
    print(f"  Found {len(ids)} investments to delete.")

    # Delete in batches
    batch_size = 50
    for i in range(0, len(ids), batch_size):
        batch = ids[i : i + batch_size]
        for item_id in batch:
            res = requests.delete(
                f"{DIRECTUS_URL}/items/jutem_investments/{item_id}",
                headers=headers,
            )
            res.raise_for_status()
        print(f"  Deleted {min(i + batch_size, len(ids))}/{len(ids)}")
        time.sleep(0.5)

    print(f"  Successfully deleted {len(ids)} investments.")


def main():
    parser = argparse.ArgumentParser(description="Upload Jutem investments to Directus")
    parser.add_argument("--execute", action="store_true", help="Actually upload (default is dry run)")
    parser.add_argument("--delete-existing", action="store_true",
                        help="Delete all existing jutem_investments for this company before uploading")
    args = parser.parse_args()

    rows = load_csv()
    token = get_token()

    # Delete existing investments if requested
    if args.execute and args.delete_existing:
        delete_existing_investments(token)

    investor_map = load_investor_mapping() if args.execute else {}

    all_refs = {row.get("reference_number", "") for row in rows}
    existing_refs = check_existing(token, all_refs) if args.execute else set()

    valid = []
    errors = []
    skipped = 0

    for i, row in enumerate(rows):
        ref = row.get("reference_number", "")
        if ref in existing_refs:
            skipped += 1
            continue

        amount = float(row.get("principal_amount", 0))
        if amount == 0:
            errors.append({"row": i + 2, "ref": ref, "error": "Zero principal"})
            continue

        # Look up investor ID
        inv_key = row["investor_name"].upper().strip()
        investor_id = investor_map.get(inv_key) if args.execute else None

        if args.execute and not investor_id:
            errors.append({"row": i + 2, "ref": ref, "error": f"Investor not found: {row['investor_name']}"})
            continue

        valid.append({
            "row": i + 2,
            "reference_number": ref,
            "investor_name": row["investor_name"],
            "investor_id": investor_id,
            "principal_amount": amount,
            "investment_date": row.get("investment_date", ""),
            "year": int(row.get("year", 0)),
            "status": row.get("status", "Ongoing"),
        })

    print(f"\nValidation results:")
    print(f"  Valid: {len(valid)}")
    print(f"  Skipped (existing): {skipped}")
    print(f"  Errors: {len(errors)}")

    if errors:
        print("\nFirst 10 errors:")
        for e in errors[:10]:
            print(f"  Row {e['row']} [{e['ref']}]: {e['error']}")

    total_amount = sum(v["principal_amount"] for v in valid)
    print(f"\nTotal investment amount: K {total_amount:,.2f}")

    if not args.execute:
        print("\n--- DRY RUN ---")
        print("First 15 investments:")
        for v in valid[:15]:
            print(
                f"  {v['reference_number']}: {v['investor_name']} | "
                f"K {v['principal_amount']:,.2f} | {v['year']} | {v['status']}"
            )
        if len(valid) > 15:
            print(f"  ... and {len(valid) - 15} more")
        print("\nRun with --execute to actually upload.")
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    ref_to_id = {}
    successes = 0
    upload_errors = []

    for i, v in enumerate(valid):
        payload = {
            "reference_number": v["reference_number"],
            "investor": v["investor_id"],
            "principal_amount": v["principal_amount"],
            "investment_date": v["investment_date"],
            "year": v["year"],
            "status": v["status"],
            "company": COMPANY_ID,
        }

        try:
            res = requests.post(
                f"{DIRECTUS_URL}/items/jutem_investments",
                headers=headers,
                json=payload,
            )
            res.raise_for_status()
            created = res.json()["data"]
            ref_to_id[v["reference_number"]] = created["id"]
            successes += 1
            print(
                f"  [{i+1}/{len(valid)}] Created: "
                f"{v['reference_number']} - {v['investor_name']} K {v['principal_amount']:,.2f}"
            )
        except requests.RequestException as e:
            err_msg = str(e)
            try:
                err_msg = res.json().get("errors", [{}])[0].get("message", err_msg)
            except Exception:
                pass
            upload_errors.append({
                "ref": v["reference_number"],
                "investor": v["investor_name"],
                "error": err_msg,
            })
            print(f"  [{i+1}/{len(valid)}] FAILED: {v['reference_number']} - {err_msg}")

        if (i + 1) % 10 == 0:
            time.sleep(0.5)

    # Write reference-to-ID mapping
    with open(OUTPUT_MAPPING, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["reference_number", "investor_name", "year", "id"])
        writer.writeheader()
        for v in valid:
            if v["reference_number"] in ref_to_id:
                writer.writerow({
                    "reference_number": v["reference_number"],
                    "investor_name": v["investor_name"],
                    "year": v["year"],
                    "id": ref_to_id[v["reference_number"]],
                })

    print(f"\n{'='*50}")
    print("Upload complete!")
    print(f"  Created: {successes}")
    print(f"  Skipped: {skipped}")
    print(f"  Failed: {len(upload_errors)}")
    print(f"  Mapping saved to: {OUTPUT_MAPPING}")

    if upload_errors:
        err_path = os.path.join(os.path.dirname(__file__), "output", "investments_upload_errors.csv")
        with open(err_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["ref", "investor", "error"])
            writer.writeheader()
            writer.writerows(upload_errors)
        print(f"  Error details saved to: {err_path}")


if __name__ == "__main__":
    main()
