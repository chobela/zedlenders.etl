#!/usr/bin/env python3
"""
Jutem Amortization Upload Script
=================================
Reads amortization_upload.csv and creates amortization entries in Directus.

- Uses loan_ref_to_id.csv to map loan_reference strings to Directus loan IDs
- The amortization.loan_reference field is a relation to loans.id
- Uploads in batches for speed

Usage:
    python3 etl/upload_amortization.py                # Dry run
    python3 etl/upload_amortization.py --execute      # Actually upload
"""

import os
import sys
import csv
import argparse
import time
import requests

# ─── Configuration ───────────────────────────────────────────────────

DIRECTUS_URL = "https://zedlenders.pickmesms.com"

CSV_PATH = os.path.join(
    os.path.dirname(__file__), "output", "amortization_upload.csv"
)
MAPPING_PATH = os.path.join(
    os.path.dirname(__file__), "output", "loan_ref_to_id.csv"
)
BATCH_SIZE = 20


def get_token():
    """Get token from environment or prompt."""
    token = os.environ.get("DIRECTUS_TOKEN")
    if not token:
        token = input("Enter Directus admin token: ").strip()
    return token


def load_mapping():
    """Load loan_reference → loan_id mapping."""
    if not os.path.exists(MAPPING_PATH):
        print(f"ERROR: {MAPPING_PATH} not found.")
        print("Run upload_loans.py --execute first.")
        sys.exit(1)

    mapping = {}
    with open(MAPPING_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            mapping[row["loan_reference"]] = int(row["loan_id"])

    print(f"Loaded {len(mapping)} loan ref → ID mappings")
    return mapping


def load_csv():
    """Load amortization CSV."""
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found.")
        print("Run jutem_etl.py --step loans first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} amortization rows from CSV")
    return rows


def check_existing_count(token):
    """Check how many amortization entries already exist."""
    headers = {"Authorization": f"Bearer {token}"}
    try:
        res = requests.get(
            f"{DIRECTUS_URL}/items/amortization",
            headers=headers,
            params={"aggregate[count]": "id"},
        )
        res.raise_for_status()
        data = res.json().get("data", [])
        if data:
            return int(data[0].get("count", {}).get("id", 0))
    except Exception:
        pass
    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Upload Jutem amortization to Directus"
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Actually upload (default is dry run)",
    )
    args = parser.parse_args()

    ref_to_id = load_mapping()
    rows = load_csv()
    token = get_token()

    existing_count = check_existing_count(token)
    print(f"Existing amortization entries: {existing_count}")

    # Validate and build payloads
    valid = []
    errors = []

    for i, row in enumerate(rows):
        ref = row.get("loan_reference", "")
        loan_id = ref_to_id.get(ref)

        if not loan_id:
            errors.append({
                "row": i + 2,
                "ref": ref,
                "error": f"No loan ID mapping for '{ref}'",
            })
            continue

        payload = {
            "loan_reference": loan_id,
            "borrower_nrc": row.get("borrower_nrc", ""),
            "month": row.get("month", ""),
            "due_date": row.get("due_date", ""),
            "amount_due": round(float(row.get("amount_due") or 0), 5),
            "interest_rate": round(float(row.get("interest_rate") or 0), 5),
            "expected_amount": round(float(row.get("expected_amount") or 0), 5),
            "profit": round(float(row.get("profit") or 0), 5),
            "status": row.get("status", "pending"),
            "sheet": row.get("sheet", ""),
        }

        valid.append({"row": i + 2, "ref": ref, "payload": payload})

    print(f"\nValidation results:")
    print(f"  Valid: {len(valid)}")
    print(f"  Errors: {len(errors)}")

    if errors:
        print("\nFirst 10 errors:")
        for e in errors[:10]:
            print(f"  Row {e['row']} [{e['ref']}]: {e['error']}")
        if len(errors) > 10:
            print(f"  ... and {len(errors) - 10} more")

    if not args.execute:
        print("\n--- DRY RUN ---")
        print("First 5 entries:")
        for v in valid[:5]:
            p = v["payload"]
            print(
                f"  {v['ref']} → loan_id {p['loan_reference']} | "
                f"{p['month']} | K {p['amount_due']:,.2f} | "
                f"{p['status']}"
            )
        if len(valid) > 5:
            print(f"  ... and {len(valid) - 5} more")
        print("\nRun with --execute to actually upload.")
        return

    # Upload in batches
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    successes = 0
    upload_errors = []
    total_batches = (len(valid) + BATCH_SIZE - 1) // BATCH_SIZE

    for batch_num in range(total_batches):
        start = batch_num * BATCH_SIZE
        end = min(start + BATCH_SIZE, len(valid))
        batch = [v["payload"] for v in valid[start:end]]

        try:
            res = requests.post(
                f"{DIRECTUS_URL}/items/amortization",
                headers=headers,
                json=batch,
            )
            res.raise_for_status()
            successes += len(batch)
            print(
                f"  Batch [{batch_num+1}/{total_batches}]: "
                f"uploaded {len(batch)} entries "
                f"({successes}/{len(valid)} total)"
            )
        except requests.RequestException as e:
            err_msg = str(e)
            try:
                err_msg = (
                    res.json()
                    .get("errors", [{}])[0]
                    .get("message", err_msg)
                )
            except Exception:
                pass
            for v in valid[start:end]:
                upload_errors.append({
                    "ref": v["ref"],
                    "row": v["row"],
                    "error": err_msg,
                })
            print(
                f"  Batch [{batch_num+1}/{total_batches}]: "
                f"FAILED - {err_msg}"
            )

        if (batch_num + 1) % 5 == 0:
            time.sleep(0.5)

    # Summary
    print(f"\n{'='*50}")
    print("Upload complete!")
    print(f"  Created: {successes}")
    print(f"  Failed: {len(upload_errors)}")

    if upload_errors:
        err_path = os.path.join(
            os.path.dirname(__file__),
            "output",
            "amortization_upload_errors.csv",
        )
        with open(err_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f, fieldnames=["ref", "row", "error"]
            )
            writer.writeheader()
            writer.writerows(upload_errors)
        print(f"  Error details saved to: {err_path}")


if __name__ == "__main__":
    main()
