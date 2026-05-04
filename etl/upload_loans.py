#!/usr/bin/env python3
"""
Jutem Loans Upload Script
=========================
Reads loans_upload.csv and creates loans in Directus.

- Looks up borrower by NRC
- Looks up loan product by name
- Looks up branch by name
- Checks for duplicate loan_reference before creating
- Stores loan_reference for later amortization/payment linking

Usage:
    python3 etl/upload_loans.py                    # Dry run (default)
    python3 etl/upload_loans.py --execute          # Actually upload
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

CSV_PATH = os.path.join(
    os.path.dirname(__file__), "output", "loans_upload.csv"
)
MAPPING_OUTPUT = os.path.join(
    os.path.dirname(__file__), "output", "loan_ref_to_id.csv"
)


def get_token():
    """Get token from environment or prompt."""
    token = os.environ.get("DIRECTUS_TOKEN")
    if not token:
        token = input("Enter Directus admin token: ").strip()
    return token


def load_csv():
    """Load and parse the loans CSV."""
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found.")
        print("Run jutem_etl.py --step loans first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} loans from CSV")
    return rows


def fetch_lookup_data(token):
    """Fetch borrowers, loan products, and branches for lookups."""
    headers = {"Authorization": f"Bearer {token}"}

    print("Fetching borrowers...")
    borrowers = {}
    res = requests.get(
        f"{DIRECTUS_URL}/users",
        headers=headers,
        params={
            "filter[role][name][_eq]": "Borrower",
            "filter[company][_eq]": COMPANY_ID,
            "fields": "id,nrc,first_name,last_name",
            "limit": -1,
        },
    )
    res.raise_for_status()
    for b in res.json().get("data", []):
        if b.get("nrc"):
            borrowers[b["nrc"]] = b
    print(f"  {len(borrowers)} borrowers loaded")

    print("Fetching loan products...")
    products = {}
    res = requests.get(
        f"{DIRECTUS_URL}/items/loan_products",
        headers=headers,
        params={
            "filter[company][_eq]": COMPANY_ID,
            "filter[is_active][_eq]": "true",
            "fields": "*",
            "limit": -1,
        },
    )
    res.raise_for_status()
    for p in res.json().get("data", []):
        products[p["loan_name"].lower()] = p
    print(f"  {len(products)} loan products loaded")

    print("Fetching branches...")
    branches = {}
    res = requests.get(
        f"{DIRECTUS_URL}/items/branches",
        headers=headers,
        params={
            "filter[company][_eq]": COMPANY_ID,
            "fields": "*",
            "limit": -1,
        },
    )
    res.raise_for_status()
    for br in res.json().get("data", []):
        branches[br["branch_name"].lower()] = br
    print(f"  {len(branches)} branches loaded")

    return borrowers, products, branches


def check_existing_loans(token):
    """Check which loan references already exist."""
    headers = {"Authorization": f"Bearer {token}"}
    existing = set()

    res = requests.get(
        f"{DIRECTUS_URL}/items/loans",
        headers=headers,
        params={
            "filter[company][_eq]": COMPANY_ID,
            "fields": "id,loan_reference",
            "limit": -1,
        },
    )
    res.raise_for_status()
    for loan in res.json().get("data", []):
        ref = loan.get("loan_reference")
        if ref:
            existing.add(ref)

    return existing


def main():
    parser = argparse.ArgumentParser(
        description="Upload Jutem loans to Directus"
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Actually upload (default is dry run)",
    )
    args = parser.parse_args()

    rows = load_csv()
    token = get_token()

    borrowers, products, branches = fetch_lookup_data(token)

    # Check existing
    print("Checking for existing loans...")
    existing_refs = check_existing_loans(token)
    print(f"  {len(existing_refs)} loans already in system")

    # Validate all rows first
    valid = []
    skipped = 0
    errors = []

    for i, row in enumerate(rows):
        ref = row.get("loan_reference", "")
        nrc = row.get("borrower_nrc", "")
        product_name = row.get("loan_product_name", "").lower()
        branch_name = row.get("branch_name", "").lower()

        if ref in existing_refs:
            skipped += 1
            continue

        borrower = borrowers.get(nrc)
        if not borrower:
            errors.append({
                "row": i + 2,
                "ref": ref,
                "error": f"Borrower NRC '{nrc}' not found",
            })
            continue

        product = products.get(product_name)
        if not product:
            errors.append({
                "row": i + 2,
                "ref": ref,
                "error": f"Loan product '{product_name}' not found",
            })
            continue

        branch_id = None
        if branch_name:
            branch = branches.get(branch_name)
            if branch:
                branch_id = branch["id"]

        amount = float(row.get("amount", 0))
        interest = float(row.get("custom_interest_rate", 0))
        loan_status = int(row.get("loan_status", 1))

        interest_decimal = interest / 100.0
        expected_return = amount * (1 + interest_decimal)
        profit = expected_return - amount

        valid.append({
            "row": row,
            "ref": ref,
            "borrower": borrower,
            "product": product,
            "branch_id": branch_id,
            "amount": amount,
            "interest": interest,
            "loan_status": loan_status,
            "expected_return": round(expected_return, 2),
            "profit": round(profit, 2),
        })

    print(f"\nValidation results:")
    print(f"  Valid: {len(valid)}")
    print(f"  Skipped (existing): {skipped}")
    print(f"  Errors: {len(errors)}")

    if errors:
        print("\nFirst 10 errors:")
        for e in errors[:10]:
            print(f"  Row {e['row']} [{e['ref']}]: {e['error']}")
        if len(errors) > 10:
            print(f"  ... and {len(errors) - 10} more")

    if not args.execute:
        print("\n--- DRY RUN ---")
        print("First 5 valid loans:")
        for v in valid[:5]:
            b = v["borrower"]
            name = f"{b['first_name']} {b['last_name']}"
            print(
                f"  {v['ref']}: {name} - "
                f"K {v['amount']:,.2f} @ {v['interest']}%"
            )
        if len(valid) > 5:
            print(f"  ... and {len(valid) - 5} more")
        print("\nRun with --execute to actually upload.")
        return

    # Upload
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    successes = 0
    upload_errors = []
    ref_to_id = {}

    for i, v in enumerate(valid):
        row = v["row"]
        payload = {
            "borrower": v["borrower"]["id"],
            "loan_product": v["product"]["id"],
            "amount": v["amount"],
            "custom_interest_rate": v["interest"],
            "interest_applied": v["interest"],
            "application_date": row.get("application_date"),
            "loan_purpose": row.get("loan_purpose", "Business"),
            "branch": v["branch_id"],
            "company": COMPANY_ID,
            "loan_status": v["loan_status"],
            "status": "published",
            "expected_return_amount": v["expected_return"],
            "profit_amount": v["profit"],
            "loan_reference": v["ref"],
        }

        if row.get("approval_date"):
            payload["approval_date"] = row["approval_date"]

        b = v["borrower"]
        name = f"{b['first_name']} {b['last_name']}"

        try:
            res = requests.post(
                f"{DIRECTUS_URL}/items/loans",
                headers=headers,
                json=payload,
            )
            res.raise_for_status()
            loan_id = res.json()["data"]["id"]
            ref_to_id[v["ref"]] = loan_id
            successes += 1
            print(
                f"  [{i+1}/{len(valid)}] Created: "
                f"{v['ref']} - {name} K {v['amount']:,.2f}"
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
            upload_errors.append({
                "ref": v["ref"],
                "name": name,
                "error": err_msg,
            })
            print(
                f"  [{i+1}/{len(valid)}] FAILED: "
                f"{v['ref']} - {err_msg}"
            )

        if (i + 1) % 10 == 0:
            time.sleep(0.5)

    # Save ref → id mapping for amortization/payment uploads
    with open(MAPPING_OUTPUT, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["loan_reference", "loan_id"])
        for ref, lid in ref_to_id.items():
            writer.writerow([ref, lid])

    # Summary
    print(f"\n{'='*50}")
    print("Upload complete!")
    print(f"  Created: {successes}")
    print(f"  Skipped (existing): {skipped}")
    print(f"  Failed: {len(upload_errors)}")
    print(f"\nLoan ref → ID mapping saved to: {MAPPING_OUTPUT}")
    print("  (Needed for amortization and payment uploads)")

    if upload_errors:
        err_path = os.path.join(
            os.path.dirname(__file__),
            "output",
            "loans_upload_errors.csv",
        )
        with open(err_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f, fieldnames=["ref", "name", "error"]
            )
            writer.writeheader()
            writer.writerows(upload_errors)
        print(f"  Error details saved to: {err_path}")


if __name__ == "__main__":
    main()
