#!/usr/bin/env python3
"""
Jutem Interest Expenses Upload Script
=======================================
Reads interest_expenses_upload.csv and creates interest expense
transactions in Directus, linked to jutem_investments.

Usage:
    python3 etl/upload_interest_expenses.py                # Dry run
    python3 etl/upload_interest_expenses.py --execute      # Actually upload
"""

import os
import sys
import csv
import argparse
import time
import requests
from datetime import datetime

# ─── Configuration ───────────────────────────────────────────────────
DIRECTUS_URL = "https://zedlenders.pickmesms.com"
COMPANY_ID = 22
ACCOUNT_ID = 25

CSV_PATH = os.path.join(os.path.dirname(__file__), "output", "interest_expenses_upload.csv")
INVESTMENT_MAPPING = os.path.join(os.path.dirname(__file__), "output", "investment_ref_to_id.csv")

# ─── Name aliases ────────────────────────────────────────────────────
# Maps (UPPER_NAME, year) from expenses CSV → (UPPER_NAME, year) in investment mapping.
# Entries without a year key apply to all years.
# Format: expense_name_upper → investment_name_upper
NAME_ALIASES = {
    "KHAYA": "KAYA",
    "AUNTY JANE": "AUNT JANE",
    "THANDI SIYOS": "THANDI SIYO",
    "MR NKATANI": "MR NKETANI",
    "MR. NKETANI": "MR NKETANI",
    "G.B": "GB",
    "RABECCA ALIFORD PHIRI": "RABECCA A. PHIRI",
    "CHILESHE": "CHILESHE SIMPITU",
}

# Year-specific aliases: (expense_name_upper, year) → investment_name_upper
# These override the generic aliases when the year matches.
YEAR_ALIASES = {
    # "Thandi" without surname maps to "Thandi" investment (2024) or "Thandi Sis" (2023)
    ("THANDI", 2023): "THANDI SIS",
    ("THANDI", 2024): "THANDI",
    # "Thandi sis" in 2024 expense maps to the 2023 investment (carried over)
    ("THANDI SIS", 2024): "THANDI SIS",
    # "Misozi" maps to "Madam Misozi" for years where only that entry exists
    ("MISOZI", 2023): "MADAM MISOZI",
    ("MISOZI", 2024): "MADAM MISOZI",
    ("MISOZI", 2025): "MADAM MISOZI",
    ("MISOZI", 2026): "MADAM MISOZI",
    # "Nketani" without "Mr" in 2024 → "Mr Nketani"
    ("NKETANI", 2024): "MR NKETANI",
    # "sarafina" (lowercase) in 2024 → "Sarafina"
    ("SARAFINA", 2024): "SARAFINA",
    # "Madam Misozi" in 2024 expense → reuse the same
    ("MADAM MISOZI", 2024): "MADAM MISOZI",
}


def normalize_investor_name(name_upper, year):
    """Resolve expense investor name to the canonical investment name."""
    # Check year-specific alias first
    key = (name_upper, year)
    if key in YEAR_ALIASES:
        return YEAR_ALIASES[key]
    # Then generic alias
    if name_upper in NAME_ALIASES:
        return NAME_ALIASES[name_upper]
    return name_upper


def get_token():
    token = os.environ.get("DIRECTUS_TOKEN")
    if not token:
        token = input("Enter Directus admin token: ").strip()
    return token


def load_csv():
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found.")
        print("Run jutem_etl.py --step interest_expenses first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} interest expense rows from CSV")
    return rows


def load_investment_mapping():
    """Load investor_name+year → investment_id mapping."""
    if not os.path.exists(INVESTMENT_MAPPING):
        print(f"WARNING: {INVESTMENT_MAPPING} not found.")
        print("Run upload_investments.py --execute first for investment linking.")
        return {}

    mapping = {}
    with open(INVESTMENT_MAPPING, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            key = (row["investor_name"].upper().strip(), int(row["year"]))
            mapping[key] = int(row["id"])

    print(f"Loaded {len(mapping)} investment mappings")
    return mapping


def get_year_from_sheet(sheet_name):
    """Extract year from sheet name like 'MARCH 2026'."""
    try:
        dt = datetime.strptime(sheet_name.strip(), "%B %Y")
        return dt.year
    except ValueError:
        return None


def check_existing(token, reference_numbers):
    """Check which references already exist."""
    headers = {"Authorization": f"Bearer {token}"}
    existing = set()

    ref_list = list(reference_numbers)
    batch_size = 50
    for i in range(0, len(ref_list), batch_size):
        batch = ref_list[i : i + batch_size]
        refs_str = ",".join(batch)
        res = requests.get(
            f"{DIRECTUS_URL}/items/transactions",
            headers=headers,
            params={
                "filter[reference_number][_in]": refs_str,
                "filter[company][_eq]": COMPANY_ID,
                "fields": "id,reference_number",
                "limit": -1,
            },
        )
        res.raise_for_status()
        for txn in res.json().get("data", []):
            ref = txn.get("reference_number")
            if ref:
                existing.add(ref)

    return existing


def fetch_account_balance(token):
    """Fetch current balance of the target account."""
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(
        f"{DIRECTUS_URL}/items/accounts",
        headers=headers,
        params={
            "filter[id][_eq]": ACCOUNT_ID,
            "fields": "id,balance",
            "limit": 1,
        },
    )
    res.raise_for_status()
    data = res.json().get("data", [])
    if not data:
        print(f"ERROR: Account {ACCOUNT_ID} not found.")
        sys.exit(1)
    account = data[0]
    print(f"Account {ACCOUNT_ID} balance: K {float(account['balance']):,.2f}")
    return account


def delete_existing_expense_transactions(token):
    """Delete all interest_expense transactions for COMPANY_ID."""
    headers = {"Authorization": f"Bearer {token}"}
    print(f"\nDeleting existing interest_expense transactions for company {COMPANY_ID}...")

    res = requests.get(
        f"{DIRECTUS_URL}/items/transactions",
        headers=headers,
        params={
            "filter[company][_eq]": COMPANY_ID,
            "filter[transaction_type][_eq]": "interest_expense",
            "fields": "id",
            "limit": -1,
        },
    )
    res.raise_for_status()
    items = res.json().get("data", [])

    # Also check for old "Interest Expense" casing
    res2 = requests.get(
        f"{DIRECTUS_URL}/items/transactions",
        headers=headers,
        params={
            "filter[company][_eq]": COMPANY_ID,
            "filter[transaction_type][_eq]": "Interest Expense",
            "fields": "id",
            "limit": -1,
        },
    )
    res2.raise_for_status()
    items.extend(res2.json().get("data", []))

    if not items:
        print("  No existing interest expense transactions found.")
        return

    ids = list({item["id"] for item in items})
    print(f"  Found {len(ids)} transactions to delete.")

    batch_size = 50
    for i in range(0, len(ids), batch_size):
        batch = ids[i : i + batch_size]
        for item_id in batch:
            res = requests.delete(
                f"{DIRECTUS_URL}/items/transactions/{item_id}",
                headers=headers,
            )
            res.raise_for_status()
        print(f"  Deleted {min(i + batch_size, len(ids))}/{len(ids)}")
        time.sleep(0.5)

    print(f"  Successfully deleted {len(ids)} transactions.")


def main():
    parser = argparse.ArgumentParser(description="Upload Jutem interest expenses to Directus")
    parser.add_argument("--execute", action="store_true", help="Actually upload (default is dry run)")
    parser.add_argument("--delete-existing", action="store_true",
                        help="Delete all existing interest expense transactions before uploading")
    args = parser.parse_args()

    rows = load_csv()
    token = get_token()

    # Delete existing transactions if requested
    if args.execute and args.delete_existing:
        delete_existing_expense_transactions(token)

    # Always load mapping (for dry run validation too)
    inv_mapping = load_investment_mapping()

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

        interest_exp = float(row.get("interest_expense", 0))
        if interest_exp <= 0:
            errors.append({"row": i + 2, "ref": ref, "error": f"Invalid interest expense: {interest_exp}"})
            continue

        # Look up investment ID by investor_name + year from sheet
        sheet_year = get_year_from_sheet(row.get("sheet", ""))
        if sheet_year:
            raw_name = row["investor_name"].upper().strip()
            canonical_name = normalize_investor_name(raw_name, sheet_year)
            inv_key = (canonical_name, sheet_year)
        else:
            inv_key = None
        investment_id = inv_mapping.get(inv_key) if inv_key else None

        valid.append({
            "row": i + 2,
            "investor_name": row["investor_name"],
            "date_expected": row.get("date_expected", ""),
            "borrowed_amount": float(row.get("borrowed_amount", 0)),
            "interest_rate": float(row.get("interest_rate", 0)),
            "paid_amount": float(row.get("paid_amount", 0)),
            "interest_expense": interest_exp,
            "reference_number": ref,
            "sheet": row.get("sheet", ""),
            "investment_id": investment_id,
        })

    print(f"\nValidation results:")
    print(f"  Valid: {len(valid)}")
    print(f"  Skipped (existing): {skipped}")
    print(f"  Errors: {len(errors)}")

    linked = [v for v in valid if v["investment_id"]]
    unlinked = [v for v in valid if not v["investment_id"]]

    total_interest = sum(v["interest_expense"] for v in valid)
    total_paid = sum(v["paid_amount"] for v in valid)
    print(f"\nTotal interest expense: K {total_interest:,.2f}")
    print(f"Total paid (principal + interest): K {total_paid:,.2f}")
    print(f"Linked to investment: {len(linked)}/{len(valid)}")

    if unlinked:
        print(f"\nUnlinked expenses ({len(unlinked)}):")
        seen_names = set()
        for v in unlinked:
            name_year = (v["investor_name"], v["sheet"])
            if name_year not in seen_names:
                seen_names.add(name_year)
                print(f"  {v['investor_name']} | {v['sheet']}")

    if not args.execute:
        print("\n--- DRY RUN ---")
        print("First 15 interest expenses:")
        for v in valid[:15]:
            link_status = "LINKED" if v["investment_id"] else "UNLINKED"
            print(
                f"  [{link_status}] {v['reference_number']}: {v['investor_name']} | "
                f"Borrowed K {v['borrowed_amount']:,.2f} @ {v['interest_rate']*100:.1f}% | "
                f"Interest K {v['interest_expense']:,.2f} | {v['date_expected']} | {v['sheet']}"
            )
        if len(valid) > 15:
            print(f"  ... and {len(valid) - 15} more")
        print("\nRun with --execute --delete-existing to actually upload.")
        return

    # Fetch current account balance
    account = fetch_account_balance(token)
    account_balance = float(account["balance"])

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    successes = 0
    upload_errors = []
    running_balance = account_balance

    for i, v in enumerate(valid):
        interest_amount = v["interest_expense"]  # The actual cost of borrowing
        paid_out = v["paid_amount"]  # Total paid out (principal + interest)

        running_balance -= paid_out
        new_balance = round(running_balance, 2)

        payload = {
            "amount": interest_amount,
            "transfer_fees": 0,
            "transaction_date": v["date_expected"],
            "payment_method": "cash",
            "transaction_type": "interest_expense",
            "reference_number": v["reference_number"],
            "notes": (
                f"Interest payment to {v['investor_name']} | "
                f"Borrowed: K{v['borrowed_amount']:,.2f} | "
                f"Rate: {v['interest_rate']*100:.1f}% | "
                f"Interest: K{v['interest_expense']:,.2f}"
            ),
            "transaction_status": "Completed",
            "is_loan_transaction": False,
            "is_debt_repayment": True,
            "is_gps_fee_payment": False,
            "new_amount": new_balance,
            "company": COMPANY_ID,
            "account": ACCOUNT_ID,
        }

        if v["investment_id"]:
            payload["investment_reference"] = v["investment_id"]

        try:
            res = requests.post(
                f"{DIRECTUS_URL}/items/transactions",
                headers=headers,
                json=payload,
            )
            res.raise_for_status()
            successes += 1
            print(
                f"  [{i+1}/{len(valid)}] Created: "
                f"{v['reference_number']} - {v['investor_name']} K {v['interest_expense']:,.2f}"
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

    # Update account balance
    final_balance = round(running_balance, 2)
    print(f"\nUpdating account {ACCOUNT_ID} balance: K {account_balance:,.2f} → K {final_balance:,.2f}")
    try:
        res = requests.patch(
            f"{DIRECTUS_URL}/items/accounts/{ACCOUNT_ID}",
            headers=headers,
            json={"balance": final_balance},
        )
        res.raise_for_status()
        print("  Account balance updated successfully")
    except requests.RequestException as e:
        print(f"  FAILED to update account balance: {e}")

    print(f"\n{'='*50}")
    print("Upload complete!")
    print(f"  Created: {successes}")
    print(f"  Skipped: {skipped}")
    print(f"  Failed: {len(upload_errors)}")
    print(f"  Account {ACCOUNT_ID} new balance: K {final_balance:,.2f}")

    if upload_errors:
        err_path = os.path.join(os.path.dirname(__file__), "output", "interest_expenses_upload_errors.csv")
        with open(err_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["ref", "investor", "error"])
            writer.writeheader()
            writer.writerows(upload_errors)
        print(f"  Error details saved to: {err_path}")


if __name__ == "__main__":
    main()
