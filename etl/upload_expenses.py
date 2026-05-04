#!/usr/bin/env python3
"""
Jutem Expenses Upload Script
==============================
Reads expenses_upload.csv and creates expense transactions in Directus.

For each expense:
- Creates a transaction in the `transactions` collection
- Deducts the expense amount from the account balance

Usage:
    python3 etl/upload_expenses.py                # Dry run
    python3 etl/upload_expenses.py --execute      # Actually upload
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
ACCOUNT_ID = 25  # Account to deduct from

CSV_PATH = os.path.join(
    os.path.dirname(__file__), "output", "expenses_upload.csv"
)


def get_token():
    """Get token from environment or prompt."""
    token = os.environ.get("DIRECTUS_TOKEN")
    if not token:
        token = input("Enter Directus admin token: ").strip()
    return token


def load_csv():
    """Load expenses CSV."""
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found.")
        print("Run jutem_etl.py --step expenses first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} expenses from CSV")
    return rows


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


def check_existing_expenses(token, reference_numbers):
    """Check which expense reference numbers already exist."""
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


def main():
    parser = argparse.ArgumentParser(
        description="Upload Jutem expenses to Directus"
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Actually upload (default is dry run)",
    )
    args = parser.parse_args()

    rows = load_csv()
    token = get_token()

    # Check for existing expenses
    all_refs = {row.get("reference_number", "") for row in rows}
    existing_refs = check_existing_expenses(token, all_refs) if args.execute else set()

    # Validate all rows
    valid = []
    errors = []
    skipped = 0

    for i, row in enumerate(rows):
        ref_num = row.get("reference_number", "")
        amount = float(row.get("amount", 0))
        description = row.get("expense_description", "").strip()

        if ref_num in existing_refs:
            skipped += 1
            continue

        if amount <= 0:
            errors.append({
                "row": i + 2,
                "ref": ref_num,
                "error": f"Invalid amount: {amount}",
            })
            continue

        if not description:
            errors.append({
                "row": i + 2,
                "ref": ref_num,
                "error": "Empty expense description",
            })
            continue

        valid.append({
            "row": i + 2,
            "amount": amount,
            "description": description,
            "transaction_date": row.get("transaction_date", ""),
            "reference_number": ref_num,
            "sheet": row.get("sheet", ""),
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

    total_amount = sum(v["amount"] for v in valid)
    print(f"\nTotal expense amount: K {total_amount:,.2f}")

    if not args.execute:
        print("\n--- DRY RUN ---")
        print("First 10 expenses:")
        for v in valid[:10]:
            print(
                f"  {v['reference_number']}: {v['description']} | "
                f"K {v['amount']:,.2f} | {v['transaction_date']} | {v['sheet']}"
            )
        if len(valid) > 10:
            print(f"  ... and {len(valid) - 10} more")
        print("\nRun with --execute to actually upload.")
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
    running_account_balance = account_balance

    for i, v in enumerate(valid):
        amount = v["amount"]

        # Deduct expense from account balance
        running_account_balance -= amount
        new_balance = round(running_account_balance, 2)

        payload = {
            "amount": amount,
            "payment_amount": amount,
            "transfer_fees": 0,
            "transaction_date": v["transaction_date"],
            "payment_method": "cash",
            "transaction_type": "Operational Cost",
            "reference_number": v["reference_number"],
            "notes": v["description"],
            "transaction_status": "Completed",
            "is_loan_transaction": False,
            "is_debt_repayment": False,
            "is_gps_fee_payment": False,
            "new_amount": new_balance,
            "company": COMPANY_ID,
            "account": ACCOUNT_ID,
        }

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
                f"{v['reference_number']} - {v['description']} K {amount:,.2f}"
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
                "ref": v["reference_number"],
                "description": v["description"],
                "error": err_msg,
            })
            print(
                f"  [{i+1}/{len(valid)}] FAILED: "
                f"{v['reference_number']} - {err_msg}"
            )

        if (i + 1) % 10 == 0:
            time.sleep(0.5)

    # Update account balance (final)
    final_balance = round(running_account_balance, 2)
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

    # Summary
    print(f"\n{'='*50}")
    print("Upload complete!")
    print(f"  Expenses created: {successes}")
    print(f"  Skipped (existing): {skipped}")
    print(f"  Failed: {len(upload_errors)}")
    print(f"  Account {ACCOUNT_ID} new balance: K {final_balance:,.2f}")

    if upload_errors:
        err_path = os.path.join(
            os.path.dirname(__file__),
            "output",
            "expenses_upload_errors.csv",
        )
        with open(err_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f, fieldnames=["ref", "description", "error"]
            )
            writer.writeheader()
            writer.writerows(upload_errors)
        print(f"  Error details saved to: {err_path}")


if __name__ == "__main__":
    main()
