#!/usr/bin/env python3
"""
Jutem Payments Upload Script
==============================
Reads payments_upload.csv and creates payment transactions in Directus.

For each payment:
- Creates a transaction in the `transactions` collection
- Deducts the payment amount from the account with id=22
- Updates the loan's paid_amount (deducts from principal)

Uses loan_ref_to_id.csv to map loan_reference strings to Directus loan IDs.

Usage:
    python3 etl/upload_payments.py                # Dry run
    python3 etl/upload_payments.py --execute      # Actually upload
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
    os.path.dirname(__file__), "output", "payments_upload.csv"
)
MAPPING_PATH = os.path.join(
    os.path.dirname(__file__), "output", "loan_ref_to_id.csv"
)


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
    """Load payments CSV."""
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: {CSV_PATH} not found.")
        print("Run jutem_etl.py --step payments first.")
        sys.exit(1)

    rows = []
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} payments from CSV")
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
    print(f"Account {ACCOUNT_ID} balance: K {account['balance']:,.2f}")
    return account


def fetch_loans_paid_amounts(token, loan_ids):
    """Fetch current paid_amount for all relevant loans."""
    headers = {"Authorization": f"Bearer {token}"}
    loans = {}

    # Fetch in batches
    loan_id_list = list(set(loan_ids))
    batch_size = 50
    for i in range(0, len(loan_id_list), batch_size):
        batch = loan_id_list[i : i + batch_size]
        ids_str = ",".join(str(lid) for lid in batch)
        res = requests.get(
            f"{DIRECTUS_URL}/items/loans",
            headers=headers,
            params={
                "filter[id][_in]": ids_str,
                "fields": "id,paid_amount,amount,loan_reference,loan_status",
                "limit": -1,
            },
        )
        res.raise_for_status()
        for loan in res.json().get("data", []):
            loans[loan["id"]] = loan

    print(f"Fetched {len(loans)} loans for paid_amount updates")
    return loans


def check_existing_payments(token, reference_numbers):
    """Check which payment reference numbers already exist."""
    headers = {"Authorization": f"Bearer {token}"}
    existing = set()

    # Check in batches
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
        description="Upload Jutem payments to Directus"
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

    existing_refs = set()

    # Validate all rows
    valid = []
    errors = []
    skipped = 0

    for i, row in enumerate(rows):
        loan_ref = row.get("loan_reference", "")
        ref_num = row.get("reference_number", "")
        amount = float(row.get("payment_amount", 0))

        if ref_num in existing_refs:
            skipped += 1
            continue

        loan_id = ref_to_id.get(loan_ref)
        if not loan_id:
            errors.append({
                "row": i + 2,
                "ref": ref_num,
                "error": f"No loan ID mapping for '{loan_ref}'",
            })
            continue

        if amount <= 0:
            errors.append({
                "row": i + 2,
                "ref": ref_num,
                "error": f"Invalid payment amount: {amount}",
            })
            continue

        valid.append({
            "row": i + 2,
            "loan_ref": loan_ref,
            "loan_id": loan_id,
            "amount": amount,
            "payment_date": row.get("payment_date", ""),
            "payment_method": row.get("payment_method", "cash"),
            "transaction_type": row.get("transaction_type", "Loan Repayment"),
            "reference_number": ref_num,
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

    # Calculate totals
    total_amount = sum(v["amount"] for v in valid)
    print(f"\nTotal payment amount: K {total_amount:,.2f}")

    if not args.execute:
        print("\n--- DRY RUN ---")
        print("First 10 payments:")
        for v in valid[:10]:
            print(
                f"  {v['reference_number']}: {v['loan_ref']} → loan_id {v['loan_id']} | "
                f"K {v['amount']:,.2f} | {v['payment_date']}"
            )
        if len(valid) > 10:
            print(f"  ... and {len(valid) - 10} more")
        print("\nRun with --execute to actually upload.")
        return

    # Fetch current account balance and loan data
    account = fetch_account_balance(token)
    account_balance = float(account["balance"])

    loan_ids = [v["loan_id"] for v in valid]
    loans = fetch_loans_paid_amounts(token, loan_ids)

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    successes = 0
    upload_errors = []

    # Track cumulative changes
    running_account_balance = account_balance
    loan_paid_updates = {}  # loan_id -> cumulative paid_amount

    for i, v in enumerate(valid):
        loan_id = v["loan_id"]
        amount = v["amount"]

        # Update running account balance
        running_account_balance += amount
        new_balance = round(running_account_balance, 2)

        # Calculate new loan paid_amount
        if loan_id not in loan_paid_updates:
            loan_data = loans.get(loan_id, {})
            loan_paid_updates[loan_id] = float(loan_data.get("paid_amount", 0) or 0)
        loan_paid_updates[loan_id] += amount

        # Create the transaction
        payload = {
            "loan": loan_id,
            "amount": amount,
            "payment_amount": amount,
            "transfer_fees": 0,
            "transaction_date": v["payment_date"],
            "payment_method": v["payment_method"],
            "transaction_type": v["transaction_type"],
            "reference_number": v["reference_number"],
            "transaction_status": "Completed",
            "is_loan_transaction": True,
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
                f"{v['reference_number']} - {v['loan_ref']} K {amount:,.2f}"
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
                "loan_ref": v["loan_ref"],
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

    # Update each loan's paid_amount
    print(f"\nUpdating paid_amount for {len(loan_paid_updates)} loans...")
    loan_update_errors = 0
    for loan_id, new_paid in loan_paid_updates.items():
        loan_data = loans.get(loan_id, {})
        loan_amount = float(loan_data.get("amount", 0) or 0)
        new_status = 10 if new_paid >= loan_amount else 7  # 10 = fully paid, 7 = partially paid

        try:
            res = requests.patch(
                f"{DIRECTUS_URL}/items/loans/{loan_id}",
                headers=headers,
                json={
                    "paid_amount": round(new_paid, 2),
                    "loan_status": new_status,
                },
            )
            res.raise_for_status()
        except requests.RequestException as e:
            loan_update_errors += 1
            print(f"  FAILED to update loan {loan_id}: {e}")

    print(f"  Loans updated: {len(loan_paid_updates) - loan_update_errors}/{len(loan_paid_updates)}")

    # Summary
    print(f"\n{'='*50}")
    print("Upload complete!")
    print(f"  Payments created: {successes}")
    print(f"  Skipped (existing): {skipped}")
    print(f"  Failed: {len(upload_errors)}")
    print(f"  Account {ACCOUNT_ID} new balance: K {final_balance:,.2f}")
    print(f"  Loans updated: {len(loan_paid_updates)}")

    if upload_errors:
        err_path = os.path.join(
            os.path.dirname(__file__),
            "output",
            "payments_upload_errors.csv",
        )
        with open(err_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f, fieldnames=["ref", "loan_ref", "error"]
            )
            writer.writeheader()
            writer.writerows(upload_errors)
        print(f"  Error details saved to: {err_path}")


if __name__ == "__main__":
    main()
