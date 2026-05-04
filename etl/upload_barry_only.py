#!/usr/bin/env python3
"""Upload only Barry's loan JUTEM-00172"""

import os
import sys

# Change CSV path in upload_loans module
import upload_loans
upload_loans.CSV_PATH = os.path.join(
    os.path.dirname(__file__), "output", "barry_loan.csv"
)

if __name__ == "__main__":
    upload_loans.main()
