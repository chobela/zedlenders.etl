-- ============================================================================
-- UPDATE LOAN STATUS TO SETTLED (10) WHERE AMORTIZATION FULLY PAID
-- ============================================================================
-- This script updates loans to "Settled" status (loan_status = 10) where:
-- - Total of paid amortization entries >= Loan's expected_return_amount
-- - Amortization status is 'paid' or 'cleared'
--
-- Example:
--   Loan: Principal=1000, Interest=300, Expected Total=1300
--   Amortizations: 1300 total where status='paid'
--   Result: loan_status = 10 (Settled)
--
-- Run this on your Directus database (PostgreSQL or MySQL)
-- ============================================================================

-- Step 1: View loans that will be updated (DRY RUN)
SELECT
    l.id AS loan_id,
    l.loan_reference,
    l.amount AS principal,
    COALESCE(l.interest_applied, l.custom_interest_rate, 0) AS interest_rate,
    l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0) AS total_expected,
    l.loan_status AS current_status,
    COALESCE(SUM(a.expected_amount), 0) AS total_paid_amortization,
    COUNT(a.id) AS paid_installments,
    CASE
        WHEN COALESCE(SUM(a.expected_amount), 0) >= (l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0)) THEN 'WILL BE SETTLED ✓'
        ELSE 'NOT FULLY PAID'
    END AS payment_status
FROM loans l
LEFT JOIN amortization a ON a.loan_reference = l.id
    AND (LOWER(a.status) = 'paid' OR LOWER(a.status) = 'cleared')
WHERE l.company = 22  -- Jutem Fund company ID
    AND l.loan_status != 10  -- Not already settled
GROUP BY l.id, l.loan_reference, l.amount, l.interest_applied, l.custom_interest_rate, l.loan_status
HAVING COALESCE(SUM(a.expected_amount), 0) >= (l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0))
ORDER BY l.id;

-- ============================================================================
-- Step 2: ACTUAL UPDATE (uncomment to execute)
-- ============================================================================
-- WARNING: Test the SELECT query above first!
-- Make sure the results are correct before running the UPDATE
-- Backup your database before running this!

/*
UPDATE loans l
SET
    loan_status = 10,
    date_updated = NOW()
WHERE l.company = 22
    AND l.loan_status != 10
    AND l.id IN (
        SELECT loan_id FROM (
            SELECT
                l2.id AS loan_id,
                COALESCE(SUM(a.expected_amount), 0) AS total_paid,
                l2.amount * (1 + COALESCE(l2.interest_applied, l2.custom_interest_rate, 0) / 100.0) AS expected_total
            FROM loans l2
            LEFT JOIN amortization a ON a.loan_reference = l2.id
                AND (LOWER(a.status) = 'paid' OR LOWER(a.status) = 'cleared')
            WHERE l2.company = 22
                AND l2.loan_status != 10
            GROUP BY l2.id, l2.amount, l2.interest_applied, l2.custom_interest_rate
            HAVING COALESCE(SUM(a.expected_amount), 0) >= (l2.amount * (1 + COALESCE(l2.interest_applied, l2.custom_interest_rate, 0) / 100.0))
        ) AS fully_paid_loans
    );
*/

-- ============================================================================
-- Step 3: Verify the update
-- ============================================================================
-- Run this after the UPDATE to verify the loans were updated correctly:
/*
SELECT
    l.id,
    l.loan_reference,
    l.amount AS principal,
    COALESCE(l.interest_applied, l.custom_interest_rate, 0) AS interest_rate,
    l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0) AS total_expected,
    l.loan_status,
    COALESCE(SUM(a.expected_amount), 0) AS total_paid,
    COUNT(a.id) AS paid_installments
FROM loans l
LEFT JOIN amortization a ON a.loan_reference = l.id
    AND (LOWER(a.status) = 'paid' OR LOWER(a.status) = 'cleared')
WHERE l.company = 22
    AND l.loan_status = 10
GROUP BY l.id, l.loan_reference, l.amount, l.interest_applied, l.custom_interest_rate, l.loan_status
ORDER BY l.date_updated DESC
LIMIT 20;
*/

-- ============================================================================
-- Step 4: REVERT/ROLLBACK (if you need to undo the update)
-- ============================================================================
-- If you need to revert the changes, you can change settled loans back to active.
-- This will revert loans that were updated to status 10 (Settled) back to status 7 (Active).

-- Option 1: Revert ALL settled loans for company 22 back to Active (7)
/*
UPDATE loans
SET
    loan_status = 7,
    date_updated = NOW()
WHERE company = 22
    AND loan_status = 10;
*/

-- Option 2: Revert only loans that were recently updated (last 24 hours)
/*
UPDATE loans
SET
    loan_status = 7,
    date_updated = NOW()
WHERE company = 22
    AND loan_status = 10
    AND date_updated >= NOW() - INTERVAL '24 hours';  -- PostgreSQL
    -- AND date_updated >= NOW() - INTERVAL 24 HOUR;  -- MySQL (use this line instead for MySQL)
*/

-- Option 3: Revert specific loans by ID
/*
UPDATE loans
SET
    loan_status = 7,
    date_updated = NOW()
WHERE company = 22
    AND loan_status = 10
    AND id IN (123, 456, 789);  -- Replace with actual loan IDs
*/

-- Option 4: Revert to Overdue (8) instead of Active (7) if they have overdue payments
/*
UPDATE loans l
SET
    loan_status = 8,
    date_updated = NOW()
WHERE l.company = 22
    AND l.loan_status = 10
    AND l.next_payment_date < CURRENT_DATE;
*/

-- ============================================================================
-- Additional: Check loans that are ALMOST paid (90%+)
-- ============================================================================
-- This helps you see which loans are close to being settled:
/*
SELECT
    l.id,
    l.loan_reference,
    l.amount AS principal,
    l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0) AS total_expected,
    COALESCE(SUM(a.expected_amount), 0) AS total_paid,
    ROUND((COALESCE(SUM(a.expected_amount), 0) / (l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0)) * 100), 2) AS percent_paid,
    (l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0)) - COALESCE(SUM(a.expected_amount), 0) AS remaining
FROM loans l
LEFT JOIN amortization a ON a.loan_reference = l.id
    AND (LOWER(a.status) = 'paid' OR LOWER(a.status) = 'cleared')
WHERE l.company = 22
    AND l.loan_status != 10
GROUP BY l.id, l.loan_reference, l.amount, l.interest_applied, l.custom_interest_rate
HAVING COALESCE(SUM(a.expected_amount), 0) >= ((l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0)) * 0.9)
    AND COALESCE(SUM(a.expected_amount), 0) < (l.amount * (1 + COALESCE(l.interest_applied, l.custom_interest_rate, 0) / 100.0))
ORDER BY percent_paid DESC;
*/
