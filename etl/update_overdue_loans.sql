-- ============================================================================
-- UPDATE LOAN STATUS TO OVERDUE (8) WHERE PAYMENTS ARE PAST DUE
-- ============================================================================
-- This script updates loans to "Overdue" status (loan_status = 8) where:
-- - There are amortization entries with status 'pending' or 'unsettled'
-- - The due_date of those entries is in the past (before today)
--
-- Example:
--   Loan: Has 1+ pending payments where due_date < CURRENT_DATE
--   Result: loan_status = 8 (Overdue)
--
-- Run this on your Directus database (PostgreSQL or MySQL)
-- ============================================================================

-- Step 1: View loans that will be updated (DRY RUN)
SELECT
    l.id AS loan_id,
    l.loan_reference,
    l.amount AS principal,
    l.loan_status AS current_status,
    COUNT(a.id) AS overdue_installments,
    MIN(a.due_date) AS earliest_overdue_date,
    MAX(a.due_date) AS latest_overdue_date,
    COALESCE(SUM(a.expected_amount), 0) AS total_overdue_amount,
    CASE
        WHEN COUNT(a.id) > 0 THEN 'WILL BE MARKED OVERDUE ⚠️'
        ELSE 'NO OVERDUE PAYMENTS'
    END AS payment_status
FROM loans l
INNER JOIN amortization a ON a.loan_reference = l.id
    AND (LOWER(a.status) = 'pending' OR LOWER(a.status) = 'unsettled')
    AND a.due_date < CURRENT_DATE
WHERE l.company = 22  -- Jutem Fund company ID
    AND l.loan_status != 8  -- Not already overdue
    AND l.loan_status != 10  -- Not settled
GROUP BY l.id, l.loan_reference, l.amount, l.loan_status
ORDER BY MIN(a.due_date) ASC;

-- ============================================================================
-- Step 2: ACTUAL UPDATE (uncomment to execute)
-- ============================================================================
-- WARNING: Test the SELECT query above first!
-- Make sure the results are correct before running the UPDATE
-- Backup your database before running this!

/*
UPDATE loans l
SET
    loan_status = 8,
    date_updated = NOW()
WHERE l.company = 22
    AND l.loan_status != 8
    AND l.loan_status != 10  -- Don't change settled loans
    AND l.id IN (
        SELECT DISTINCT loan_id FROM (
            SELECT
                l2.id AS loan_id
            FROM loans l2
            INNER JOIN amortization a ON a.loan_reference = l2.id
                AND (LOWER(a.status) = 'pending' OR LOWER(a.status) = 'unsettled')
                AND a.due_date < CURRENT_DATE
            WHERE l2.company = 22
                AND l2.loan_status != 8
                AND l2.loan_status != 10
        ) AS overdue_loans
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
    l.loan_status,
    COUNT(a.id) AS overdue_installments,
    MIN(a.due_date) AS earliest_overdue_date,
    COALESCE(SUM(a.expected_amount), 0) AS total_overdue_amount
FROM loans l
INNER JOIN amortization a ON a.loan_reference = l.id
    AND (LOWER(a.status) = 'pending' OR LOWER(a.status) = 'unsettled')
    AND a.due_date < CURRENT_DATE
WHERE l.company = 22
    AND l.loan_status = 8
GROUP BY l.id, l.loan_reference, l.amount, l.loan_status
ORDER BY l.date_updated DESC
LIMIT 20;
*/

-- ============================================================================
-- Step 4: REVERT/ROLLBACK (if you need to undo the update)
-- ============================================================================
-- If you need to revert the changes, you can change overdue loans back to active.

-- Option 1: Revert ALL overdue loans for company 22 back to Active (7)
/*
UPDATE loans
SET
    loan_status = 7,
    date_updated = NOW()
WHERE company = 22
    AND loan_status = 8;
*/

-- Option 2: Revert only loans that were recently updated (last 24 hours)
/*
UPDATE loans
SET
    loan_status = 7,
    date_updated = NOW()
WHERE company = 22
    AND loan_status = 8
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
    AND loan_status = 8
    AND id IN (123, 456, 789);  -- Replace with actual loan IDs
*/

-- ============================================================================
-- Additional: Update loans back to ACTIVE when overdue payments are cleared
-- ============================================================================
-- This query finds loans marked as overdue (8) but have NO overdue payments
-- and updates them back to active (7):
/*
UPDATE loans l
SET
    loan_status = 7,
    date_updated = NOW()
WHERE l.company = 22
    AND l.loan_status = 8
    AND NOT EXISTS (
        SELECT 1
        FROM amortization a
        WHERE a.loan_reference = l.id
            AND (LOWER(a.status) = 'pending' OR LOWER(a.status) = 'unsettled')
            AND a.due_date < CURRENT_DATE
    );
*/

-- ============================================================================
-- Additional: Check loans by days overdue
-- ============================================================================
-- This helps you see how many days overdue each loan is:
/*
SELECT
    l.id,
    l.loan_reference,
    l.amount AS principal,
    l.loan_status,
    COUNT(a.id) AS overdue_installments,
    MIN(a.due_date) AS earliest_overdue_date,
    CURRENT_DATE - MIN(a.due_date) AS days_overdue,
    COALESCE(SUM(a.expected_amount), 0) AS total_overdue_amount
FROM loans l
INNER JOIN amortization a ON a.loan_reference = l.id
    AND (LOWER(a.status) = 'pending' OR LOWER(a.status) = 'unsettled')
    AND a.due_date < CURRENT_DATE
WHERE l.company = 22
GROUP BY l.id, l.loan_reference, l.amount, l.loan_status
ORDER BY days_overdue DESC;
*/
