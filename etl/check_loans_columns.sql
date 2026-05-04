-- Check the structure of the loans table
SELECT column_name, data_type
FROM information_schema.columns
WHERE table_name = 'loans'
ORDER BY ordinal_position;
