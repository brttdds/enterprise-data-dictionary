SELECT
    TABLE_SCHEMA        AS [Schema],
    TABLE_NAME          AS [Source Table],
    COLUMN_NAME         AS [Field],
    DATA_TYPE           AS [Data Type],
    ''                  AS [Definition],
    'Pending'           AS [Confirmed],
    ''                  AS [Notes],
    ''                  AS [Joins With]
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'YourSchemaName'
  AND TABLE_NAME IN (
      'TableOne',
      'TableTwo',
      'TableThree'
  )
ORDER BY TABLE_NAME, ORDINAL_POSITION;
