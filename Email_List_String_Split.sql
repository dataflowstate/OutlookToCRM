WITH SplitTo AS (
    SELECT
        RecordID,
        Subject,
        value AS ToName,
        ROW_NUMBER() OVER (PARTITION BY RecordID ORDER BY (SELECT NULL)) AS RowNum
    FROM [dbo].[SentEmailContacts]
    CROSS APPLY STRING_SPLIT(ToName, ';')
),
SplitToAddress AS (
    SELECT
        RecordID,
        value AS ToAddress,
        ROW_NUMBER() OVER (PARTITION BY RecordID ORDER BY (SELECT NULL)) AS RowNum
    FROM [dbo].[SentEmailContacts]
    CROSS APPLY STRING_SPLIT(ToAddress, ';')
),
SplitCC AS (
    SELECT
        RecordID,
        Subject,
        value AS CCName,
        ROW_NUMBER() OVER (PARTITION BY RecordID ORDER BY (SELECT NULL)) AS RowNum
    FROM [dbo].[SentEmailContacts]
    CROSS APPLY STRING_SPLIT(CCName, ';')
),
SplitCCAddress AS (
    SELECT
        RecordID,
        value AS CCAddress,
        ROW_NUMBER() OVER (PARTITION BY RecordID ORDER BY (SELECT NULL)) AS RowNum
    FROM [dbo].[SentEmailContacts]
    CROSS APPLY STRING_SPLIT(CCAddress, ';')
)
-- Combine ToName and ToAddress
SELECT 
    t.RecordID,
    t.Subject,
    t.ToName,
    ta.ToAddress,
    NULL AS CCName,  -- Placeholder for CC columns in this union
    NULL AS CCAddress
FROM SplitTo t
INNER JOIN SplitToAddress ta
    ON t.RecordID = ta.RecordID AND t.RowNum = ta.RowNum

UNION ALL

-- Combine CCName and CCAddress
SELECT 
    c.RecordID,
    c.Subject,
    NULL AS ToName,  -- Placeholder for To columns in this union
    NULL AS ToAddress,
    c.CCName,
    ca.CCAddress
FROM SplitCC c
INNER JOIN SplitCCAddress ca
    ON c.RecordID = ca.RecordID AND c.RowNum = ca.RowNum
ORDER BY RecordID, Subject, ToName, CCName;
