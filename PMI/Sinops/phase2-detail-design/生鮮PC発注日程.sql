WITH days AS (
    SELECT
        d.DeliveryDate,
        (DATEDIFF(day, CONVERT(date, '18991231'), d.DeliveryDate) % 7) + 1 AS [WeekDay]
    FROM (
        SELECT DATEADD(day, n.i, CONVERT(date, GETDATE())) AS DeliveryDate
        FROM (
            SELECT TOP (36) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) - 8 AS i
            FROM (VALUES (0),(0),(0),(0),(0),(0),(0),(0),(0),(0)) AS a(n)
            CROSS JOIN (VALUES (0),(0),(0),(0)) AS b(n)
        ) AS n
    ) AS d
),
base AS (
    SELECT
        s.BranchCD,
        s.LineCD,
        M.DivisionCD,
        s.SourceFlag AS SupplierCD,
        s.Bin,
        s.[WeekDay],
        s.BasicLockDays,
        t.LockTime,
        CASE WHEN v.lt < 1 THEN 1 ELSE v.lt END AS 理想LT,
        MAX(CASE WHEN s.Bin <> 0 THEN 1 ELSE 0 END) OVER (
            PARTITION BY s.BranchCD, s.LineCD, s.SourceFlag, s.[WeekDay]
        ) AS HasOtherBin
    FROM dbOrder.dbo.shT_ProductionPlanSched AS s
    inner join MasterDB.dbo.Lines M
    ON
    	s.lineCD=M.LineCD
    CROSS APPLY (SELECT RTRIM(s.BasicLockTime) AS LockTime) AS t
    CROSS APPLY (
        SELECT s.BasicLockDays + CASE WHEN t.LockTime < '02:30:00' THEN 1 ELSE 0 END AS lt
    ) AS v
    WHERE s.BasicLockTime IS NOT NULL
      AND (s.BranchCD = 999 OR s.BranchCD > 1500)
      AND EXISTS (
            SELECT 1
            FROM dbOrder.dbo.shM_SupplierPC AS p
            WHERE p.SupplierCD = s.SourceFlag
      )
),
-- TOP 让每个店、Line、供应商、星期先定成一行再复用，避免日程表被重复扫描。
pick AS (
    SELECT TOP (2147483647)
        BranchCD, DivisionCD, SupplierCD, [WeekDay], BasicLockDays, LockTime, 理想LT
    FROM (
        SELECT
            BranchCD, DivisionCD, SupplierCD, [WeekDay], BasicLockDays, LockTime, 理想LT,
            ROW_NUMBER() OVER (
                PARTITION BY BranchCD, DivisionCD, SupplierCD, [WeekDay]
                ORDER BY COUNT(*) DESC, 理想LT DESC, BasicLockDays DESC, LockTime
            ) AS rn
        FROM base
        WHERE Bin <> 0 OR HasOtherBin = 0
        GROUP BY BranchCD, DivisionCD, SupplierCD, [WeekDay], BasicLockDays, LockTime, 理想LT
    ) AS g
    WHERE rn = 1
    ORDER BY DivisionCD, SupplierCD, BranchCD, [WeekDay]
),
sched AS (
    SELECT BranchCD, DivisionCD, SupplierCD, [WeekDay], BasicLockDays, LockTime, 理想LT
    FROM (
        SELECT
            BranchCD, DivisionCD, SupplierCD, [WeekDay], BasicLockDays, LockTime, 理想LT,
            ROW_NUMBER() OVER (
                PARTITION BY BranchCD, DivisionCD, SupplierCD, [WeekDay]
                ORDER BY pri
            ) AS rn
        FROM (
            SELECT
                BranchCD, DivisionCD, SupplierCD, [WeekDay], BasicLockDays, LockTime, 理想LT,
                0 AS pri
            FROM pick
            UNION ALL
            SELECT
                st.BranchCD, cm.DivisionCD, cm.SupplierCD, cm.[WeekDay],
                cm.BasicLockDays, cm.LockTime, cm.理想LT,
                1
            FROM (
                SELECT DISTINCT BranchCD, DivisionCD, SupplierCD
                FROM pick
                WHERE BranchCD <> 999
            ) AS st
            INNER JOIN pick AS cm
                ON  cm.BranchCD = 999
                AND cm.DivisionCD = st.DivisionCD
                AND cm.SupplierCD = st.SupplierCD
        ) AS u
    ) AS x
    WHERE rn = 1
)
SELECT
    RIGHT('0000' + CONVERT(varchar(4), r.BranchCD), 4) AS BranchCD,
    RIGHT('0000' + CONVERT(varchar(4), r.DivisionCD), 4) AS DivisionCD,
    RIGHT('000000000' + CONVERT(varchar(9), r.SupplierCD), 9) AS SupplierCD,
    CONVERT(char(8), DATEADD(day, -r.理想LT, d.DeliveryDate), 112) AS 发注日,
    CONVERT(char(8), d.DeliveryDate, 112) AS 纳品日,
    CONVERT(char(8), DATEADD(day, -r.BasicLockDays, d.DeliveryDate), 112)
        + '_' + REPLACE(LEFT(r.LockTime, 5), ':', '') AS 截止时间,
    r.理想LT,
    r.BasicLockDays AS 实际LT
FROM sched AS r
INNER JOIN days AS d
    ON d.[WeekDay] = r.[WeekDay]
ORDER BY r.SupplierCD, r.DivisionCD, r.BranchCD, d.DeliveryDate;
