CREATE PROCEDURE "TNX_PL_By_Branch"
(
    IN  FromDate DATE,
    IN  ToDate   DATE
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

    DECLARE branch_columns NVARCHAR(5000) DEFAULT '';
    DECLARE full_sql       NVARCHAR(20000) DEFAULT '';

    SELECT IFNULL(
        STRING_AGG(
            'SUM(CASE WHEN T1."BPLId" = ' || TO_NVARCHAR("BPLId") ||
            ' THEN T1."Debit" - T1."Credit" ELSE 0 END) AS "' || REPLACE("BPLName", '"', '') || '"',
            ', '
        ),
        ''
    ) INTO branch_columns
    FROM OBPL;

    full_sql := 
    'SELECT 
        T3."FatherNum"    AS "Level1_Account",
        T4."AcctName"     AS "Level1_Name",
        T1."Account"      AS "AccountCode",
        T2."AcctName"     AS "AccountName"' ||

        CASE 
            WHEN :branch_columns <> '' THEN ',' || :branch_columns
            ELSE ''
        END ||

        ', SUM(T1."Debit" - T1."Credit") AS "Total"
    FROM "JDT1" T1
    INNER JOIN "OJDT" T0 ON T1."TransId" = T0."TransId"
    INNER JOIN "OACT" T2 ON T1."Account" = T2."AcctCode"
    LEFT JOIN "OACT" T3 ON T2."FatherNum" = T3."AcctCode"
    LEFT JOIN "OACT" T4 ON T3."FatherNum" = T4."AcctCode"
    WHERE 
        T2."GroupMask" IN (4,5,6,7,8) 
        AND T0."RefDate" BETWEEN ''' || TO_NVARCHAR(FromDate) || ''' AND ''' || TO_NVARCHAR(ToDate) || '''
    GROUP BY 
        T3."FatherNum", 
        T4."AcctName", 
        T1."Account", 
        T2."AcctName"
    ORDER BY 
        T3."FatherNum", 
        T1."Account"';

    EXECUTE IMMEDIATE :full_sql;

END;