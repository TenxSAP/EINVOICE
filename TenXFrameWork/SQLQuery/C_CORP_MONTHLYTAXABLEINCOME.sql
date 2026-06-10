
--PROFIT TREND
CREATE PROCEDURE C_CORP_MONTHLYTAXABLEINCOME
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    WITH MONTHLY_PROFIT AS
    (
        SELECT
            TO_VARCHAR(H."RefDate", 'YYYY-MM') AS "Period",

            SUM
            (
                CASE
                    WHEN A."ActType" = 'I'
                        THEN J."Credit" - J."Debit"

                    WHEN A."ActType" = 'E'
                        THEN J."Debit" - J."Credit"

                    ELSE 0
                END
            ) AS "ProfitBeforeTax"

        FROM OJDT H

        INNER JOIN JDT1 J
            ON H."TransId" = J."TransId"

        INNER JOIN OACT A
            ON J."Account" = A."AcctCode"

        WHERE H."RefDate" BETWEEN :FromDate AND :ToDate
          AND A."ActType" IN ('I', 'E')

        GROUP BY TO_VARCHAR(H."RefDate", 'YYYY-MM')
    ),

    MONTHLY_ADJ AS
    (
        SELECT
            TO_VARCHAR(T0."U_FDate", 'YYYY-MM') AS "Period",

            IFNULL(SUM(T0."U_CTaxVal"), 0) AS "AddBackAmount",

            IFNULL(SUM(T0."U_CTaxVal"), 0) AS "DeductionAmount"

        FROM "@TNX_CTAXCALCU" T0

        WHERE T0."U_DST" = 'A'
          AND T0."U_FDate" BETWEEN :FromDate AND :ToDate

        GROUP BY TO_VARCHAR(T0."U_FDate", 'YYYY-MM')
    )

    SELECT
        P."Period",

        P."ProfitBeforeTax",

        P."ProfitBeforeTax"
            + IFNULL(A."AddBackAmount", 0)
            - IFNULL(A."DeductionAmount", 0) AS "TaxableIncome"

    FROM MONTHLY_PROFIT P

    LEFT JOIN MONTHLY_ADJ A
        ON P."Period" = A."Period"

    ORDER BY P."Period";

END;
 