--lIABILITY TRENDS

CREATE PROCEDURE C_CORP_LIABILITYTREND
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    WITH MONTHLY_DATA AS
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
    )

    SELECT
        "Period",

        "ProfitBeforeTax",

        CASE
            WHEN "ProfitBeforeTax" > 375000
            THEN ("ProfitBeforeTax" - 375000) * 0.09
            ELSE 0
        END AS "EstimatedCorporateTax"

    FROM MONTHLY_DATA

    ORDER BY "Period";

END;