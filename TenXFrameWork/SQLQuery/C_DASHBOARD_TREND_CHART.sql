--TREND CHART
CREATE PROCEDURE "C_DASHBOARD_TREND_CHART"
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT
        TO_NVARCHAR("DocDate",'DD Mon') AS "Day",

        SUM(
            CASE
                WHEN "U_STATUS1" = 'SUCCESS'
                THEN 1
                ELSE 0
            END
        ) AS "Success",

        SUM(
            CASE
                WHEN "U_STATUS1" = 'FAILED'
                THEN 1
                ELSE 0
            END
        ) AS "Failure",

        SUM(
            CASE
                WHEN "U_STATUS1" IS NULL
                     OR TRIM("U_STATUS1") = ''
                THEN 1
                ELSE 0
            END
        ) AS "Pending"

    FROM OINV
    WHERE "DocDate" BETWEEN :FromDate AND :ToDate

    GROUP BY "DocDate"
    ORDER BY "DocDate";

END;

 