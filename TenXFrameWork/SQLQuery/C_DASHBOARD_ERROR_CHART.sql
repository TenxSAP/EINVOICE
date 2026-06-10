CREATE PROCEDURE "C_DASHBOARD_ERROR_CHART"
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT TOP 5
        IFNULL(
            LEFT(TO_NVARCHAR("U_ERRORMSG"),100),
            'Unknown'
        ) AS "Label",

        COUNT(*) AS "Value"

    FROM OINV

    WHERE "U_STATUS1" = 'FAILED'
      AND "DocDate" BETWEEN :FromDate AND :ToDate

    GROUP BY LEFT(TO_NVARCHAR("U_ERRORMSG"),100)

    ORDER BY COUNT(*) DESC;

END;
 