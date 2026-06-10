--STATUS MIX
CREATE PROCEDURE "C_DASHBOARD_STATUSMIX"
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT 'Success' AS "Label",
           COUNT(*)  AS "Value"
    FROM OINV
    WHERE "DocDate" BETWEEN :FromDate AND :ToDate
      AND "U_STATUS1" = 'SUCCESS'

    UNION ALL

    SELECT 'Failed',
           COUNT(*)
    FROM OINV
    WHERE "DocDate" BETWEEN :FromDate AND :ToDate
      AND "U_STATUS1" = 'FAILED'

    UNION ALL

    SELECT 'Pending',
           COUNT(*)
    FROM OINV
    WHERE "DocDate" BETWEEN :FromDate AND :ToDate
      AND (
            "U_STATUS1" IS NULL
            OR TRIM("U_STATUS1") = ''
          );

END;
 