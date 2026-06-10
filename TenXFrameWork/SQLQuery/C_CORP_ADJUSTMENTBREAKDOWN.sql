
CREATE PROCEDURE C_CORP_ADJUSTMENTBREAKDOWN
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    /* Corporate Tax Adjustment Breakdown */

    SELECT
        T0."U_CTax" AS "AdjustmentType",

        SUM
        (
            IFNULL(T0."U_CTaxVal", 0)
        ) AS "AdjustmentAmount"

    FROM "@TNX_CTAXCALCU" T0

    WHERE T0."U_DST" = 'A'

      AND T0."U_FDate"
            BETWEEN :FromDate AND :ToDate

    GROUP BY T0."U_CTax"

    ORDER BY
        SUM(IFNULL(T0."U_CTaxVal", 0)) DESC;

END;
 