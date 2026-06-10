--VAT BY  BUSINESS UNIT
CREATE PROCEDURE C_VAT_BUSINESSUNIT
(
    IN FromDate DATE,
    IN ToDate DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT
        IFNULL(B."BPLName", 'Head Office') AS "BusinessUnit",
        SUM(T1."VatSum") AS "OutputVAT"
    FROM OINV T0
    INNER JOIN INV1 T1
        ON T0."DocEntry" = T1."DocEntry"
    LEFT JOIN OBPL B
        ON T0."BPLId" = B."BPLId"
    WHERE T0."CANCELED" = 'N'
        AND T0."DocDate" BETWEEN :FromDate AND :ToDate
    GROUP BY IFNULL(B."BPLName", 'Head Office')
    ORDER BY SUM(T1."VatSum") DESC;

END;
 