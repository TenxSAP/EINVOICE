--Output VS Input Vat, & NET VAT 
CREATE PROCEDURE C_VAT_MONTHLYVATCHART
(
    IN FromDate DATE,
    IN ToDate DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT
        "Period",
        SUM("OutputVAT") AS "OutputVAT",
        SUM("InputVAT") AS "InputVAT",
        SUM("OutputVAT") - SUM("InputVAT") AS "NetVATPayable"
    FROM
    (

        /* AR Invoice */
        SELECT
            TO_VARCHAR(T0."DocDate",'YYYY-MM') AS "Period",
            SUM(T1."VatSum") AS "OutputVAT",
            0 AS "InputVAT"
        FROM OINV T0
        INNER JOIN INV1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate
        GROUP BY TO_VARCHAR(T0."DocDate",'YYYY-MM')

        UNION ALL

        /* AR Credit Memo */
        SELECT
            TO_VARCHAR(T0."DocDate",'YYYY-MM') AS "Period",
            SUM(T1."VatSum" * -1) AS "OutputVAT",
            0 AS "InputVAT"
        FROM ORIN T0
        INNER JOIN RIN1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate
        GROUP BY TO_VARCHAR(T0."DocDate",'YYYY-MM')

        UNION ALL

        /* AP Invoice */
        SELECT
            TO_VARCHAR(T0."DocDate",'YYYY-MM') AS "Period",
            0 AS "OutputVAT",
            SUM(T1."VatSum") AS "InputVAT"
        FROM OPCH T0
        INNER JOIN PCH1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate
        GROUP BY TO_VARCHAR(T0."DocDate",'YYYY-MM')

        UNION ALL

        /* AP Credit Memo */
        SELECT
            TO_VARCHAR(T0."DocDate",'YYYY-MM') AS "Period",
            0 AS "OutputVAT",
            SUM(T1."VatSum" * -1) AS "InputVAT"
        FROM ORPC T0
        INNER JOIN RPC1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate
        GROUP BY TO_VARCHAR(T0."DocDate",'YYYY-MM')

    ) X

    GROUP BY "Period"
    ORDER BY "Period";

END;

 
