--VAT Exception Summary
CREATE PROCEDURE C_VAT_EXCEPTIONSUMMARY
(
    IN FromDate DATE,
    IN ToDate DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT
        "ExceptionType",
        COUNT(*) AS "ExceptionCount",
        SUM("VATAmount") AS "VATAmount"
    FROM
    (

        /* AR Invoice Exceptions */
        SELECT
            CASE
                WHEN IFNULL(C."LicTradNum",'') = '' THEN 'Missing TRN'
                WHEN IFNULL(T1."TaxCode",'') = '' THEN 'Missing Tax Code'
                WHEN T1."VatSum" = 0
                     AND T1."LineTotal" > 0 THEN 'Zero VAT on Taxable Line'
                ELSE 'OK'
            END AS "ExceptionType",

            T1."VatSum" AS "VATAmount"

        FROM OINV T0
        INNER JOIN INV1 T1
            ON T0."DocEntry" = T1."DocEntry"
        INNER JOIN OCRD C
            ON T0."CardCode" = C."CardCode"

        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate


        UNION ALL


        /* AP Invoice Exceptions */
        SELECT
            CASE
                WHEN IFNULL(C."LicTradNum",'') = '' THEN 'Missing TRN'
                WHEN IFNULL(T1."TaxCode",'') = '' THEN 'Missing Tax Code'
                WHEN T1."VatSum" = 0
                     AND T1."LineTotal" > 0 THEN 'Input VAT Not Eligible / Zero VAT'
                ELSE 'OK'
            END AS "ExceptionType",

            T1."VatSum" AS "VATAmount"

        FROM OPCH T0
        INNER JOIN PCH1 T1
            ON T0."DocEntry" = T1."DocEntry"
        INNER JOIN OCRD C
            ON T0."CardCode" = C."CardCode"

        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate

    ) X

    WHERE "ExceptionType" <> 'OK'

    GROUP BY "ExceptionType"

    ORDER BY COUNT(*) DESC;

END;

 