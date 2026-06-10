--VAT EXCEPTION DETAILS
CREATE PROCEDURE C_VAT_EXCEPTIONDETAILS
(
    IN FromDate DATE,
    IN ToDate DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT
        T0."DocNum" AS "DocumentNo",
        'Sales Invoice' AS "Type",
        T0."CardName" AS "PartyName",
        C."LicTradNum" AS "TRN",

        CASE
            WHEN IFNULL(C."LicTradNum",'') = '' THEN 'Customer TRN Missing'
            WHEN IFNULL(T1."TaxCode",'') = '' THEN 'Tax Code Missing'
            WHEN T1."VatSum" = 0
                 AND T1."LineTotal" > 0 THEN 'Zero VAT on Taxable Line'
            ELSE 'OK'
        END AS "Issue",

        T1."VatSum" AS "VATAmount",
        T0."DocDate"

    FROM OINV T0

    INNER JOIN INV1 T1
        ON T0."DocEntry" = T1."DocEntry"

    INNER JOIN OCRD C
        ON T0."CardCode" = C."CardCode"

    WHERE T0."CANCELED" = 'N'
        AND T0."DocDate" BETWEEN :FromDate AND :ToDate
        AND
        (
            IFNULL(C."LicTradNum",'') = ''
            OR IFNULL(T1."TaxCode",'') = ''
            OR (T1."VatSum" = 0 AND T1."LineTotal" > 0)
        )

    UNION ALL

    SELECT
        T0."DocNum" AS "DocumentNo",
        'Purchase Invoice' AS "Type",
        T0."CardName" AS "PartyName",
        C."LicTradNum" AS "TRN",

        CASE
            WHEN IFNULL(C."LicTradNum",'') = '' THEN 'Supplier TRN Missing'
            WHEN IFNULL(T1."TaxCode",'') = '' THEN 'Tax Code Missing'
            WHEN T1."VatSum" = 0
                 AND T1."LineTotal" > 0 THEN 'Input VAT Not Eligible'
            ELSE 'OK'
        END AS "Issue",

        T1."VatSum" AS "VATAmount",
        T0."DocDate"

    FROM OPCH T0

    INNER JOIN PCH1 T1
        ON T0."DocEntry" = T1."DocEntry"

    INNER JOIN OCRD C
        ON T0."CardCode" = C."CardCode"

    WHERE T0."CANCELED" = 'N'
        AND T0."DocDate" BETWEEN :FromDate AND :ToDate
        AND
        (
            IFNULL(C."LicTradNum",'') = ''
            OR IFNULL(T1."TaxCode",'') = ''
            OR (T1."VatSum" = 0 AND T1."LineTotal" > 0)
        )

    ORDER BY "DocDate" DESC;

END;