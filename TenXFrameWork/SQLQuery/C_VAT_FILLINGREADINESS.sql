--VAT Filling Readiness

CREATE PROCEDURE C_VAT_FILLINGREADINESS
(
    IN FromDate DATE,
    IN ToDate DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT
        'Sales invoices posted' AS "Checklist",
        CASE
            WHEN COUNT(*) > 0 THEN 'Complete'
            ELSE 'Pending'
        END AS "Status",
        COUNT(*) AS "Count"
    FROM OINV
    WHERE "CANCELED" = 'N'
        AND "DocDate" BETWEEN :FromDate AND :ToDate

    UNION ALL

    SELECT
        'Purchase invoices posted' AS "Checklist",
        CASE
            WHEN COUNT(*) > 0 THEN 'Complete'
            ELSE 'Pending'
        END AS "Status",
        COUNT(*) AS "Count"
    FROM OPCH
    WHERE "CANCELED" = 'N'
        AND "DocDate" BETWEEN :FromDate AND :ToDate

    UNION ALL

    SELECT
        'Supplier TRN validation' AS "Checklist",
        CASE
            WHEN COUNT(*) = 0 THEN 'Complete'
            ELSE 'Issues Found'
        END AS "Status",
        COUNT(*) AS "Count"
    FROM OPCH T0
    INNER JOIN OCRD C
        ON T0."CardCode" = C."CardCode"
    WHERE T0."CANCELED" = 'N'
        AND T0."DocDate" BETWEEN :FromDate AND :ToDate
        AND IFNULL(C."LicTradNum", '') = ''

    UNION ALL

    SELECT
        'Tax code validation' AS "Checklist",
        CASE
            WHEN COUNT(*) = 0 THEN 'Complete'
            ELSE 'Issues Found'
        END AS "Status",
        COUNT(*) AS "Count"
    FROM
    (
        SELECT T1."TaxCode"
        FROM OINV T0
        INNER JOIN INV1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate
            AND IFNULL(T1."TaxCode", '') = ''

        UNION ALL

        SELECT T1."TaxCode"
        FROM OPCH T0
        INNER JOIN PCH1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
            AND T0."DocDate" BETWEEN :FromDate AND :ToDate
            AND IFNULL(T1."TaxCode", '') = ''

    ) X;

END;

 