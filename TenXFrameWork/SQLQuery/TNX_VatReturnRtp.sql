CREATE PROCEDURE "TNX_VAT_RETURN_RPT"
(
    IN I_FromDate DATE,
    IN I_ToDate   DATE
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

WITH HEADER AS
(
    SELECT
        CURRENT_DATE AS "SubmissionDate",
        IFNULL("TaxIdNum",'') AS "TRN",
        IFNULL("CompnyName",'') AS "TaxablePersonNameEnglish",
        IFNULL("PrintHeadr",'') AS "TaxablePersonNameArabic",
        IFNULL("CompnyAddr",'') AS "TaxablePersonAddress",
        'O' AS "DocumentStatus",
        CASE 
            WHEN UPPER(IFNULL("Country",'')) = 'ABU DHABI' THEN 'Abu Dhabi'
            WHEN UPPER(IFNULL("Country",'')) = 'SHARJAH' THEN 'Sharjah'
            WHEN UPPER(IFNULL("Country",'')) = 'AJMAN' THEN 'Ajman'
            WHEN UPPER(IFNULL("Country",'')) = 'UMM AL QUWAIN' THEN 'Umm Al Quwain'
            WHEN UPPER(IFNULL("Country",'')) = 'RAS AL KHAIMAH' THEN 'Ras Al Khaimah'
            WHEN UPPER(IFNULL("Country",'')) = 'FUJAIRAH' THEN 'Fujairah'
            ELSE 'Dubai'
        END AS "CompanyEmirate"
    FROM OADM
),

SALES AS
(
    SELECT 
        IFNULL(T1."AcctCode",'') AS "AccountNo",
        T1."LineTotal" AS "Amount",
        T1."VatSum" AS "VATAmount"
    FROM OINV T0
    INNER JOIN INV1 T1 ON T0."DocEntry" = T1."DocEntry"
    WHERE T0."CANCELED" = 'N'
      AND T0."DocDate" BETWEEN :I_FromDate AND :I_ToDate

    UNION ALL

    SELECT 
        IFNULL(T1."AcctCode",'') AS "AccountNo",
        -T1."LineTotal" AS "Amount",
        -T1."VatSum" AS "VATAmount"
    FROM ORIN T0
    INNER JOIN RIN1 T1 ON T0."DocEntry" = T1."DocEntry"
    WHERE T0."CANCELED" = 'N'
      AND T0."DocDate" BETWEEN :I_FromDate AND :I_ToDate
),

PURCHASE AS
(
    SELECT 
        IFNULL(T1."AcctCode",'') AS "AccountNo",
        T1."LineTotal" AS "Amount",
        T1."VatSum" AS "VATAmount"
    FROM OPCH T0
    INNER JOIN PCH1 T1 ON T0."DocEntry" = T1."DocEntry"
    WHERE T0."CANCELED" = 'N'
      AND T0."DocDate" BETWEEN :I_FromDate AND :I_ToDate

    UNION ALL

    SELECT 
        IFNULL(T1."AcctCode",'') AS "AccountNo",
        -T1."LineTotal" AS "Amount",
        -T1."VatSum" AS "VATAmount"
    FROM ORPC T0
    INNER JOIN RPC1 T1 ON T0."DocEntry" = T1."DocEntry"
    WHERE T0."CANCELED" = 'N'
      AND T0."DocDate" BETWEEN :I_FromDate AND :I_ToDate
),

SALES_ACC AS
(
    SELECT IFNULL(MIN("AccountNo"), '') AS "SalesAccountNo"
    FROM SALES
    WHERE IFNULL("VATAmount", 0) <> 0
      AND IFNULL("AccountNo", '') <> ''
),

PURCHASE_ACC AS
(
    SELECT IFNULL(MIN("AccountNo"), '') AS "PurchaseAccountNo"
    FROM PURCHASE
    WHERE IFNULL("VATAmount", 0) <> 0
      AND IFNULL("AccountNo", '') <> ''
),

SUMMARY AS
(
    SELECT
        H."CompanyEmirate",
        SA."SalesAccountNo",
        PA."PurchaseAccountNo",

        IFNULL((SELECT SUM("Amount") FROM SALES WHERE IFNULL("VATAmount",0) <> 0),0) AS "SalesAmount",
        IFNULL((SELECT SUM("VATAmount") FROM SALES WHERE IFNULL("VATAmount",0) <> 0),0) AS "SalesVAT",
        IFNULL((SELECT SUM("Amount") FROM SALES WHERE IFNULL("VATAmount",0) = 0),0) AS "ZeroSalesAmount",

        IFNULL((SELECT SUM("Amount") FROM PURCHASE WHERE IFNULL("VATAmount",0) <> 0),0) AS "PurchaseAmount",
        IFNULL((SELECT SUM("VATAmount") FROM PURCHASE WHERE IFNULL("VATAmount",0) <> 0),0) AS "PurchaseVAT"
    FROM HEADER H
    CROSS JOIN SALES_ACC SA
    CROSS JOIN PURCHASE_ACC PA
),

FINAL AS
(
    SELECT '1a' AS "BoxNo",
           CASE WHEN "CompanyEmirate" = 'Abu Dhabi' THEN "SalesAccountNo" ELSE '' END AS "AccountNo",
           'Standard rated supplies in Abu Dhabi' AS "Description",
           CASE WHEN "CompanyEmirate" = 'Abu Dhabi' THEN "SalesAmount" ELSE 0 END AS "Amount_AED",
           CASE WHEN "CompanyEmirate" = 'Abu Dhabi' THEN "SalesVAT" ELSE 0 END AS "VAT_AED",
           0 AS "Adjustment_AED"
    FROM SUMMARY

    UNION ALL SELECT '1b',
           CASE WHEN "CompanyEmirate" = 'Dubai' THEN "SalesAccountNo" ELSE '' END,
           'Standard rated supplies in Dubai',
           CASE WHEN "CompanyEmirate" = 'Dubai' THEN "SalesAmount" ELSE 0 END,
           CASE WHEN "CompanyEmirate" = 'Dubai' THEN "SalesVAT" ELSE 0 END,
           0 FROM SUMMARY

    UNION ALL SELECT '1c',
           CASE WHEN "CompanyEmirate" = 'Sharjah' THEN "SalesAccountNo" ELSE '' END,
           'Standard rated supplies in Sharjah',
           CASE WHEN "CompanyEmirate" = 'Sharjah' THEN "SalesAmount" ELSE 0 END,
           CASE WHEN "CompanyEmirate" = 'Sharjah' THEN "SalesVAT" ELSE 0 END,
           0 FROM SUMMARY

    UNION ALL SELECT '1d',
           CASE WHEN "CompanyEmirate" = 'Ajman' THEN "SalesAccountNo" ELSE '' END,
           'Standard rated supplies in Ajman',
           CASE WHEN "CompanyEmirate" = 'Ajman' THEN "SalesAmount" ELSE 0 END,
           CASE WHEN "CompanyEmirate" = 'Ajman' THEN "SalesVAT" ELSE 0 END,
           0 FROM SUMMARY

    UNION ALL SELECT '1e',
           CASE WHEN "CompanyEmirate" = 'Umm Al Quwain' THEN "SalesAccountNo" ELSE '' END,
           'Standard rated supplies in Umm Al Quwain',
           CASE WHEN "CompanyEmirate" = 'Umm Al Quwain' THEN "SalesAmount" ELSE 0 END,
           CASE WHEN "CompanyEmirate" = 'Umm Al Quwain' THEN "SalesVAT" ELSE 0 END,
           0 FROM SUMMARY

    UNION ALL SELECT '1f',
           CASE WHEN "CompanyEmirate" = 'Ras Al Khaimah' THEN "SalesAccountNo" ELSE '' END,
           'Standard rated supplies in Ras Al Khaimah',
           CASE WHEN "CompanyEmirate" = 'Ras Al Khaimah' THEN "SalesAmount" ELSE 0 END,
           CASE WHEN "CompanyEmirate" = 'Ras Al Khaimah' THEN "SalesVAT" ELSE 0 END,
           0 FROM SUMMARY

    UNION ALL SELECT '1g',
           CASE WHEN "CompanyEmirate" = 'Fujairah' THEN "SalesAccountNo" ELSE '' END,
           'Standard rated supplies in Fujairah',
           CASE WHEN "CompanyEmirate" = 'Fujairah' THEN "SalesAmount" ELSE 0 END,
           CASE WHEN "CompanyEmirate" = 'Fujairah' THEN "SalesVAT" ELSE 0 END,
           0 FROM SUMMARY

    UNION ALL SELECT '2', '', 'Supplies subject to reverse charge provisions', 0, 0, 0 FROM SUMMARY
    UNION ALL SELECT '3', '', 'Zero rated supplies', "ZeroSalesAmount", 0, 0 FROM SUMMARY
    UNION ALL SELECT '4', '', 'Supplies to registered customers in other GCC implementing states', 0, 0, 0 FROM SUMMARY
    UNION ALL SELECT '5', '', 'Exempt supplies', 0, 0, 0 FROM SUMMARY
    UNION ALL SELECT '6', '', 'Import VAT accounted through UAE customs', 0, 0, 0 FROM SUMMARY
    UNION ALL SELECT '7', '', 'Amendments or corrections to Output figures', 0, 0, 0 FROM SUMMARY

    UNION ALL SELECT '8',
           "SalesAccountNo",
           'Output Totals A1 / A2',
           "SalesAmount" + "ZeroSalesAmount",
           "SalesVAT",
           0 FROM SUMMARY

    UNION ALL SELECT '9',
           "PurchaseAccountNo",
           'Standard rated expenses',
           "PurchaseAmount",
           "PurchaseVAT",
           0 FROM SUMMARY

    UNION ALL SELECT '10', '', 'Expenses subject to reverse charge provisions', 0, 0, 0 FROM SUMMARY
    UNION ALL SELECT '11', '', 'Amendments or corrections to Input figures', 0, 0, 0 FROM SUMMARY

    UNION ALL SELECT '12A',
           "PurchaseAccountNo",
           'Input Totals A3 / A4',
           "PurchaseAmount",
           "PurchaseVAT",
           0 FROM SUMMARY

    UNION ALL SELECT '12', '', 'Total value of due tax for the period A5', 0, "SalesVAT", 0 FROM SUMMARY
    UNION ALL SELECT '13', '', 'Total value of recoverable tax for the period A6', 0, "PurchaseVAT", 0 FROM SUMMARY
    UNION ALL SELECT '14', '', 'Net VAT due or reclaimed for the period A7', 0, "SalesVAT" - "PurchaseVAT", 0 FROM SUMMARY
)

SELECT
    H."SubmissionDate",
    H."TRN",
    H."TaxablePersonNameEnglish",
    H."TaxablePersonNameArabic",
    H."TaxablePersonAddress",
    H."DocumentStatus",
    :I_FromDate AS "VATReturnPeriodFrom",
    :I_ToDate AS "VATReturnPeriodTo",
    F."BoxNo",
    F."AccountNo",
    F."Description",
    ROUND(IFNULL(F."Amount_AED",0),3) AS "Amount_AED",
    ROUND(IFNULL(F."VAT_AED",0),3) AS "VAT_AED",
    ROUND(IFNULL(F."Adjustment_AED",0),3) AS "Adjustment_AED"
FROM FINAL F
CROSS JOIN HEADER H
ORDER BY
    CASE F."BoxNo"
        WHEN '1a' THEN 1
        WHEN '1b' THEN 2
        WHEN '1c' THEN 3
        WHEN '1d' THEN 4
        WHEN '1e' THEN 5
        WHEN '1f' THEN 6
        WHEN '1g' THEN 7
        WHEN '2' THEN 8
        WHEN '3' THEN 9
        WHEN '4' THEN 10
        WHEN '5' THEN 11
        WHEN '6' THEN 12
        WHEN '7' THEN 13
        WHEN '8' THEN 14
        WHEN '9' THEN 15
        WHEN '10' THEN 16
        WHEN '11' THEN 17
        WHEN '12A' THEN 18
        WHEN '12' THEN 19
        WHEN '13' THEN 20
        WHEN '14' THEN 21
        ELSE 99
    END;

END;