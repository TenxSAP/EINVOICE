CREATE PROCEDURE "TNX_E_OINV"
(
    IN DocEntry NVARCHAR(30)
)
AS
BEGIN

SELECT

-- ================= HEADER =================
T0."DocEntry",
'Invoice_' || T0."DocNum" AS "ReferenceNumber",
YEAR(T0."DocDate") AS "FinancialYear",
'Peppol' AS "edelivery",
'Invoice_' || T0."DocNum" AS "InvoiceID",
T0."DocDate" AS "IssueDate",
T0."DocDueDate" AS "DueDate",
'380' AS "InvoiceTypeCode",
IFNULL(T0."Comments",'') AS "Note",
T0."DocCur" AS "DocumentCurrencyCode",
T0."DocCur" AS "VATCurrencyCode",
IFNULL(T0."Project",'') AS "AccountingCost",
IFNULL(T0."NumAtCard",'') AS "BuyerReference",

CASE 
    WHEN IFNULL(T0."NumAtCard",'') = '' 
        THEN CAST(T0."DocNum" AS NVARCHAR)
    ELSE ''
END AS "Purchaseorderreference",
-- ================= INVOICE PERIOD =================
T0."DocDate" AS "StartDate",
T0."DocDate" AS "EndDate",

-- ================= EXTRA =================
'' AS "Invoicetransactiontypecode",
T0."TaxDate" AS "VATtaxpointdate",
ADM."TaxIdNum" AS "PrincipleID",
'' AS "BeneficiaryID",
T0."DocRate" AS "Currencyexchangerate",

'' AS "Salesorderreference",
'' AS "Despatchadvicereference",
'' AS "Receivingadvicereference",
'' AS "Customsreferencenumber",
'' AS "FrequencyofBilling",

-- ================= BILLING =================
'' AS "RefInvoiceID",
T0."DocDate" AS "RefIssueDate",
'' AS "CreditReason",

-- ================= ADDITIONAL DOCUMENT =================
IFNULL(T0."NumAtCard",'') AS "AdditionalDocumentReferenceID",
'OP' AS "SchemeID",
'130' AS "DocumentTypecode",
'PO Reference' AS "DocumentDescription",
'' AS "EmbeddedDocumentBinaryObject",
'' AS "mimeCode",
'' AS "FileName",
IFNULL(T0."NumAtCard",'') AS "ExternalReferenceID",

-- ================= SUPPLIER =================
'' AS "SellerCode",
'0235' AS "SupplierSchemeID",
ADM."TaxIdNum" AS "SellerElectronicAddress",

'' AS "SellerLegalSchemeIdentifier",
'EID' AS "SellerLegalRegistrationType",
'' AS "SellerTradeLicense",
'' AS "SellerEmiratesID",
'' AS "SellerPassport",
'' AS "SellerCabinetDecision",
'' AS "SellerAuthorityname",
'' AS "SellerPassportCountry",

IFNULL(AD1."Street",'NA') AS "SupplierStreet",
'' AS "SupplierStreet2",
IFNULL(AD1."City",'NA') AS "SupplierCity",
'' AS "SupplierArea",
IFNULL(AD1."ZipCode",'00000') AS "SupplierZip",

ADM."Country" AS "SupplierCountry",
ADM."TaxIdNum" AS "SupplierVAT",
'' AS "SupplierTaxScheme",
ADM."CompnyName" AS "SupplierName",
'' AS "SupplierCompanyID",
''AS "SupplierContactName",
ADM."E_Mail" AS "SupplierEmail",
ADM."Phone1" AS "SupplierPhone",


 

-- ================= CUSTOMER =================
-- ================= CUSTOMER =================
'' AS "BuyerCode",

'0235' AS "CustomerSchemeID",
IFNULL(C."LicTradNum",'') AS "CustomerVAT",

-- ADDRESS
IFNULL(A."Street",'') AS "CustStreet",
'' AS "CustStreet2",
IFNULL(A."City",'') AS "CustCity",
'' AS "CustArea",
IFNULL(A."ZipCode",'') AS "CustZip",
IFNULL(A."Country",'') AS "CustCountry",

-- BASIC DETAILS
IFNULL(C."CardName",'') AS "CustomerName",
IFNULL(C."E_Mail",'') AS "CustomerEmail",
IFNULL(C."Cellular",'') AS "CustomerPhone",

-- 🔥 REQUIRED FOR YOUR VB.NET MODEL
'' AS "Buyerlegalregistrationidentifiertype",
'' AS "BuyerCommercialorTradelicense",
'' AS "BuyerEmiratesID",
'' AS "BuyerPassport",
'' AS "BuyerCabinetDecision",
'' AS "BuyerAuthorityName",
'' AS "BuyerPassportIssuingCountrycode",

-- 🔥 CONTACT (VERY IMPORTANT)
IFNULL(C."CntctPrsn",'') AS "ContactPerson",
IFNULL(C."Cellular",'') AS "Phone",
IFNULL(C."E_Mail",'') AS "Email",


-- ================= DELIVERY =================
T0."DocDate" AS "DeliveryDate",
'MAIN' AS "DeliveryName",
'LOC' AS "DeliverySchemeID",
T0."DocEntry" AS "DeliveryCode",

IFNULL(T12."StreetB",'') AS "DeliveryStreet",
IFNULL(T12."CityB",'') AS "DeliveryCity",
IFNULL(T12."ZipCodeB",'') AS "DeliveryZip",
IFNULL(T12."CountryB",'') AS "DeliveryCountry",
'' AS "DeliveryStreet2",


-- ================= PAYMENT =================
'' AS "PayeeName",
'Bank Transfer' AS "PaymentInstructions",
T0."DocDueDate" AS "PaymentDueDate",
'30' AS "PaymentMeansCode",

PT."PymntGroup" AS "PaymentTerms",

'' AS "PayeeAccountID",
'' AS "AccountID",
'' AS "AccountName",
'' AS "BranchID",

-- ================= TOTAL =================
-- ================= TOTAL =================
-- Line totals (sum of lines before tax)
SUM(T1."LineTotal") OVER (PARTITION BY T0."DocEntry") AS "LineExtensionAmount",
-- Tax exclusive
(T0."DocTotal" - T0."VatSum") AS "TaxExclusiveAmount",
-- Tax inclusive (full invoice amount)
T0."DocTotal" AS "TaxInclusiveAmount",
-- Charges (if no extra charges → 0)
0 AS "ChargeTotalAmount",
-- Prepaid (if not using → 0)
0 AS "PrepaidAmount",
-- Rounding (if not using → 0)
0 AS "PayableRoundingAmount",
-- Final payable
T0."DocTotal" AS "PayableAmount",
-- Discount total
T0."DiscSum" AS "TotalDiscountAmount",
-- ================= ALLOWANCE =================
T0."DiscSum" AS "AllowanceAmount",
'1' AS "AllowanceID",
0 AS "ChargeIndicator",                 -- 🔥 FIXED (BOOLEAN)
'95' AS "AllowanceReason",
'0' AS "Multiplier",

(T0."DocTotal" - T0."VatSum") AS "BaseAmount",   -- 🔥 IMPORTANT

-- 🔥 REQUIRED FOR VB.NET
'S' AS "TaxCategoryID",
T0."VatSum" AS "TaxPercent",

'' AS "ExemptionCode",
'' AS "ExemptionReason",
-- ================= LINE =================

T1."LineNum" + 1 AS "LineID",
T1."ItemCode" AS "ItemName",
T1."Dscription",
T1."Quantity",
T1."Price",
T1."LineTotal",

-- UOM Mapping
CASE 
WHEN T1."UomCode" IN ('Manual','Nos') THEN 'EA'
WHEN T1."UomCode" = 'KG' THEN 'KGM'
WHEN T1."UomCode" = 'L' THEN 'LTR'
ELSE 'EA'
END AS "UomCode",

-- TAX
CASE T1."VatGroup"
WHEN 'O1' THEN 'S'
WHEN 'O2' THEN 'O'
WHEN 'X0' THEN 'E'
ELSE 'Z'
END AS "TaxCode",

T1."VatPrcnt" AS "TaxPercent",
'VAT' AS "TaxSchemeId",
'' AS "TaxExemptionReasonCode",
'' AS "TaxExemptionReason",

-- ITEM TYPE / SERVICE
'' AS "GoodsType",
'' AS "ServiceCode",
'UNSPSC' AS "ServiceAccountingSchemeidentifier",
'Goods' AS "ItemType",

-- STANDARD ITEM IDENTIFICATION
'' AS "StandardSchemeID",
'' AS "StandardSchemeCode",

-- COMMODITY
'' AS "CommodityListID",
'' AS "CommodityCode",

-- LINE ALLOWANCE
'' AS "LineAllowanceID",
'false' AS "LineChargeIndicator",
'' AS "LineAllowanceReason",
'' AS "LineMultiplier",
'' AS "LineCurrencyID",
'' AS "LineAllowanceAmount",
'' AS "LineBaseCurrencyID",
'' AS "LineBaseAmount",

-- PRICE DETAILS
'1' AS "BaseQuantity",
'' AS "BaseQuantityUOM",
T1."Price" AS "Price",
'' AS "AlwChgAmt",
T1."LineTotal" AS "AlwChgBaseAmt"

FROM OINV T0
INNER JOIN INV1 T1 ON T0."DocEntry" = T1."DocEntry"
LEFT JOIN OADM ADM ON 1=1
LEFT JOIN ADM1 AD1 ON 1=1
LEFT JOIN OCRD C ON C."CardCode" = T0."CardCode"
LEFT JOIN CRD1 A ON A."CardCode" = C."CardCode" AND A."AdresType"='B'
LEFT JOIN INV12 T12 ON T12."DocEntry" = T0."DocEntry"
LEFT JOIN OCTG PT ON PT."GroupNum" = T0."GroupNum"
LEFT JOIN OITM ITM ON ITM."ItemCode" = T1."ItemCode"

WHERE T0."DocEntry" = :DocEntry;

END;