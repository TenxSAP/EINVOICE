CREATE PROCEDURE "TNX_ZATCA_ORIN"
(
    IN DocEntry NVARCHAR(30)
)
AS
BEGIN

    SELECT
        /* ---------------- Header ---------------- */
        'save'                                      AS "type",
        'SCN-' || T0."DocNum"                      AS "credit_note_number",
        T0."DocDate"                               AS "credit_note_date",
        T1."BaseRef"                               AS "invoice_number",
        T0."DocDate"                               AS "invoice_date",
        T0."U_REASON"                              AS "reason_for_issuance",
        PT."PymntGroup"                            AS "payment_terms",

        /* ---------------- Billed To ---------------- */
        ''                                         AS "other_id_type",
        ''                                         AS "other_id",
        T0."CardName"                              AS "name",
        C."E_Mail"                                 AS "email",
        C."Cellular"                               AS "phone",
        T12."CountryB"                             AS "country_code",
        CR."Name"                                  AS "country",

        /* ---------------- Item Details ---------------- */
        T1."ItemCode"                              AS "item_id",
        CASE 
        T1."VatGroup" 
        WHEN 'O1' THEN 'S'
        WHEN 'O2' THEN 'O'
        WHEN 'X0' THEN 'E'
        WHEN 'A1' THEN 'A'
        ELSE 'Z'       END                    	   AS "vat_category_code",
        ''                                         AS "vat_exemption_code",
        T1."VatPrcnt" / 100                        AS "vat_rate",
        T1."Dscription"                            AS "item_name",
        T1."DiscPrcnt"                             AS "item_price_discount",
        'PERCENTAGE'                               AS "price_allowance_indicator",
        T1."Quantity"                              AS "quantity",
        T1."Price"                                 AS "unit_rate",
        CASE 
            WHEN IFNULL(T1."DiscPrcnt",0) > 0 THEN 'true'
            ELSE 'false'
        END                                        AS "is_discount_applicable",
        T1."UomCode"                               AS "unit_of_measure",

        /* ---------------- Other ---------------- */
        'SYSTEM_GENERATED'                         AS "document_number_type",
        '30'                                       AS "payment_means",
        T0."DocCur"                                AS "invoice_currency",
        'B2C'                                      AS "business_kind"
        

    FROM ORIN T0
    INNER JOIN RIN1  T1  ON T0."DocEntry" = T1."DocEntry"
    INNER JOIN RIN12 T12 ON T0."DocEntry" = T12."DocEntry"
    INNER JOIN OCRD  C   ON C."CardCode"  = T0."CardCode"
    LEFT  JOIN OCRY  CR  ON CR."Code"     = T12."CountryB"
    LEFT  JOIN OCTG  PT  ON PT."GroupNum" = T0."GroupNum"

    WHERE
        T0."DocEntry" = :DocEntry
        AND T0."CreateDate" > '2025-03-01';

END;