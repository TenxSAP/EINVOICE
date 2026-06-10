CREATE PROCEDURE "TNX_ZATCA_OINV"
(
    IN DocEntry NVARCHAR(30)
)
AS
BEGIN

    SELECT
        T0."DocEntry",
        'save'                                       AS "type",
        'INV-' || T0."DocNum" 					     AS "invoice_number",
        T0."DocDate"                                 AS "invoice_date",
        T0."CardCode"                                AS "CusCode",
        ''											AS "logo_url",        
        IFNULL(
            TO_CHAR(
                TO_TIME(LPAD(T0."DocTime", 4, '0')),
                'HH24:MI:SS'
            ),
            '00:00:00'
        )                                            AS "DocTime",
        T0."QRCodeSrc",
        T0."NumAtCard",
        '388'                                        AS "DocType",
        'simplified'                                 AS "InvType",       
        T1."LineTotal",
 		CASE  T1."VatGroup" 
        WHEN 'O1' THEN 'S'
        WHEN 'O2' THEN 'O'
        WHEN 'X0' THEN 'E'
        WHEN 'A1' THEN 'A'
        ELSE 'Z'       END                    	   	 AS "vat_category_code",
        ''                                           AS "vat_exemption_code",
        ''                                           AS "vat_exemption_reason",
        T1."VatPrcnt"                                AS "vat_rate",
        
        T1."ItemCode"                                AS "item_id",
        T1."Dscription"                              AS "item_name",
        T1."DiscPrcnt"								 AS "item_price_discount", 
          'PERCENTAGE'                               AS "price_allowance_indicator",
        T1."Quantity"                                AS "quantity",
        T1."Price"                                   AS "unit_rate",        
        T1."UomCode"                                 AS "unit_of_measure",        
        ''                                           AS "service_code",
       CASE 
            WHEN IFNULL(T1."DiscPrcnt",0) > 0 THEN 'true'
            ELSE 'false'
        END  								 AS "is_discount_applicable",
		      
		      --billto
        T0."CardName"                                AS "name",		
        C."Cellular"                                 AS "phone",
        CR."Name"                                    AS "country",
        T12."CountryB"                               AS "country_code",
        C."E_Mail"                                   AS "email",
        
        
        T0."Comments"                                AS "notes",
          ''                                         AS "payment_means",
           0.0										 AS "header_discount",
        PT."PymntGroup"                              AS "payment_terms",
           T0."DocCur"                               AS "invoice_currency",
          'B2C'                                      AS "business_kind"
    FROM OINV T0
    INNER JOIN INV1  T1  ON T0."DocEntry" = T1."DocEntry"
    INNER JOIN INV12 T12 ON T0."DocEntry" = T12."DocEntry"
    INNER JOIN NNM1  S   ON T0."Series"   = S."Series"
    INNER JOIN OCRD  C   ON C."CardCode"  = T0."CardCode"
    LEFT  JOIN OCRY  CR  ON CR."Code"     = T12."CountryB"
    LEFT  JOIN OCTG  PT  ON PT."GroupNum" = T0."GroupNum"

    WHERE
        T0."DocEntry" = :DocEntry
        AND T0."CreateDate" > '2025-03-01';

END;