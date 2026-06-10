SELECT 
    T0."DocEntry",
    T0."DocNum",
    T0."CardCode",
    T0."CardName",
    T0."DocDate",
    T0."DocTotal",
    T0."DocStatus",
    T0."EDocStatus",
    CASE 
        WHEN T0."EDocStatus" = 'G' THEN 'SUCCESS'
        WHEN T0."EDocStatus" = 'N' THEN 'FAILED'
        WHEN T0."EDocStatus" = 'P' THEN 'PENDING'
        ELSE 'NOT GENERATED'
    END AS "EInvoice_Status"
 
FROM OINV T0 Where T0."UpdateDate"=CURRENT_DATE and LEFT(LPAD(T0."UpdateTS",6,'0'),4)>=TO_CHAR(ADD_SECONDS(CURRENT_TIMESTAMP,-60), 'HH24MI')
ORDER BY T0."DocNum" DESC;