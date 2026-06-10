SELECT T0."DocEntry",T0."DocNum", T0."DocDate", T0."CardCode", 
T0."CardName", T0."VatSum",T0."DocTotal",T0."DocCur",T0."DocStatus",T0."ObjType" FROM ORDR T0 WHERE 
T0."UpdateDate"=CURRENT_DATE and LEFT(LPAD(T0."UpdateTS",6,'0'),4)>=TO_CHAR(ADD_SECONDS(CURRENT_TIMESTAMP,-60), 'HH24MI')