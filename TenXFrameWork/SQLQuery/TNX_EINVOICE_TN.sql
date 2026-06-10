CREATE PROCEDURE TNX_EINVOICE_TN
(
	in object_type nvarchar(30), 				-- SBO Object Type
	in transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
	in num_of_cols_in_key int,
	in list_of_key_cols_tab_del nvarchar(255),
	in list_of_cols_val_tab_del nvarchar(255),
	out error int,
	out error_message nvarchar (200))
LANGUAGE SQLSCRIPT
AS

Cnt1 int;
Line int;

    -- Declare the counter variable
    counter INT;
    counter1 INT;
    counter2 INT;
    countNum INT;
    
begin

error := 0;
error_message := N'Ok';




--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE
--------------	ADD	YOUR	CODE	HERE
------------------------------------------AR INVOICE--------------------------------------------------

 IF :object_type = '13' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         FROM "OINV" WHERE "DocEntry" = :list_of_cols_val_tab_del AND IFNULL("DiscSum",0)>0;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Document discount should be 0, you cannot use document level discount';
         END IF;
 
     END;
    END IF;
    --mobile no should be less than 15 char
    IF :object_type = '13' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
  BEGIN
      counter := 0;
  
      SELECT COUNT(*) INTO counter
      FROM OINV T0
      INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode"
      WHERE T0."DocEntry" = :list_of_cols_val_tab_del
        AND (
              T1."Cellular" IS NULL
              OR LENGTH(TRIM(T1."Cellular")) = 0
              OR LENGTH(T1."Cellular") > 15
            );
  
      IF :counter > 0 THEN
          error := -133;
          error_message := 'Phone number is mandatory and must not be more than 15 characters long';
      END IF;
  END;
 END IF;
 
   


     --country
IF :object_type = '13' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
BEGIN
    counter := 0;

    SELECT COUNT(*) INTO counter
    FROM OINV  T0
    INNER JOIN INV12 T12 ON T0."DocEntry" = T12."DocEntry"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del
      AND (T12."CountryB" IS NULL OR LENGTH(TRIM(T12."CountryB")) = 0);

    IF :counter > 0 THEN
        error := -133;
        error_message := 'Country is mandatory. Please maintain the Country in the Bill-To Address.';
    END IF;
END;
END IF;

--COUNTRY VALIDATION--
  IF :object_type = '13' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         
FROM "OINV"  T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN CRD1 T2 ON T1."CardCode" = T2."CardCode"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T2."Country" is null;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Please Enter The City Of Customer In Business Partner Master Data';
         END IF;
 
     END;
    END IF;



  IF :object_type = '13' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         
FROM "OINV"  T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN CRD1 T2 ON T1."CardCode" = T2."CardCode"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T2."City" is null;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Please Enter The City Of Customer In Business Partner Master Data';
         END IF;
 
     END;
    END IF;
    
    
    
    
 IF :object_type = '13' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         
FROM "OINV"  T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN CRD1 T2 ON T1."CardCode" = T2."CardCode"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T1."Cellular" is null;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Please Enter The Moblie Of Customer In Business Partner Master Data';
         END IF;
         -- Postal Code (max 10)
    SELECT COUNT(*) INTO counter
    
    FROM "OINV" T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN "CRD1" T2 ON T1."CardCode" = T2."CardCode" 
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T1."County" is null
      AND (T2."ZipCode" IS NULL OR LENGTH(TRIM(T2."ZipCode")) = 0 OR LENGTH(T2."ZipCode") > 127);
      

    IF :counter > 0 THEN
        error := -133;
        error_message := 'Postal Code is mandatory and must not be more than 10 characters long,(ZipCode) in Business Master';
    END IF;

    --------------------------------------------------
    -- District (max 127)
    SELECT COUNT(*) INTO counter
    FROM "OINV" T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN "CRD1" T2 ON T1."CardCode" = T2."CardCode" 
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T1."County" is null
      AND (T2."County" IS NULL OR LENGTH(TRIM(T2."County")) = 0 OR LENGTH(T2."County") > 127);

    IF :counter > 0 THEN
        error := -133;
        error_message := 'District is mandatory and must not be more than 127 characters long(County)in Business Master';
    END IF;
         
 
     END;
    END IF;
    
    
    
 IF :object_type = '13' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         
FROM "OINV"  T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN CRD1 T2 ON T1."CardCode" = T2."CardCode"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T1."CardName" is null;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Please Enter The Name Of Customer In Business Partner Master Data';
         END IF;
         
 
     END;
    END IF;
    
    
     IF :object_type = '13' AND (:transaction_type = 'A' or :transaction_type = 'C') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         FROM "OINV" WHERE "DocEntry" = :list_of_cols_val_tab_del AND "CANCELED"='C';    
         -- If MnfDate are missing
         IF :counter > 0 THEN
             error := -132;
             error_message := 'Cannot cancel the document';
         END IF; SELECT COUNT(*) INTO counter
         FROM "OCRD" WHERE "CardCode" = :list_of_cols_val_tab_del AND ("ZipCode" IS NULL OR "Building" IS NULL or "Cellular" IS NULL);    
        
 
     END;
    END IF;
    
    /*
     IF :object_type = '13' AND (:transaction_type = 'A' or :transaction_type = 'C') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         FROM "OINV" WHERE "DocEntry" = :list_of_cols_val_tab_del AND "U_Type" NOT IN ('Credit','Cash');    
         -- If MnfDate are missing
         IF :counter > 0 THEN
             error := -132;
             error_message := 'Please Select The Invoice Type';
         END IF; SELECT COUNT(*) INTO counter
         FROM "OCRD" WHERE "CardCode" = :list_of_cols_val_tab_del AND ("ZipCode" IS NULL OR "Building" IS NULL or "Cellular" IS NULL);    
        
 
     END;
    END IF;
    */
    IF :object_type = '13'  AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    SELECT COUNT(*) INTO counter
    FROM "OINV" T0
    INNER JOIN "OCRD" T1 ON T0."CardCode" = T1."CardCode"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del
    AND (
          T1."Cellular" IS NULL
          OR LENGTH(TRIM(T1."Cellular")) = 0
          OR (LEFT(T1."Cellular",1) NOT IN ('0','+'))
          OR LENGTH(T1."Cellular") < 5
          OR LENGTH(T1."Cellular") > 16
          OR REPLACE_REGEXPR('[^0-9]' IN SUBSTRING(T1."Cellular",2) WITH '') 
             <> SUBSTRING(T1."Cellular",2)
        );

    IF :counter > 0 THEN
        error := -133;
        error_message := 
        'Mobile number must start with 0 or + and contain 4–15 digits.';
    END IF;

END IF;

IF :object_type = '13'
   AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    SELECT COUNT(*) INTO counter
    FROM "ADM1"
    WHERE  (
            "Street" IS NULL
            OR LENGTH(TRIM("Street")) = 0
            OR LENGTH("Street") > 127
          );

    IF :counter > 0 THEN
        error := -133;
        error_message := 
        'Seller Street (BT-35) must be between 1 and 127 characters.';
    END IF;

END IF;

/*IF :object_type = '13' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
BEGIN

    counter := 0;   

    -- 🔹 2. Payment Means Mandatory Validation
    SELECT COUNT(*) INTO counter
    FROM "OINV"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND ("U_PAYMEANS" IS NULL OR "U_PAYMEANS" = '');

    IF :counter > 0 THEN
        error := -133;
        error_message := 'Payment Means is Mandatory. Please Select Payment Means in UDF.';
    END IF;

END;
END IF;*/
    
    
    
    
    --------------inv------

IF (:object_type = '13' And (:transaction_type = 'A'  OR :transaction_type = 'U' )) then
 BEGIN
 CountNum :=0;
 counter :=0;
 counter1 :=0;
 counter2 :=0;
 
SELECT COUNT(*) INTO CountNum
FROM OINV T0
INNER JOIN INV1 T1 ON T0."DocEntry" = T1."DocEntry"
LEFT JOIN DLN1 T3 ON T1."BaseEntry" = T3."DocEntry"  AND T1."BaseLine" = T3."LineNum" AND T1."BaseType" = '15'
LEFT JOIN ODLN T2 ON T3."DocEntry" = T2."DocEntry"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del
AND T0."isIns" = 'N' AND T0."DocType" = 'I'AND T0."CANCELED" = 'N'
AND T1."BaseType" <> '15';  

IF (CountNum > 0) THEN 
    error := -1304;
    error_message := 'AR Invoice must be based on a Delivery Note'; 
END IF;
END;
END IF;




IF (:object_type = '13' AND (:transaction_type = 'A' OR :transaction_type = 'U')) THEN
BEGIN
  DECLARE CountNum INT;

  SELECT COUNT(*) INTO CountNum
  FROM OINV T0
  INNER JOIN INV1 T1 ON T0."DocEntry" = T1."DocEntry"
  LEFT JOIN DLN1 T3 ON T1."BaseEntry" = T3."DocEntry" 
                   AND T1."BaseLine" = T3."LineNum"
                   AND T1."BaseType" = 15
  LEFT JOIN ODLN T2 ON T3."DocEntry" = T2."DocEntry"
  WHERE T0."DocEntry" = :list_of_cols_val_tab_del
    AND T0."isIns" = 'N'
    AND T0."DocType" = 'I'
    AND T1."BaseType" = 15 
    AND COALESCE(T1."Price", 0) <> COALESCE(T3."Price", 0);

  IF (CountNum > 0) THEN
    error := -1304;
    error_message := 'Unit Price must match the Delivery Note price.';
  END IF;
END;
END IF;






       
------------------------------------------CREDITMEMO-----------------------------------------------------

--

   IF :object_type = '14'  AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    SELECT COUNT(*) INTO counter
    FROM "ORIN" T0
    INNER JOIN "OCRD" T1 ON T0."CardCode" = T1."CardCode"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del
    AND (
          T1."Cellular" IS NULL
          OR LENGTH(TRIM(T1."Cellular")) = 0
          OR (LEFT(T1."Cellular",1) NOT IN ('0','+'))
          OR LENGTH(T1."Cellular") < 5
          OR LENGTH(T1."Cellular") > 16
          OR REPLACE_REGEXPR('[^0-9]' IN SUBSTRING(T1."Cellular",2) WITH '') 
             <> SUBSTRING(T1."Cellular",2)
        );

    IF :counter > 0 THEN
        error := -133;
        error_message := 
        'Customer Mobile number must start with 0 or + and contain 4–15 digits In Business Partner Master Data.';
    END IF;

END IF;
/*
--Payment Means Field
IF :object_type = '14' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
BEGIN

    counter := 0;   

    -- 🔹 2. Payment Means Mandatory Validation
    SELECT COUNT(*) INTO counter
    FROM "ORIN"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND ("U_PAYMEANS" IS NULL OR "U_PAYMEANS" = '');

    IF :counter > 0 THEN
        error := -133;
        error_message := 'Payment Means is Mandatory. Please Select Payment Means in UDF.';
    END IF;

END;
END IF;*/
    
--Street field should not be empty in Company Details
IF :object_type = '14' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    SELECT COUNT(*) INTO counter
    FROM "ADM1"
    WHERE  (
            "Street" IS NULL
            OR LENGTH(TRIM("Street")) = 0
            OR LENGTH("Street") > 127
          );

    IF :counter > 0 THEN
        error := -133;
        error_message := 
        'Seller Street (BT-35) must be between 1 and 127 characters.';
    END IF;

END IF;

 IF :object_type = '14' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         FROM "ORIN" T0 WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND (IFNULL(T0."U_REASON",'')='' or IFNULL(T0."U_REASON",'')='-') AND T0. "DocType" = 'I';    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Credit note reason is mandatory.';
         END IF;
     END;
    END IF;
    IF :object_type = '14' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         FROM "ORIN" T0 INNER JOIN "RIN1" T1 ON T0."DocEntry"=T1."DocEntry" 
         WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND (IFNULL(T1."BaseRef",'')='' AND IFNULL(T0."U_INVNO",'')='') AND T0. "DocType" = 'I';    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Credit note base document number is mandatory';
         END IF;
     END;
    END IF; 
    
    
     IF :object_type = '14' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         FROM "ORIN" WHERE "DocEntry" = :list_of_cols_val_tab_del AND IFNULL("DiscSum",0)>0;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Document discount should be 0, you cannot use document level discount';
         END IF;
 
     END;
    END IF;




   IF :object_type = '14' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         
FROM "ORIN"  T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN CRD1 T2 ON T1."CardCode" = T2."CardCode"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T2."City" is null;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Please Enter The City Of Customer In Business Partner Master Data';
         END IF;
 
     END;
    END IF;
    
    
    
    
 IF :object_type = '14' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         
FROM "ORIN"  T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN CRD1 T2 ON T1."CardCode" = T2."CardCode"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T1."Cellular" is null;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Please Enter The Moblie Of Customer In Business Partner Master Data';
         END IF;
 
     END;
    END IF;
    
    
 IF :object_type = '14' AND (:transaction_type = 'A' or :transaction_type = 'U') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         
FROM "ORIN"  T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN CRD1 T2 ON T1."CardCode" = T2."CardCode"
WHERE T0."DocEntry" = :list_of_cols_val_tab_del AND T1."CardName" is null;    
         IF :counter > 0 THEN
             error := -133;
             error_message := 'Please Enter The Name Of Customer In Business Partner Master Data';
         END IF;
 
     END;
    END IF;
    
    
     IF :object_type = '14' AND (:transaction_type = 'A' or :transaction_type = 'C') THEN
     BEGIN
         counter := 0;
         SELECT COUNT(*) INTO counter
         FROM "ORIN" WHERE "DocEntry" = :list_of_cols_val_tab_del AND "CANCELED"='C';    
         -- If MnfDate are missing
         IF :counter > 0 THEN
             error := -132;
             error_message := 'Cannot cancel the document';
         END IF; SELECT COUNT(*) INTO counter
         FROM "OCRD" WHERE "CardCode" = :list_of_cols_val_tab_del AND ("ZipCode" IS NULL OR "Building" IS NULL or "Cellular" IS NULL);    
        
 
     END;
    END IF;
     

End;