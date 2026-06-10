CREATE PROCEDURE SBO_SP_TransactionNotification
(
    IN object_type NVARCHAR(30),                -- SBO Object Type
    IN transaction_type NCHAR(1),               -- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
    IN num_of_cols_in_key INT,
    IN list_of_key_cols_tab_del NVARCHAR(255),
    IN list_of_cols_val_tab_del NVARCHAR(255)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
	CountNum integer :=0;
	counter integer := 0 ;
counter1 integer := 0 ;
counter2 integer := 0 ;
    error INT;                                -- Result (0 for no error)
    error_message NVARCHAR(200);              -- Error string to be displayed

    -- Declare the counter variable

BEGIN
    error := 0;
    error_message := N'Ok';
  ---------------------------------------------
    IF :object_type = '10000044' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
    BEGIN
        counter := 0;
        SELECT COUNT(*) INTO counter
        FROM "OBTN" WHERE "SysNumber" = :list_of_cols_val_tab_del AND "MnfDate" IS NULL;     
        -- If MnfDate are missing
        IF :counter > 0 THEN
            error := -309001;
            error_message := 'Manufacture Date is missing!';
        END IF;

    END;
    END IF;
--------------------------------------------------------------


-----------GRPO---------------

/*IF :object_type = '20' AND :transaction_type = 'A'OR :transaction_type = 'U' THEN 
    IF EXISTS (
        SELECT 1 FROM OPDN WHERE "DocEntry" = :list_of_cols_val_tab_del AND "DocTotal" > 5000 AND "BaseType" = -1  -- No linked PO
    ) THEN
        error := 1;
        error_message := 'GRPO over 5000 is not allowed without a PO....!';
    END IF;
END IF;*/

-------------------

------------------------------------------ARINVOICE--------------------------------------------------

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

    
    


    SELECT :error, :error_message FROM dummy;

END;