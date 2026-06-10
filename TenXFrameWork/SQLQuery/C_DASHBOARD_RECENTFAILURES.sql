--TOP 5 RECENT FAILURES
CREATE PROCEDURE "C_DASHBOARD_RECENTFAILURES"
(

)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT TOP 5
        "DocNum",

        IFNULL("CardName",'') AS "Customer",

        IFNULL("U_STATUS1",'Pending') AS "Status",

        "DocDate"

    FROM OINV

    WHERE "U_STATUS1" = 'FAILED'


    ORDER BY "DocDate" DESC;

END;

 