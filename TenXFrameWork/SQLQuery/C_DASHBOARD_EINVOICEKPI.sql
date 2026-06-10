--INVOICE DASHBOARD KPIS DYNAMIC
CREATE PROCEDURE "C_DASHBOARD_EINVOICEKPI"
(

)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT 
        'Total Invoices' AS "Title",
        TO_NVARCHAR(COUNT(*)) AS "Value",
        'Total invoices' AS "Sub"
    FROM OINV

    UNION ALL

    SELECT 
        'Success' AS "Title",
        TO_NVARCHAR(
            SUM(
                CASE 
                    WHEN "U_STATUS1" = 'SUCCESS' 
                    THEN 1 
                    ELSE 0 
                END
            )
        ) AS "Value",
        'Accepted / synced' AS "Sub"
    FROM OINV

    UNION ALL

    SELECT 
        'Failed' AS "Title",
        TO_NVARCHAR(
            SUM(
                CASE 
                    WHEN "U_STATUS1" = 'FAILED' 
                    THEN 1 
                    ELSE 0 
                END
            )
        ) AS "Value",
        'Requires action' AS "Sub"
    FROM OINV

    UNION ALL

    SELECT 
        'Pending' AS "Title",
        TO_NVARCHAR(
            SUM(
                CASE 
                    WHEN "U_STATUS1" IS NULL 
                         OR TRIM("U_STATUS1") = ''
                    THEN 1 
                    ELSE 0 
                END
            )
        ) AS "Value",
        'Not sent to document' AS "Sub"
    FROM OINV
   -- WHERE "DocDate" BETWEEN :FROMDATE AND :TODATE

    UNION ALL

    SELECT 
        'Success Rate' AS "Title",

        TO_NVARCHAR
        (
            ROUND
            (
                CASE 
                    WHEN 
                    (
                        SUM(CASE WHEN "U_STATUS1" = 'SUCCESS' THEN 1 ELSE 0 END) +
                        SUM(CASE WHEN "U_STATUS1" = 'FAILED' THEN 1 ELSE 0 END)
                    ) = 0
                    THEN 0

                    ELSE
                    (
                        SUM(CASE WHEN "U_STATUS1" = 'SUCCESS' THEN 1 ELSE 0 END) * 100.0
                    ) /
                    (
                        SUM(CASE WHEN "U_STATUS1" = 'SUCCESS' THEN 1 ELSE 0 END) +
                        SUM(CASE WHEN "U_STATUS1" = 'FAILED' THEN 1 ELSE 0 END)
                    )
                END,
            2)
        ) AS "Value",

        'Processing efficiency' AS "Sub"

    FROM OINV

    UNION ALL

    SELECT 
        'Avg Response' AS "Title",
        '1.8s' AS "Value",
        'Gateway time' AS "Sub"
    FROM DUMMY;

END;