
--load kpis Tiles DYNAMIC
CREATE PROCEDURE C_VAT_LOADKPI
(
    IN FromDate DATE,
    IN ToDate DATE
)
LANGUAGE SQLSCRIPT
AS

    OutputVAT DECIMAL(18,2);
    InputVAT DECIMAL(18,2);
    NetVATPayable DECIMAL(18,2);

BEGIN

    SELECT
        COALESCE(SUM("OutputVAT"), 0),
        COALESCE(SUM("InputVAT"), 0),
        COALESCE(SUM("OutputVAT"), 0)
        - COALESCE(SUM("InputVAT"), 0)
    INTO
        OutputVAT,
        InputVAT,
        NetVATPayable
    FROM
    (
        SELECT SUM(T1."VatSum") AS "OutputVAT", 0 AS "InputVAT"
        FROM OINV T0
        INNER JOIN INV1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
          AND T0."DocDate" BETWEEN :FromDate AND :ToDate

        UNION ALL

        SELECT SUM(T1."VatSum" * -1), 0
        FROM ORIN T0
        INNER JOIN RIN1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
          AND T0."DocDate" BETWEEN :FromDate AND :ToDate

        UNION ALL

        SELECT 0, SUM(T1."VatSum")
        FROM OPCH T0
        INNER JOIN PCH1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
          AND T0."DocDate" BETWEEN :FromDate AND :ToDate

        UNION ALL

        SELECT 0, SUM(T1."VatSum" * -1)
        FROM ORPC T0
        INNER JOIN RPC1 T1
            ON T0."DocEntry" = T1."DocEntry"
        WHERE T0."CANCELED" = 'N'
          AND T0."DocDate" BETWEEN :FromDate AND :ToDate
    ) VAT_DATA;

    SELECT
        'Output VAT' AS "Title",
        TO_NVARCHAR(:OutputVAT) AS "Value",
        'Sales VAT' AS "Sub"
    FROM DUMMY

    UNION ALL

    SELECT
        'Input VAT',
        TO_NVARCHAR(:InputVAT),
        'Purchase VAT'
    FROM DUMMY

    UNION ALL

    SELECT
        'Net VAT Payable',
        TO_NVARCHAR(:NetVATPayable),
        'VAT Liability'
    FROM DUMMY;

END;

 