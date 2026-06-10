CREATE PROCEDURE C_CORP_COMPUTATION
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    /* Corporate Tax Computation */

    WITH PROFIT AS
    (
        SELECT
            SUM
            (
                CASE
                    WHEN A."ActType" = 'I'
                        THEN J."Credit" - J."Debit"

                    WHEN A."ActType" = 'E'
                        THEN J."Debit" - J."Credit"

                    ELSE 0
                END
            ) AS "AccountingProfit"

        FROM OJDT H

        INNER JOIN JDT1 J
            ON H."TransId" = J."TransId"

        INNER JOIN OACT A
            ON J."Account" = A."AcctCode"

        WHERE H."RefDate"
                BETWEEN :FromDate AND :ToDate

          AND A."ActType" IN ('I', 'E')
    ),

    ADJ AS
    (
        SELECT
            IFNULL(SUM(T0."U_CTaxVal"), 0)
                AS "NonDeductibleExpenses",

            IFNULL(SUM(T0."U_CTaxVal"), 0)
                AS "OtherDeductibleAdjustments"

        FROM "@TNX_CTAXCALCU" T0

        WHERE T0."U_DST" = 'A'

          AND T0."U_FDate"
                BETWEEN :FromDate AND :ToDate
    )

    SELECT
        'Accounting Profit' AS "Particulars",
        P."AccountingProfit" AS "Amount"

    FROM PROFIT P

    UNION ALL

    SELECT
        'Add: Non-deductible Expenses',
        A."NonDeductibleExpenses"

    FROM ADJ A

    UNION ALL

    SELECT
        'Less: Other Deductible Adjustments',
        A."OtherDeductibleAdjustments" * -1

    FROM ADJ A

    UNION ALL

    SELECT
        'Taxable Income',

        P."AccountingProfit"
        + A."NonDeductibleExpenses"
        - A."OtherDeductibleAdjustments"

    FROM PROFIT P

    CROSS JOIN ADJ A

    UNION ALL

    SELECT
        'Corporate Tax @ 9%',

        CASE
            WHEN
                P."AccountingProfit"
                + A."NonDeductibleExpenses"
                - A."OtherDeductibleAdjustments"
                > 375000

            THEN
            (
                P."AccountingProfit"
                + A."NonDeductibleExpenses"
                - A."OtherDeductibleAdjustments"
                - 375000
            ) * 0.09

            ELSE 0
        END AS "Amount"

    FROM PROFIT P

    CROSS JOIN ADJ A;

END;