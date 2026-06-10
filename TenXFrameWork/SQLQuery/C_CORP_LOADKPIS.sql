--CORPORATE TAX DASHBOARD KPIS DYNAMIC


CREATE PROCEDURE C_CORP_LOADKPIS
(
    IN FromDate DATE,
    IN ToDate DATE
)
LANGUAGE SQLSCRIPT
AS

    ProfitBeforeTax DECIMAL(18,2);
    AddBackAmount DECIMAL(18,2);
    DeductionAmount DECIMAL(18,2);
    TaxableIncome DECIMAL(18,2);
    Adjustments DECIMAL(18,2);
    EstimatedCorporateTax DECIMAL(18,2);

BEGIN

    /* PROFIT CALCULATION */
    SELECT
        IFNULL(SUM(
            CASE
                WHEN A."ActType" = 'I' THEN J."Credit" - J."Debit"
                WHEN A."ActType" = 'E' THEN J."Debit" - J."Credit"
                ELSE 0
            END
        ),0)
    INTO ProfitBeforeTax
    FROM OJDT H
    INNER JOIN JDT1 J ON H."TransId" = J."TransId"
    INNER JOIN OACT A ON J."Account" = A."AcctCode"
    WHERE H."RefDate" BETWEEN :FromDate AND :ToDate
      AND A."ActType" IN ('I','E');


    /* ADJUSTMENTS */
    SELECT
        IFNULL(SUM(T0."U_CTaxVal"),0),
        IFNULL(SUM(T0."U_CTaxVal"),0)
    INTO AddBackAmount, DeductionAmount
    FROM "@TNX_CTAXCALCU" T0
    WHERE T0."U_DST" = 'A'
      AND T0."U_FDate" BETWEEN :FromDate AND :ToDate;


    TaxableIncome :=
        ProfitBeforeTax + AddBackAmount - DeductionAmount;

    Adjustments :=
        AddBackAmount - DeductionAmount;

    IF TaxableIncome > 375000 THEN
        EstimatedCorporateTax :=
            (TaxableIncome - 375000) * 0.09;
    ELSE
        EstimatedCorporateTax := 0;
    END IF;


    /* FINAL KPI OUTPUT (ROW FORMAT) */

    SELECT 'Profit Before Tax' AS "Title",
           TO_NVARCHAR(:ProfitBeforeTax) AS "Value",
           'Income - Expense' AS "Sub"
    FROM DUMMY

    UNION ALL

    SELECT 'Taxable Income',
           TO_NVARCHAR(:TaxableIncome),
           'After Adjustments'
    FROM DUMMY

    UNION ALL

    SELECT 'Adjustments',
           TO_NVARCHAR(:Adjustments),
           'AddBack - Deduction'
    FROM DUMMY

    UNION ALL

    SELECT 'Estimated Corporate Tax',
           TO_NVARCHAR(:EstimatedCorporateTax),
           '9% Above 375K'
    FROM DUMMY;

END;