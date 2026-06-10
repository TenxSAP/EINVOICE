CREATE PROCEDURE C_CORP_EXCEPTIONS
(
    IN FromDate DATE,
    IN ToDate   DATE
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    /* Corporate Tax Exceptions */

    SELECT
        H."TransId" AS "TransactionNo",

        H."RefDate",

        J."Account",

        A."AcctName",

        J."LineMemo",

        J."Debit",

        J."Credit",

        CASE
            WHEN LOWER(A."AcctName")
                LIKE '%entertainment%'

            THEN 'Possibly Non-Deductible'

            WHEN LOWER(A."AcctName")
                LIKE '%penalty%'

            THEN 'Penalty / Fine Review'

            WHEN LOWER(A."AcctName")
                LIKE '%fine%'

            THEN 'Penalty / Fine Review'

            WHEN LOWER(A."AcctName")
                LIKE '%provision%'

            THEN 'Provision Not Tax Deductible'

            WHEN LOWER(A."AcctName")
                LIKE '%related party%'

            THEN 'Transfer Pricing Support Missing'

            ELSE 'Review Required'
        END AS "Issue",

        CASE
            WHEN LOWER(A."AcctName")
                    LIKE '%penalty%'

                 OR LOWER(A."AcctName")
                    LIKE '%fine%'

            THEN 'High'

            WHEN LOWER(A."AcctName")
                    LIKE '%provision%'

            THEN 'High'

            WHEN LOWER(A."AcctName")
                    LIKE '%related party%'

            THEN 'High'

            ELSE 'Medium'
        END AS "Risk"

    FROM OJDT H

    INNER JOIN JDT1 J
        ON H."TransId" = J."TransId"

    INNER JOIN OACT A
        ON J."Account" = A."AcctCode"

    WHERE H."RefDate"
            BETWEEN :FromDate AND :ToDate

      AND A."ActType" = 'E'

      AND
      (
            LOWER(A."AcctName")
                LIKE '%entertainment%'

            OR LOWER(A."AcctName")
                LIKE '%penalty%'

            OR LOWER(A."AcctName")
                LIKE '%fine%'

            OR LOWER(A."AcctName")
                LIKE '%provision%'

            OR LOWER(A."AcctName")
                LIKE '%related party%'
      )

    ORDER BY H."RefDate" DESC;

END;
 