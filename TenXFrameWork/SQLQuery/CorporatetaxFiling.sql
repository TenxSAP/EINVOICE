SELECT T0."DocEntry",  T0."Object",     T0."U_CMN"  AS "Status",
    T0."Code",
    T0."U_CMR"  AS "Start Period",
    T0."U_SO"   AS "End Period",
    T0."U_GRPO" AS "Alert Before Days",
    T0."U_DCNF" AS "Submission Date",
 
 
    LAST_DAY(
        TO_DATE(
            YEAR(CURRENT_DATE) || '-' || 
            LPAD(T0."U_SO",2,'0') || '-01',
            'YYYY-MM-DD')) AS "End Period Last Date",
 
    ADD_DAYS( LAST_DAY( TO_DATE (YEAR(CURRENT_DATE) || '-' || 
                LPAD(T0."U_SO",2,'0') || '-01', 'YYYY-MM-DD')),
        -IFNULL(T0."U_GRPO",0)) AS "Reminder Date"
 
FROM "@TNX_CORPTAX" T0
 
WHERE IFNULL(T0."U_CMN",'') = 'O' AND LPAD(T0."U_SO",2,'0') = LPAD(MONTH(CURRENT_DATE),2,'0')
 
    AND CURRENT_DATE >= ADD_DAYS(LAST_DAY(TO_DATE(YEAR(CURRENT_DATE) || '-' || 
                                        LPAD(T0."U_SO",2,'0') || '-01', 'YYYY-MM-DD')), -IFNULL(T0."U_GRPO",0))
                                          AND CURRENT_DATE = ADD_DAYS(LAST_DAY(TO_DATE(YEAR(CURRENT_DATE) || '-' || 
                                    LPAD(T0."U_SO",2,'0') || '-01', 'YYYY-MM-DD')),
                            -IFNULL(T0."U_GRPO",0))  And T0."UpdateDate"=CURRENT_DATE 
                            and LEFT(LPAD(T0."UpdateTime",6,'0'),4)>=TO_CHAR(ADD_SECONDS(CURRENT_TIMESTAMP,-60), 'HH24MI')