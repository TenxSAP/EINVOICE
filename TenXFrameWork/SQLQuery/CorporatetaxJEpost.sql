SELECT T0."DocEntry", T0."U_AIPD",
    T0."U_DAT",
    T0."U_FDate",
    T0."U_TDate",
    T0."U_PPeriod",
    T0."U_CTax",
    T0."U_CTaxVal",
    T0."U_JEPOSTD",
    T0."U_JENo",
    T0."U_Status" FROM "@TNX_CTAXCALCU" T0 WHERE T0."U_DST" = 'O'
    AND IFNULL(T0."U_JENo", '') = '' And T0."UpdateDate"=CURRENT_DATE and LEFT(LPAD(T0."UpdateTime",6,'0'),4)>=TO_CHAR(ADD_SECONDS(CURRENT_TIMESTAMP,-60), 'HH24MI')