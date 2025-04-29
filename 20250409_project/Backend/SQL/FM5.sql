
Query FM5_OBU_FC9450B


PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_FC9450B.DataID,
    OBU_FC9450B.CurrencyType,
    OBU_FC9450B.DealID,
    OBU_FC9450B.BookAmount,
    OBU_FC9450B.ProfitLossAmount
FROM
    OBU_FC9450B
WHERE
		OBU_FC9450B.DataMonthString = [DataMonthParam];
	
	
	
