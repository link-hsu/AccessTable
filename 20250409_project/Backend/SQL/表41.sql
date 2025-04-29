Query 表41_DBU_DL9360_LIST

SELECT DBU_DL9360.DataID,
       DBU_DL9360.DataMonthString,
       DBU_DL9360.Trader,
       DBU_DL9360.SettlementDate,
       DBU_DL9360.DealID,
       DBU_DL9360.Counterparty,
       DBU_DL9360.ProfitLossAmount
FROM DBU_DL9360
WHERE
    DBU_DL9360.DataMonthString = "2024/11"
    AND Mid(DBU_DL9360.Counterparty, Len(DBU_DL9360.Counterparty)-3, 2) <> "TW"
Order by DBU_DL9360.ProfitLossAmount DESC



PARAMETERS DataMonthParam TEXT;
SELECT DBU_DL9360.DataID,
       DBU_DL9360.DataMonthString,
       DBU_DL9360.Trader,
       DBU_DL9360.SettlementDate,
       DBU_DL9360.DealID,
       DBU_DL9360.Counterparty,
       DBU_DL9360.ProfitLossAmount
FROM DBU_DL9360
WHERE
    DBU_DL9360.DataMonthString = [DataMonthParam]
    AND Mid(DBU_DL9360.Counterparty, Len(DBU_DL9360.Counterparty)-3, 2) <> "TW"
Order by DBU_DL9360.ProfitLossAmount DESC



New 轉換為美元

Query 表41_DBU_DL9360_Subtotal

SELECT 
    DBU_DL9360.Trader,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] > 0, [DBU_DL9360].[ProfitLossAmount], 0)) AS SumProfit,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] < 0, [DBU_DL9360].[ProfitLossAmount], 0)) AS SumLoss,
    clsRate.Rate AS clsRate_USD,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] > 0, [DBU_DL9360].[ProfitLossAmount], 0)) 
        / clsRate.Rate AS SumProfit_USD,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] < 0, [DBU_DL9360].[ProfitLossAmount], 0)) 
        / clsRate.Rate AS SumLoss_USD
FROM 
    DBU_DL9360, 
    表2_CloseRate_USDTWD AS clsRate 
WHERE 
    DBU_DL9360.DataMonthString = "2024/11"
    AND Mid(DBU_DL9360.Counterparty, Len(DBU_DL9360.Counterparty)-3, 2) <> "TW"
GROUP BY 
    DBU_DL9360.Trader, 
    clsRate.Rate;    
    
    
    
PARAMETERS DataMonthParam TEXT;
SELECT 
    DBU_DL9360.Trader,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] > 0, [DBU_DL9360].[ProfitLossAmount], 0)) AS SumProfit,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] < 0, [DBU_DL9360].[ProfitLossAmount], 0)) AS SumLoss,
    clsRate.Rate AS clsRate_USD,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] > 0, [DBU_DL9360].[ProfitLossAmount], 0)) 
        / clsRate.Rate AS SumProfit_USD,
    Sum(IIf([DBU_DL9360].[ProfitLossAmount] < 0, [DBU_DL9360].[ProfitLossAmount], 0)) 
        / clsRate.Rate AS SumLoss_USD
FROM 
    DBU_DL9360, 
    表2_CloseRate_USDTWD AS clsRate 
WHERE 
    DBU_DL9360.DataMonthString = [DataMonthParam]
    AND Mid(DBU_DL9360.Counterparty, Len(DBU_DL9360.Counterparty)-3, 2) <> "TW"
GROUP BY 
    DBU_DL9360.Trader, 
    clsRate.Rate;
