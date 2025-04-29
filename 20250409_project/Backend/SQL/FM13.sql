New 修改 fv.PL_Amt_USD < 0 只針對 PL_Amt_USD
Query FM13_FXDebtEvaluation_Subtotal_FVandAdjust

SELECT 
    ref.Country, 
    SUM(fv.Book_Value) AS Book_Value, 
    SUM(IIf(fv.PL_Amt_USD < 0, fv.PL_Amt_USD, 0)) AS PL_Amt_USD
FROM 
    FXDebtReferCountry AS ref
LEFT JOIN FXDebtEvaluation AS fv 
    ON ref.Issuer = fv.Issuer
WHERE 
    fv.DataMonthString = "2024/11"
GROUP BY 
    ref.Country;



PARAMETERS DataMonthParam TEXT;
SELECT 
    ref.Country, 
    SUM(fv.Book_Value) AS Book_Value, 
    SUM(IIf(fv.PL_Amt_USD < 0, fv.PL_Amt_USD, 0)) AS PL_Amt_USD
FROM 
    FXDebtReferCountry AS ref
LEFT JOIN FXDebtEvaluation AS fv 
    ON ref.Issuer = fv.Issuer
WHERE 
    fv.DataMonthString = [DataMonthParam]
GROUP BY 
    ref.Country;




New Query FM13_FXDebtEvaluation_Subtotal_Impairment
    
SELECT 
    ref.Country, 
    SUM(ip.CurrImpairmentCost) AS CurrImpairmentCost
FROM 
    FXDebtReferCountry AS ref
LEFT JOIN AssetsImpairment AS ip
    ON ref.Issuer = ip.BondName
WHERE 
    ip.DataMonthString = "2024/11"
GROUP BY 
    ref.Country;


PARAMETERS DataMonthParam TEXT;    
SELECT 
    ref.Country, 
    SUM(ip.CurrImpairmentCost) AS CurrImpairmentCost
FROM 
    FXDebtReferCountry AS ref
LEFT JOIN AssetsImpairment AS ip
    ON ref.Issuer = ip.BondName
WHERE 
    ip.DataMonthString = [DataMonthParam]
GROUP BY 
    ref.Country;
