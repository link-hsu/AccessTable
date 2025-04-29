
New version 基準日 最後一個工作日
Query AI240_DBU_DL6850_LIST

SELECT DBU_DL6850.DataID,
			 DBU_DL6850.DataMonthString,
			 DBU_DL6850.DealID,
			 DBU_DL6850.SettlementDate,
			 DBU_DL6850.ContractDate,
       DBU_DL6850.BuyCurrency,
       DBU_DL6850.BuyAmount,
       DBU_DL6850.SellAmount,
       DBU_DL6850.SellCurrency
From DBU_DL6850
WHERE DBU_DL6850.DataMonthString = "2024/11"
			AND ContractDate <= DBU_DL6850.DataDate
      AND SettlementDate > DBU_DL6850.DataDate
      
      
PARAMETERS DataMonthParam TEXT;         
SELECT DBU_DL6850.DataID,
			 DBU_DL6850.DataMonthString,
			 DBU_DL6850.DealID,
			 DBU_DL6850.SettlementDate,
			 DBU_DL6850.ContractDate,
       DBU_DL6850.BuyCurrency,
       DBU_DL6850.BuyAmount,
       DBU_DL6850.SellAmount,
       DBU_DL6850.SellCurrency
From DBU_DL6850
WHERE DBU_DL6850.DataMonthString = [DataMonthParam]
			AND ContractDate <= DBU_DL6850.DataDate
      AND SettlementDate > DBU_DL6850.DataDate
      
      
      
      
      
      
 New version - 基準日為月底工作日
Query AI240_DBU_DL6850_Subtoal

SELECT
    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 0 AND 10, '基準日後0-10天',
        IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 11 AND 30, '基準日後11-30天',
            IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 31 AND 90, '基準日後31-90天',
                IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 91 AND 180, '基準日後91-180天',
                    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 181 AND 365, '基準日後181天-1年', '超過基準日後一年')
                )
            )
        )
    ) AS DaysRange,
    Sum(IIf(DBU_DL6850.BuyCurrency = 'TWD', DBU_DL6850.BuyAmount, 0)) AS SumBuyAmountTWD,
    Sum(IIf(DBU_DL6850.SellCurrency = 'TWD', DBU_DL6850.SellAmount, 0)) AS SumSellAmountTWD
FROM 
    DBU_DL6850
WHERE 
    DBU_DL6850.DataMonthString = "2024/11"
    AND DBU_DL6850.ContractDate <= DBU_DL6850.DataDate
    AND DBU_DL6850.SettlementDate > DBU_DL6850.DataDate
GROUP BY 
    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 0 AND 10, '基準日後0-10天',
        IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 11 AND 30, '基準日後11-30天',
            IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 31 AND 90, '基準日後31-90天',
                IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 91 AND 180, '基準日後91-180天',
                    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 181 AND 365, '基準日後181天-1年', '超過基準日後一年')
                )
            )
        )
    );
    
    
PARAMETERS DataMonthParam TEXT;    
SELECT 
    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 0 AND 10, '基準日後0-10天',
        IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 11 AND 30, '基準日後11-30天',
            IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 31 AND 90, '基準日後31-90天',
                IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 91 AND 180, '基準日後91-180天',
                    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 181 AND 365, '基準日後181天-1年', '超過基準日後一年')
                )
            )
        )
    ) AS DaysRange,
    Sum(IIf(DBU_DL6850.BuyCurrency = 'TWD', DBU_DL6850.BuyAmount, 0)) AS SumBuyAmountTWD,
    Sum(IIf(DBU_DL6850.SellCurrency = 'TWD', DBU_DL6850.SellAmount, 0)) AS SumSellAmountTWD
FROM 
    DBU_DL6850
WHERE
    DBU_DL6850.DataMonthString = [DataMonthParam]
    AND DBU_DL6850.ContractDate <= DBU_DL6850.DataDate
    AND DBU_DL6850.SettlementDate > DBU_DL6850.DataDate
GROUP BY 
    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 0 AND 10, '基準日後0-10天',
        IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 11 AND 30, '基準日後11-30天',
            IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 31 AND 90, '基準日後31-90天',
                IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 91 AND 180, '基準日後91-180天',
                    IIf(DBU_DL6850.SettlementDate - DBU_DL6850.DataDate BETWEEN 181 AND 365, '基準日後181天-1年', '超過基準日後一年')
                )
            )
        )
    );


 
