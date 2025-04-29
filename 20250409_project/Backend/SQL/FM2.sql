增加考慮篩選每月最後一天匯率
new version 
Query FM2_OBU_MM4901B_LIST

SELECT 
  om.DataID,
  om.DataMonthString,
  om.DealId,
  om.CounterParty,
  om.CurrencyType,
  om.Amount,
  e.BaseCurrency,
  e.QuoteCurrency,
  e.Rate
FROM 
  OBU_MM4901B AS om
LEFT JOIN 
  CloseRate AS e
    ON om.CurrencyType = e.QuoteCurrency
WHERE 
  om.DataMonthString = "2024/12"
  AND e.DataDate = (
    SELECT MAX(clsLast.DataDate) 
    FROM CloseRate AS clsLast 
    WHERE clsLast.DataMonthString = "2024/12"
  )
  AND (
    e.BaseCurrency <> "TWD"
    OR (e.BaseCurrency = "TWD" AND e.QuoteCurrency = "USD")
  );
  
  
PARAMETERS DataMonthParam TEXT;
SELECT 
  om.DataID,
  om.DataMonthString,
  om.DealId,
  om.CounterParty,
  om.CurrencyType,
  om.Amount,
  e.BaseCurrency,
  e.QuoteCurrency,
  e.Rate
FROM 
  OBU_MM4901B AS om
LEFT JOIN 
  CloseRate AS e
    ON om.CurrencyType = e.QuoteCurrency
WHERE 
  om.DataMonthString = [DataMonthParam]
  AND e.DataDate = (
    SELECT MAX(clsLast.DataDate) 
    FROM CloseRate AS clsLast 
    WHERE clsLast.DataMonthString = [DataMonthParam]
  )
  AND (
    e.BaseCurrency <> "TWD"
    OR (e.BaseCurrency = "TWD" AND e.QuoteCurrency = "USD")
  );
  





增加考慮篩選每月最後一天匯率
new version
Query FM2_OBU_MM4901B_Subtotal_v2

SELECT 
    om.CounterParty,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    ) AS Category,
    SUM(IIf(om.CurrencyType = 'USD', 
            om.Amount, 
            om.Amount / e.Rate)) AS total_amount_USD
FROM 
    OBU_MM4901B AS om
    LEFT JOIN CloseRate AS e 
    ON om.CurrencyType = e.QuoteCurrency
WHERE 
    om.DataMonthString = '2024/11' 
    AND 
    (
        e.BaseCurrency <> 'TWD' 
        OR (e.BaseCurrency = 'TWD' AND e.QuoteCurrency = 'USD')
    )
		AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = '2024/11'
		)
GROUP BY 
    om.CounterParty,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    );




PARAMETERS DataMonthParam TEXT;    
SELECT 
    om.CounterParty,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    ) AS Category,
    SUM(IIf(om.CurrencyType = 'USD', 
            om.Amount, 
            om.Amount / e.Rate)) AS total_amount_USD
FROM 
    OBU_MM4901B AS om
    LEFT JOIN CloseRate AS e 
    ON om.CurrencyType = e.QuoteCurrency
WHERE 
    om.DataMonthString = [DataMonthParam]
    AND 
    (
        e.BaseCurrency <> 'TWD' 
        OR (e.BaseCurrency = 'TWD' AND e.QuoteCurrency = 'USD')
    )
		AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = [DataMonthParam]
		)    
GROUP BY 
    om.CounterParty,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    );








增加考慮篩選每月最後一天匯率
new version

Query FM2_OBU_MM4901B_Subtotal_BankCode

SELECT 
    om.CounterParty,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    ) AS Category,
    SUM(IIf(om.CurrencyType = 'USD', 
            om.Amount, 
            om.Amount / e.Rate)) AS total_amount_USD,
    bk.BankCode
FROM 
    (OBU_MM4901B AS om
    LEFT JOIN CloseRate AS e 
	    ON om.CurrencyType = e.QuoteCurrency)
    LEFT JOIN BankDirectory AS bk
	    ON om.CounterParty = bk.SWIFT
WHERE 
    om.DataMonthString = '2024/11' 
    AND 
    (
        e.BaseCurrency <> 'TWD' 
        OR (e.BaseCurrency = 'TWD' AND e.QuoteCurrency = 'USD')
    )
		AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = '2024/11'
		)    
GROUP BY 
    om.CounterParty,
    bk.BankCode,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    );
    
    
    
    
PARAMETERS DataMonthParam TEXT;    
SELECT 
    om.CounterParty,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    ) AS Category,
    SUM(IIf(om.CurrencyType = 'USD', 
            om.Amount, 
            om.Amount / e.Rate)) AS total_amount_USD,
    bk.BankCode
FROM 
    (OBU_MM4901B AS om
    LEFT JOIN CloseRate AS e 
	    ON om.CurrencyType = e.QuoteCurrency)
    LEFT JOIN BankDirectory AS bk
	    ON om.CounterParty = bk.SWIFT
WHERE 
    om.DataMonthString = [DataMonthParam]
    AND 
    (
        e.BaseCurrency <> 'TWD' 
        OR (e.BaseCurrency = 'TWD' AND e.QuoteCurrency = 'USD')
    )
		AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = [DataMonthParam]
		)    
GROUP BY 
    om.CounterParty,
    bk.BankCode,
    IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(om.CounterParty, 3) <> "OBU" AND MID(om.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(om.CounterParty, 3) = "OBU" AND MID(om.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    );
