Query 表2_CloseRate_USDTWD

SELECT
		CloseRate.BaseCurrency,
		CloseRate.QuoteCurrency,
		CloseRate.Rate
FROM 
    CloseRate
WHERE
		CloseRate.BaseCurrency = "USD"
		AND CloseRate.QuoteCurrency = "TWD"
		AND CloseRate.DataMonthString = "2025/03";

PARAMETERS DataMonthParam TEXT;
SELECT
		CloseRate.BaseCurrency,
		CloseRate.QuoteCurrency,
		CloseRate.Rate
FROM 
    CloseRate
WHERE
		CloseRate.BaseCurrency = "USD"
		AND CloseRate.QuoteCurrency = "TWD"
		AND CloseRate.DataMonthString = [DataMonthParam];
