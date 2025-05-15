

Query AI822_CloseRate_TWDCNY

PARAMETERS DataMonthParam TEXT;
SELECT
		CloseRate.BaseCurrency,
		CloseRate.QuoteCurrency,
		CloseRate.Rate
FROM 
    CloseRate
WHERE
		CloseRate.BaseCurrency = "TWD"
		AND CloseRate.QuoteCurrency = "CNY"
		AND CloseRate.DataMonthString = [DataMonthParam]
    AND CloseRate.DataDate = (
    SELECT MAX(clsLast.DataDate)
    FROM CloseRate AS clsLast
    WHERE 
        clsLast.DataMonthString = [DataMonthParam]
        AND clsLast.BaseCurrency = "TWD" 
        AND clsLast.QuoteCurrency = "CNY"
);

-- ====================================
Query AI822_DBU_AC5092B_DepositList

PARAMETERS DataMonthParam TEXT;
SELECT
    DBU_AC5092B.AccountID,
    DBU_AC5092B.AccountName,
    DBU_AC5092B.CurrencyType,
    DBU_AC5092B.DebitBalance_Tday
FROM 
    DBU_AC5092B
WHERE
    DBU_AC5092B.DataMonthString = [DataMonthParam]
    AND DBU_AC5092B.AccountID IN ('886', '890', '891')
    AND DBU_AC5092B.CurrencyType = 'CNY';


-- ===================================

Query AI822_OBU_DBU_MM4901B_LendingList


PARAMETERS DataMonthParam TEXT;
SELECT
    om.DealDate,
    om.DealId,
    om.CounterParty,
    om.CurrencyType,
    om.Amount,
    e.Rate
FROM 
    OBU_MM4901B AS om
    LEFT JOIN CloseRate AS e 
	    ON om.CurrencyType = e.QuoteCurrency
WHERE
    om.DataMonthString = [DataMonthParam]
    AND om.CounterParty IN ('BKCHTWTP', 'COMMTWTP', 'PCBCTWTP')
    AND e.BaseCurrency = 'TWD'
    AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = [DataMonthParam]
    )

UNION ALL

SELECT
    dm.DealDate,
    dm.DealId,
    dm.CounterParty,
    dm.CurrencyType,
    dm.Amount,
    e.Rate
FROM 
    DBU_MM4901B AS dm
    LEFT JOIN CloseRate AS e 
	    ON dm.CurrencyType = e.QuoteCurrency
WHERE
    dm.DataMonthString = [DataMonthParam]
    AND dm.CounterParty IN ('BKCHTWTP', 'COMMTWTP', 'PCBCTWTP')
    AND e.BaseCurrency = 'TWD'
    AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = [DataMonthParam]
    );



-- ================================

Query AI822_OBU_DBU_MM4901B_LendingTotal

PARAMETERS DataMonthParam TEXT;
SELECT
    'OBU' AS Source,
    SUM(om.Amount * e.Rate) AS Total_TWD_Amount
FROM 
    OBU_MM4901B AS om
    LEFT JOIN CloseRate AS e 
	    ON om.CurrencyType = e.QuoteCurrency
WHERE
    om.DataMonthString = [DataMonthParam]
    AND om.CounterParty IN ('BKCHTWTP', 'COMMTWTP', 'PCBCTWTP')
    AND e.BaseCurrency = 'TWD'
    AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = [DataMonthParam]
    )

UNION ALL

SELECT
    'DBU' AS Source,
    SUM(dm.Amount * e.Rate) AS Total_TWD_Amount
FROM 
    DBU_MM4901B AS dm
    LEFT JOIN CloseRate AS e 
	    ON dm.CurrencyType = e.QuoteCurrency
WHERE
    dm.DataMonthString = [DataMonthParam]
    AND dm.CounterParty IN ('BKCHTWTP', 'COMMTWTP', 'PCBCTWTP')
    AND e.BaseCurrency = 'TWD'
    AND e.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE clsLast.DataMonthString = [DataMonthParam]
    );












BKCHTWTP
COMMTWTP
PCBCTWTP
