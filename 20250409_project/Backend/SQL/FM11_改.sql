Old


PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_AC5411B.DataID,
    OBU_AC5411B.DataMonthString,
    OBU_AC5411B.AccountCode, 
    ACCAccount.AccountTitle, 
    OBU_AC5411B.MonthAmount
FROM 
    OBU_AC5411B 
INNER JOIN 
    ACCAccount 
ON 
    OBU_AC5411B.AccountCode = ACCAccount.AccountCode
WHERE 
    OBU_AC5411B.AccountCode IN
	    ("425020203", "425020211", "425020212", "425020229",
	     "410331203", "410331211", "410331212", "410331229", 
	     "410332203", "410332211", "410332212", "410332229",
	     	     
	     "431110105", "431110125", "431110127", "431110147",
	     "435030105", "435030125", "435030127",
	     
	     "525030205", "525030225", "525030227",
	     "531110105", "531110125", "531110127", "531110147",
	     
	     "410997201",
	     "425050205", "425050225", "425050227",
	     "450110105", "450110125", "450110127", "450110143",
	     "450130105", "450130125", "450130127", "450130147",
	     "525050205", "525050225", "525050227",
	     "550110105", "550110125", "550110127", "550110143",
	     "550130105", "550130125", "550130147")
	    'FVPL債務工具息
	    'FVOCI債務工具息
	    'AC債務工具息
	    'FVOCI處分利益
	    'AC處分利益
	    'FVPL處分損失
	    'FVOCI處分損失
	    '拆放證券公司息 OSU
	    'FVPL金融資產評價利益
	    'FVOCI債務工具減損迴轉利益
	    'AC債務工具減損迴轉利益
	    'FVPL金融資產評價損失
	    'FVOCI債務工具減損損失
	    'AC債務工具減損損失
    AND OBU_AC5411B.DataMonthString = [DataMonthParam];







    

Old


Query FM11_OBU_AC5411B

SELECT
    d.DataID,
    d.DataMonthString,
    d.AccountCode, 
    a.AccountTitle, 
    d.MonthAmount
FROM 
    OBU_AC5411B AS d
INNER JOIN 
    ACCAccount AS a 
ON 
    d.AccountCode = a.AccountCode
WHERE 
    d.AccountCode IN
	    ("425020203", "425020211", "425020212", "425020229"
	     "410331203", "410331211", "410331212", "410331229", 
	     "410332203", "410332211", "410332212", "410332229",	     
	     "431110105", "431110125", "431110127", "431110147",
	     "435030105", "435030125", "435030127",
	     
	     "525030205", "525030225", "525030227",
	     "531110105", "531110125", "531110127", "531110147",
	     
	     "410997201",
	     "425050205", "425050225", "425050227",
	     "450110105", "450110125", "450110127", "450110143",
	     "450130105", "450130125", "450130127", "450130147",
	     "525050205", "525050225", "525050227",
	     "550110105", "550110125", "550110127", "550110143",
	     "550130105", "550130125", "550130147")
	    'FVPL債務工具息
	    'FVOCI債務工具息
	    'AC債務工具息
	    'FVOCI處分利益
	    'AC處分利益
	    'FVPL處分損失
	    'FVOCI處分損失
	    '拆放證券公司息 OSU
	    'FVPL金融資產評價利益
	    'FVOCI債務工具減損迴轉利益
	    'AC債務工具減損迴轉利益
	    'FVPL金融資產評價損失
	    'FVOCI債務工具減損損失
	    'AC債務工具減損損失
    AND d.DataMonthString = "2024/11";
***remark: 其中 550130147 AC債務工具投資減損損失-金融債券

New

Query FM11_OBU_AC5411B

---------------------------------------
PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_AC5411B.DataID,
    OBU_AC5411B.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    OBU_AC5411B.MonthAmount
FROM
    AccountCodeMap
INNER JOIN
    OBU_AC5411B
ON
    AccountCodeMap.AccountCode = OBU_AC5411B.AccountCode
WHERE
    AccountCodeMap.Category IN ("InterestRevenue" , "GainOnDisposal", "LossOnDisposal", "OSU息", "ValuationProfit", "ValuationLoss")
    AND OBU_AC5411B.DataMonthString = [DataMonthParam];


---------------------------------------
PARAMETERS DataMonthParam TEXT;
SELECT
    oa.DataID,
    oa.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    oa.MonthAmount
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC5411B.DataID,
            OBU_AC5411B.AccountCode,
            OBU_AC5411B.DataMonthString,
            OBU_AC5411B.MonthAmount
        FROM
            OBU_AC5411B
        WHERE
            OBU_AC5411B.DataMonthString = [DataMonthParam]
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ("InterestRevenue" , "GainOnDisposal", "LossOnDisposal", "OSU息", "ValuationProfit", "ValuationLoss");
---------------------------------------

New



PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_AC5411B.DataID,
    OBU_AC5411B.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    OBU_AC5411B.MonthAmount
FROM
    AccountCodeMap
INNER JOIN
    OBU_AC5411B
ON
    AccountCodeMap.AccountCode = OBU_AC5411B.AccountCode
WHERE
    AccountCodeMap.Category IN ("InterestRevenue" , "GainOnDisposal", "LossOnDisposal", "OSU息", "ValuationProfit", "ValuationLoss")
    AND OBU_AC5411B.DataMonthString = '2024/11';



---------------------------------------
PARAMETERS DataMonthParam TEXT;
SELECT
    oa.DataID,
    oa.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    oa.MonthAmount
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC5411B.DataID,
            OBU_AC5411B.AccountCode,
            OBU_AC5411B.DataMonthString,
            OBU_AC5411B.MonthAmount
        FROM
            OBU_AC5411B
        WHERE
            OBU_AC5411B.DataMonthString = '2024/11'
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ("InterestRevenue" , "GainOnDisposal", "LossOnDisposal", "OSU息", "ValuationProfit", "ValuationLoss");



Query FM11_OBU_AC5411B_Subtotal

---------------------------------------
PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.Category, 
    SUM(OBU_AC5411B.MonthAmount) AS SubtotalBalance
FROM
    AccountCodeMap
INNER JOIN
    OBU_AC5411B
ON
    AccountCodeMap.AccountCode = OBU_AC5411B.AccountCode
WHERE
    AccountCodeMap.Category IN ("InterestRevenue" , "GainOnDisposal", "LossOnDisposal", "OSU息", "ValuationProfit", "ValuationLoss")
    AND OBU_AC5411B.DataMonthString = [DataMonthParam]
GROUP BY
    AccountCodeMap.Category





---------------------------------------
PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.Category, 
    SUM(oa.MonthAmount) AS SubtotalBalance
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC5411B.AccountCode,
            OBU_AC5411B.MonthAmount
        FROM
            OBU_AC5411B
        WHERE
            OBU_AC5411B.DataMonthString = [DataMonthParam]
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ("InterestRevenue" , "GainOnDisposal", "LossOnDisposal", "OSU息", "ValuationProfit", "ValuationLoss")
GROUP BY
    AccountCodeMap.Category;
