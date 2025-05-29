1.更新匯率、餘額A、餘額C及表一~三與Table1~3表頭日期
2.更新持有到期債券風險餘額表(A:V):以複製貼上值方式更新(AC)
3.更新外幣債評估表並改分頁名:以複製貼上值方式更新(僅AC部位) 貼A:G欄位
4.將評估表的內容貼到外幣債公允價值評估
5.檢核:確認檢核表檢核欄位是否皆正確



更新匯率
餘額A
更新持有到期債券風險餘額表(A:V):以複製貼上值方式更新(AC)
3.更新外幣債評估表並改分頁名:以複製貼上值方式更新(僅AC部位) 貼A:G欄位
4.將評估表的內容貼到外幣債公允價值評估
5.檢核:確認檢核表檢核欄位是否皆正確


減損表

空白
=K4*VLOOKUP(B4,匯率!B:D,3,0)
欄位K Market_Value
欄位B Ccy

"公允價值評估未實現利益" W欄位
=IF(L5>=0,L5,0)*匯率!$D$12
L欄位 PL_Amt_USD

"公允價值評估未實現損失" X欄位
=IF(L5<0,L5,0)*匯率!$D$12		
		


- 表一 表一：按攤銷後成本法衡量之債務工具
    - 國內投資
        - 按攤銷後成本衡量之債務工具投資
            - 帳列金額
                =VLOOKUP(122,餘額A!A:C,3,0)
            - 公允價值
                = 帳列金額 + 公允價值評估未實現利益 + 公允價值評估未實現損失
            - "公允價值評估未實現利益"
                =SUM(持有到期債券風險餘額表!W:W)
                W欄位
            - "公允價值評估未實現損失"
                =SUM(持有到期債券風險餘額表!X:X)
                X欄位

    - 國外投資
        - 按攤銷後成本衡量之債務工具投資
            - 帳列金額
            =VLOOKUP(122,餘額C!A:C,3,0)-VLOOKUP(122,餘額A!A:C,3,0)
            - 公允價值
                = 帳列金額 + 公允價值評估未實現利益 + 公允價值評估未實現損失
            - "公允價值評估未實現利益"
                =SUM(外幣債公允價值評估!T:T)
                W欄位
            - "公允價值評估未實現損失"
                =SUM(外幣債公允價值評估!U:U)
                X欄位



- 表三：嵌入式衍生工具明細表
    - 可轉換公司債資產交換 強制透過損益按公允價值衡量之金融資產
        - 原始金額
        =VLOOKUP(120057307,餘額A!A:C,3,0)
        - 本期帳列金額
        =VLOOKUP(120057307,餘額A!A:C,3,0)+
        VLOOKUP(120077307,餘額A!A:C,3,0)+
        VLOOKUP(120079031,餘額A!A:C,3,0)
	



ai605		
		
- 表三：嵌入式衍生工具明細表		
    - 可轉換公司債資產交換 強制透過損益按公允價值衡量之金融資產 		
        - 原始金額		
120057307		
        - 本期帳列金額		
120057307   強制FVPL金融資產-衍生性ＳＷＡＰ-ASS               
120077307   強制FVPL金融資產評價調整-ＳＷＡＰ-ASS             
120079031   強制FVPL金融資產評價調整-CVA-SWAP-ASS             



PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.Category,
    SUM(ab.Amount) As SubtotalAmount
FROM AccountCodeMap
INNER JOIN
    (
        SELECT AccountBalance.AccountCode, AccountBalance.Amount
        FROM AccountBalance
        WHERE AccountBalance.DataMonthString = [DataMonthParam]
        AND AccountBalance.BalanceType = '餘額A'
    ) AS ab
ON
    AccountCodeMap.AccountCode = ab.AccountCode
WHERE
    AccountCodeMap.GroupFlag IN ('Derivative')
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust')
    AND AccountCodeMap.AssetMeasurementSubType IN ('FVPL_SWAP', 'FVPL_CVASWAP')
GROUP BY
    AccountCodeMap.Category;
