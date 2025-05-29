1.更新餘額C
2.至OM下載債券交易明細表：固定收益->債券->報表->債券交易明細表,依交易日->依右邊畫面勾選
3.至OM下載CBAS-BUY：衍生性商品->CBAS交易登錄及維護作業->交易報表->買進可轉換公司債資產交換簽報單->KEY入日期區間
4.至OM下載CBAS-SELL：衍生性商品->CBAS交易登錄及維護作業->交易報表->CBAS執行買回簽報單->僅勾選解約
5.更新AI405表頭月份(注意首買欄位數值,需為初級市場買進,目前需驗證交易單,手動調整)
6.依上手口述:交易量等於發行(意指本行發行)+初級+次級買賣合計數
7.目前交易量欄位不含RP履約數字,僅申報RP新作
8.外幣債目前交易皆為海外發行券種




- 台幣

    - 公債
=SUMIF(債券交易明細!AD:AD,"A",債券交易明細!M:M)+SUMIF(債券交易明細!AD:AD,"H",債券交易明細!M:M)
其中債券交易明細表

AD欄位為 
=LEFT(K2)

K欄位 債券代號

M欄位 面額
    - 金融債
=SUMIF(債券交易明細!AD:AD,"G",債券交易明細!M:M)

    - 公司債
=SUMIF(債券交易明細!AD:AD,"B",債券交易明細!M:M)

- CBAS
    - 公司債
=SUM('CBAS-BUY'!L$1:L$74)+SUM('CBAS-SELL'!K$1:K$50)

'CBAS-BUY'!L$1:L$74
L欄位 名目本金

'CBAS-SELL'!K$1:K$50
K欄位 買回本金



PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf(Left(BondTransactionDetails.Issuer, 1) IN ("A", "H"), "公債",
        IIf(Left(BondTransactionDetails.Issuer, 1) = "G", "金融債", "其他")) AS 台幣債類別,
    SUM(BondTransactionDetails.Cost) AS 面值
FROM 
    BondTransactionDetails
WHERE
    BondTransactionDetails.DataMonthString = [DataMonthParam]
    Left(BondTransactionDetails.Issuer, 1) IN ("銀行", "票券")
GROUP BY 
    IIf(Left(BondTransactionDetails.BondCode, 1) IN ("A", "H"), "公債",
        IIf(Left(BondTransactionDetails.BondCode, 1) = "G", "公司債", "其他"));




PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf(Left(BondTransactionDetails.BondCode, 1) IN ("A", "H"), "公債", 
        IIf(Left(BondTransactionDetails.BondCode, 1) = "G", "金融債", 
            IIf(Left(BondTransactionDetails.BondCode, 1) = "B", "公司債", "其他")
        )
    ) AS 台幣債類別, 
    SUM(BondTransactionDetails.FaceValue) AS 面值
FROM 
    BondTransactionDetails
WHERE 
    BondTransactionDetails.DataMonthString = [DataMonthParam] 
GROUP BY 
    IIf(Left(BondTransactionDetails.BondCode, 1) IN ("A", "H"), "公債", 
        IIf(Left(BondTransactionDetails.BondCode, 1) = "G", "金融債", 
            IIf(Left(BondTransactionDetails.BondCode, 1) = "B", "公司債", "其他")
        )
    );



- 外幣
    - 公債
        X*月底匯率
    - 金融債
        X*月底匯率


PARAMETERS DataMonthParam TEXT;
SELECT
		CloseRate.BaseCurrency,
		CloseRate.QuoteCurrency,
		CloseRate.Rate
FROM 
    CloseRate
WHERE
		CloseRate.BaseCurrency = "TWD"
		AND CloseRate.DataMonthString = [DataMonthParam]
    AND CloseRate.DataDate = (
    SELECT MAX(clsLast.DataDate)
    FROM CloseRate AS clsLast
    WHERE 
        clsLast.DataMonthString = [DataMonthParam]
        AND clsLast.BaseCurrency = "TWD"
);



- 交易量
    - 政府債券
        台幣+CBAS+外幣
    - 金融債券
        台幣+CBAS+外幣+月底匯率AUD
    - 公司債
        台幣


- 持有餘額
    - 政府債券
=IF(ISNA(VLOOKUP(120050101,'C'!$A:$C,3,0)),0,(VLOOKUP(120050101,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(120050103,'C'!$A:$C,3,0)),0,(VLOOKUP(120050103,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(125010101,'C'!$A:$C,3,0)),0,(VLOOKUP(125010101,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(121110101,'C'!$A:$C,3,0)),0,(VLOOKUP(121110101,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(121110103,'C'!$A:$C,3,0)),0,(VLOOKUP(121110103,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(122010101,'C'!$A:$C,3,0)),0,(VLOOKUP(122010101,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(122010103,'C'!$A:$C,3,0)),0,(VLOOKUP(122010103,'C'!$A:$C,3,0)))

120050101
120050103
125010101
121110101
121110103
122010101
122010103


    - 金融債券
=IF(ISNA(+VLOOKUP(122010147,'C'!$A:$C,3,0)),0,(+VLOOKUP(122010147,'C'!$A:$C,3,0)))+
 IF(ISNA(+VLOOKUP(121110147,'C'!$A:$C,3,0)),0,(+VLOOKUP(121110147,'C'!$A:$C,3,0)))+
 IF(ISNA(+VLOOKUP(120050147,'C'!$A:$C,3,0)),0,(+VLOOKUP(120050147,'C'!$A:$C,3,0)))-
 (IF(ISNA(+VLOOKUP(122010147,'C'!$A:$C,3,0)),0,(+VLOOKUP(122010147,'C'!$A:$C,3,0)))+
 IF(ISNA(+VLOOKUP(121110147,'C'!$A:$C,3,0)),0,(+VLOOKUP(121110147,'C'!$A:$C,3,0)))+
 IF(ISNA(+VLOOKUP(120050147,'C'!$A:$C,3,0)),0,(+VLOOKUP(120050147,'C'!$A:$C,3,0))))

為什麼這邊加了一樣的職，卻又扣掉一樣的值

122010147
121110147
120050147

122010147
121110147
120050147



    - 公司債
=IF(ISNA(+VLOOKUP(120057307,'C'!$A:$C,3,0)),0,(+VLOOKUP(120057307,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(120050121,'C'!$A:$C,3,0)),0,(VLOOKUP(120050121,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(120050123,'C'!$A:$C,3,0)),0,(VLOOKUP(120050123,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(121110121,'C'!$A:$C,3,0)),0,(VLOOKUP(121110121,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(121110123,'C'!$A:$C,3,0)),0,(VLOOKUP(121110123,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(122010121,'C'!$A:$C,3,0)),0,(VLOOKUP(122010121,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(122010123,'C'!$A:$C,3,0)),0,(VLOOKUP(122010123,'C'!$A:$C,3,0)))+
 IF(ISNA(VLOOKUP(121110127,'C'!$A:$C,3,0)),0,(VLOOKUP(121110127,'C'!$A:$C,3,0)))-
 (IF(ISNA(VLOOKUP(121110127,'C'!$A:$C,3,0)),0,(VLOOKUP(121110127,'C'!$A:$C,3,0))))


120057307
120050121
120050123
121110121
121110123
122010121
122010123

121110127
121110127 相互扣掉





ai405		
		
		
- 持有餘額		
    - 政府債券		
120050101		強制FVPL金融資產-公債-中央政府(我國)              
120050103		強制FVPL金融資產-公債-地方政府(我國)              
125010101		附賣回票券及債券投資-公債                         
121110101		FVOCI債務工具-公債-中央政府(我國)                 
121110103		FVOCI債務工具-公債-地方政府(我國)                 
122010101		AC債務工具投資-公債-中央政府(我國)                
122010103		AC債務工具投資-公債-地方政府(我國)                
    - 金融債券		
122010147		AC債務工具投資-金融債券-海外                      
121110147		FVOCI債務工具-金融債券-海外                       
120050147		#N/A
		
122010147		AC債務工具投資-金融債券-海外                      
121110147		FVOCI債務工具-金融債券-海外                       
120050147		#N/A
		
    - 公司債		
120057307		強制FVPL金融資產-衍生性ＳＷＡＰ-ASS               
120050121		強制FVPL金融資產-普通公司債(公營)                 
120050123		強制FVPL金融資產-普通公司債(民營)                 
121110121		FVOCI債務工具-普通公司債（公營）                  
121110123		FVOCI債務工具-普通公司債（民營）                  
122010121		AC債務工具投資-普通公司債(公營)                   
122010123		AC債務工具投資-普通公司債(民營)                   
		
121110127		FVOCI債務工具-普通公司債(民營)(外國)              
121110127		FVOCI債務工具-普通公司債(民營)(外國)              




PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementSubType & "_" & AccountCodeMap.Category As MeasurementCategory,
    SUM(ab.Amount) As SubtotalAmount
FROM AccountCodeMap
INNER JOIN
    (
        SELECT AccountBalance.AccountCode, AccountBalance.Amount
        FROM AccountBalance
        WHERE AccountBalance.DataMonthString = [DataMonthParam]
        AND AccountBalance.BalanceType = '餘額C'
    ) AS ab
ON
    AccountCodeMap.AccountCode = ab.AccountCode
WHERE
    AccountCodeMap.GroupFlag IN ('外幣債', '台幣債', 'RPRS', 'Derivative')
    AND AccountCodeMap.Category IN ('Cost')
    AND AccountCodeMap.SingleOrSubtotal = 'Single'
GROUP BY
    AccountCodeMap.AssetMeasurementType,
    AccountCodeMap.AssetMeasurementSubType,
    AccountCodeMap.Category;
