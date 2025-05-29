1.公債

    原始取得成本1
=+VLOOKUP(120050101,E!$A:$C,3,0)/1000+
VLOOKUP(120050103,E!$A:$C,3,0)/1000+
VLOOKUP(121110101,E!$A:$C,3,0)/1000+
VLOOKUP(121110103,E!$A:$C,3,0)/1000+
VLOOKUP(122010101,E!$A:$C,3,0)/1000+
VLOOKUP(122010103,E!$A:$C,3,0)/1000

    透過損益按公允價值衡量之金融資產2 A
=IF(ISNA(VLOOKUP(120050101,E!$A:$C,3,0)/1000),0,(VLOOKUP(120050101,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(120050103,E!$A:$C,3,0)/1000),0,(VLOOKUP(120050103,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(120070101,E!$A:$C,3,0)/1000),0,(VLOOKUP(120070101,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(120070103,E!$A:$C,3,0)/1000),0,(VLOOKUP(120070103,E!$A:$C,3,0)/1000))

    透過其他綜合損益按公允價值衡量之金融資產2 B
=IF(ISNA(VLOOKUP(121110101,E!$A:$C,3,0)/1000),0,(VLOOKUP(121110101,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(121110103,E!$A:$C,3,0)/1000),0,(VLOOKUP(121110103,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(121130101,E!$A:$C,3,0)/1000),0,(VLOOKUP(121130101,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(121130103,E!$A:$C,3,0)/1000),0,(VLOOKUP(121130103,E!$A:$C,3,0)/1000))

    按攤銷後成本衡量之債務工具投資2 C
=IF(ISNA(VLOOKUP(122010101,E!$A:$C,3,0)/1000),0,(VLOOKUP(122010101,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(122010103,E!$A:$C,3,0)/1000),0,(VLOOKUP(122010103,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(122030101,E!$A:$C,3,0)/1000),0,(VLOOKUP(122030101,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(122030103,E!$A:$C,3,0)/1000),0,(VLOOKUP(122030103,E!$A:$C,3,0)/1000))


2.公司債
2.1.公營事業
    原始取得成本1
=IF(ISNA(+VLOOKUP(120050121,E!$A:$C,3,0)/1000),0,(+VLOOKUP(120050121,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(121110121,E!$A:$C,3,0)/1000),0,(VLOOKUP(121110121,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(122010121,E!$A:$C,3,0)/1000),0,(VLOOKUP(122010121,E!$A:$C,3,0)/1000))

    透過損益按公允價值衡量之金融資產2 A
=IF(ISNA(VLOOKUP(120050121,E!$A:$C,3,0)/1000),0,(VLOOKUP(120050121,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(120070121,E!$A:$C,3,0)/1000),0,(VLOOKUP(120070121,E!$A:$C,3,0)/1000))

    透過其他綜合損益按公允價值衡量之金融資產2 B
=+VLOOKUP(121110121,E!$A:$C,3,0)/1000+VLOOKUP(121130121,E!$A:$C,3,0)/1000

    按攤銷後成本衡量之債務工具投資2 C
=+IF(ISNA(VLOOKUP(122010121,E!$A:$C,3,0)/1000),0,(VLOOKUP(122010121,E!$A:$C,3,0)/1000))

2.2.民營企業-國內公司債
    原始取得成本1
=+VLOOKUP(120050123,E!$A:$C,3,0)/1000+VLOOKUP(121110123,E!$A:$C,3,0)/1000+VLOOKUP(122010123,E!$A:$C,3,0)/1000

    透過損益按公允價值衡量之金融資產2 A
=IF(ISNA(VLOOKUP(120050123,E!$A:$C,3,0)/1000),0,(VLOOKUP(120050123,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(120070123,E!$A:$C,3,0)/1000),0,(VLOOKUP(120070123,E!$A:$C,3,0)/1000))

    透過其他綜合損益按公允價值衡量之金融資產2 B
=+VLOOKUP(121110123,E!$A:$C,3,0)/1000+VLOOKUP(121130123,E!$A:$C,3,0)/1000

    按攤銷後成本衡量之債務工具投資2 C
=+VLOOKUP(122010123,E!$A:$C,3,0)/1000+VLOOKUP(122030123,E!$A:$C,3,0)/1000

3.股票及股權投資-民營企業

    原始取得成本1
=IF(ISNA(+VLOOKUP(1200503,E!$A:$C,3,0)/1000),0,(+VLOOKUP(1200503,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(1210103,E!$A:$C,3,0)/1000),0,(VLOOKUP(1210103,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(15501,E!$A:$C,3,0)/1000),0,(VLOOKUP(15501,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(121019901,E!$A:$C,3,0)/1000),0,(VLOOKUP(121019901,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(150019901,E!$A:$C,3,0)/1000),0,(VLOOKUP(150019901,E!$A:$C,3,0)/1000))-C19

    透過損益按公允價值衡量之金融資產2 A
=+VLOOKUP(1200503,E!$A:$C,3,0)/1000+VLOOKUP(1200703,E!$A:$C,3,0)/1000-D19

    透過其他綜合損益按公允價值衡量之金融資產2 B
=IF(ISNA(+VLOOKUP(1210103,E!$A:$C,3,0)/1000),0,(+VLOOKUP(1210103,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(1210303,E!$A:$C,3,0)/1000),0,(VLOOKUP(1210303,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(1210199,E!$A:$C,3,0)/1000),0,(VLOOKUP(1210199,E!$A:$C,3,0)/1000))+
IF(ISNA(VLOOKUP(1210399,E!$A:$C,3,0)/1000),0,(VLOOKUP(1210399,E!$A:$C,3,0)/1000))-E19

    按攤銷後成本衡量之債務工具投資2 C
    
    採用權益法之投資-淨額2 E
=+VLOOKUP(15001,E!$A:$C,3,0)/1000+
VLOOKUP(15003,E!$A:$C,3,0)/1000

4.受益憑證-其他

    原始取得成本1
=+VLOOKUP(1200505,E!$A:$C,3,0)/1000

    透過損益按公允價值衡量之金融資產2 A
=+VLOOKUP(1200505,E!$A:$C,3,0)/1000+
VLOOKUP(1200705,E!$A:$C,3,0)/1000

    透過其他綜合損益按公允價值衡量之金融資產2 B

    按攤銷後成本衡量之債務工具投資2 C

5.新台幣可轉讓定期存單-中央銀行發行

    原始取得成本1
=VLOOKUP(121110911,E!A:C,3,0)/1000+
VLOOKUP(122010911,E!A:C,3,0)/1000

    透過損益按公允價值衡量之金融資產2 A

    透過其他綜合損益按公允價值衡量之金融資產2 B
=VLOOKUP(121110911,E!A:C,3,0)/1000+
VLOOKUP(121130911,E!A:C,3,0)/1000

    按攤銷後成本衡量之債務工具投資2 C
=VLOOKUP(122010911,E!A:C,3,0)/1000+
VLOOKUP(122030911,E!A:C,3,0)/1000

6.商業本票-民營企業

    原始取得成本1
=VLOOKUP(120050903,E!A:C,3,0)/1000

    透過損益按公允價值衡量之金融資產2 A
=+VLOOKUP(120050903,E!$A:$C,3,0)/1000+
VLOOKUP(120070903,E!$A:$C,3,0)/1000

    透過其他綜合損益按公允價值衡量之金融資產2 B

    按攤銷後成本衡量之債務工具投資2 C

7.國外機構發行-在國外發行-長期債票券6

    原始取得成本1
=VLOOKUP(140010147,E!A:C,3,0)/1000

    透過損益按公允價值衡量之金融資產2 A

    透過其他綜合損益按公允價值衡量之金融資產2 B
=VLOOKUP(140010147,E!A:C,3,0)/1000+
VLOOKUP(140030147,E!A:C,3,0)/1000

    按攤銷後成本衡量之債務工具投資2 C





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
        AND AccountBalance.BalanceType = '餘額E'
    ) AS ab
ON
    AccountCodeMap.AccountCode = ab.AccountCode
WHERE
    AccountCodeMap.GroupFlag IN ('台幣債', '股權', '受益憑證', 'Bill', 'CP')
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust', 'ImpairmentLoss')
GROUP BY
    AccountCodeMap.AssetMeasurementType,
    AccountCodeMap.AssetMeasurementSubType,
    AccountCodeMap.Category;



PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AccountTitle,
    ab.Amount
FROM AccountCodeMap
INNER JOIN
    (
        SELECT AccountBalance.AccountCode, AccountBalance.Amount
        FROM AccountBalance
        WHERE AccountBalance.DataMonthString = [DataMonthParam]
        AND AccountBalance.BalanceType = '餘額E'
    ) AS ab
ON
    AccountCodeMap.AccountCode = ab.AccountCode
WHERE
    AccountCodeMap.GroupFlag IN ('台幣債')
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust', 'ImpairmentLoss')
    AND AccountCodeMap.SingleOrSubtotal = 'Single'
    AND AccountCodeMap.AssetType = 'Company'
