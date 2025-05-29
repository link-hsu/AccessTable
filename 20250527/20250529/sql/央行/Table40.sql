- 台幣數字
    - 衍生性金融資產
        - 國內
            - 強制透過損益按公允價值衡量之金融資產2
                - 成本 A
=+VLOOKUP(120057307,E!$A:$C,3,0)/1000

                - 公允價值6 B
=VLOOKUP(120057307,E!$A:$C,3,0)/1000+VLOOKUP(120077307,E!$A:$C,3,0)/1000+VLOOKUP(120079031,E!$A:$C,3,0)/1000

            - 避險之金融資產4
                - 成本 A
                - 公允價值6 B
        - OBU
            - 強制透過損益按公允價值衡量之金融資產2
                - 成本 A
                - 公允價值6 B
            - 避險之金融資產4
                - 成本 A
                - 公允價值6 B
        - 國外
            - 強制透過損益按公允價值衡量之金融資產2
                - 成本 A
                - 公允價值6 B
            - 避險之金融資產4
                - 成本 A
                - 公允價值6 B
    - 衍生性金融負債
        - 國內
            - 持有供交易之金融負債3
                - 成本 A
                - 公允價值6 B
            - 避險之金融負債5
                - 成本 A
                - 公允價值6 B
        - OBU
            - 持有供交易之金融負債3
                - 成本 A
                - 公允價值6 B
            - 避險之金融負債5
                - 成本 A
                - 公允價值6 B
        - 國外
            - 持有供交易之金融負債3
                - 成本 A
                - 公允價值6 B
            - 避險之金融負債5
                - 成本 A
                - 公允價值6 B        
- 外幣數字
    - 衍生性金融資產
        - 國內
            - 強制透過損益按公允價值衡量之金融資產2
                - 成本 A
                - 公允價值6 B
=VLOOKUP(1200771,E!$A:$C,3,0)/1000+Q28+VLOOKUP(120079011,E!$A:$C,3,0)/1000

            - 避險之金融資產4
                - 成本 A
                - 公允價值6 B
        - OBU
            - 強制透過損益按公允價值衡量之金融資產2
                - 成本 A
                - 公允價值6 B
            - 避險之金融資產4
                - 成本 A
                - 公允價值6 B
        - 國外
            - 強制透過損益按公允價值衡量之金融資產2
                - 成本 A
                - 公允價值6 B
            - 避險之金融資產4
                - 成本 A
                - 公允價值6 B
    - 衍生性金融負債
        - 國內
            - 持有供交易之金融負債3
                - 成本 A
                - 公允價值6 B
=VLOOKUP(2200371,E!$A:$C,3,0)/1000+S28
            - 避險之金融負債5
                - 成本 A
                - 公允價值6 B
        - OBU
            - 持有供交易之金融負債3
                - 成本 A
                - 公允價值6 B
            - 避險之金融負債5
                - 成本 A
                - 公允價值6 B
        - 國外
            - 持有供交易之金融負債3
                - 成本 A
                - 公允價值6 B
            - 避險之金融負債5
                - 成本 A
                - 公允價值6 B        





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
    AccountCodeMap.GroupFlag IN ('Derivative')
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust')
GROUP BY
    AccountCodeMap.AssetMeasurementSubType,
    AccountCodeMap.Category;
