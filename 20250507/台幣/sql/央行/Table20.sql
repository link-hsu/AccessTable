餘額E

- 民營企業
    - 其他到期日
        - 公債
=IF(ISNA(+VLOOKUP(225010101,E!A:C,3,0)/1000),0,(+VLOOKUP(225010101,E!A:C,3,0)/1000))
        - 公司債
=IF(ISNA(+VLOOKUP(225010105,E!A:C,3,0)/1000),0,(+VLOOKUP(225010105,E!A:C,3,0)/1000))
        - 商業本票
=IF(ISNA(+VLOOKUP(225010303,E!A:C,3,0)/1000),0,(+VLOOKUP(225010303,E!A:C,3,0)/1000))



3.與債券RP/RS報表核對是否有餘額
這部分要詢問是核對什麼，是否也抓出表的資料可以直接核對



PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementSubType & "_" & AccountCodeMap.Category As MeasurementCategory,
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
    AccountCodeMap.GroupFlag IN ('RPRS');
