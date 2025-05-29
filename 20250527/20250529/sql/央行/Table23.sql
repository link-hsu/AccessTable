- 融資性商業本票
    - 加權平均利率
初級市場買入
        - 1-30天
=IF(SUMIF(承銷交易!AF:AF,30,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,30,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,30,承銷交易!S:S))
        - 31-90天
=IF(SUMIF(承銷交易!AF:AF,90,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,90,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,90,承銷交易!S:S))
        - 91-180天
=IF(SUMIF(承銷交易!AF:AF,180,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,180,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,180,承銷交易!S:S))
        - 181-270天
=IF(SUMIF(承銷交易!AF:AF,270,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,270,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,270,承銷交易!S:S))
        - 271-365天
=IF(SUMIF(承銷交易!AF:AF,365,承銷交易!S:S)=0,0,SUMIF(承銷交易!AF:AF,365,承銷交易!AE:AE)/SUMIF(承銷交易!AF:AF,365,承銷交易!S:S))


承銷交易AF欄位
=IF(V2<=30,30,IF(V2<=90,90,IF(V2<=180,180,IF(V2<=270,270,365))))

AE欄位
=S2*U2

V欄位 天數
S欄位 面額
U欄位 成交利率


我有 BillTransactionDetails 資料表，
我想將Days欄位0-30天，30-90天，90-180天，180-270天，270-365天各自分組，
然後將各自的加總算出來，其中加總項目為
FaceValue欄位 * TradeYield欄位
，請問microsoft access要怎麼寫?


=IF(V23<=30,30,IF(V23<=90,90,IF(V23<=180,180,IF(V23<=270,270,365))))


PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf([BillTransactionByTradeDate.Days] >= 0 AND [BillTransactionByTradeDate.Days] <= 30, '0-30天',
    IIf([BillTransactionByTradeDate.Days] > 30 AND [BillTransactionByTradeDate.Days] <= 90, '31-90天',
    IIf([BillTransactionByTradeDate.Days] > 90 AND [BillTransactionByTradeDate.Days] <= 180, '91-180天',
    IIf([BillTransactionByTradeDate.Days] > 180 AND [BillTransactionByTradeDate.Days] <= 270, '181-270天',
    IIf([BillTransactionByTradeDate.Days] > 270 AND [BillTransactionByTradeDate.Days] <= 365, '271-365天', '其他'))))) AS DayPeriod,
    SUM([BillTransactionByTradeDate.FaceValue] * [BillTransactionByTradeDate.TradeYield]) AS 'FaceValue*TradeYield',
    SUM([BillTransactionByTradeDate.FaceValue]) AS 'FaceValue'
FROM 
    BillTransactionByTradeDate
WHERE
    BillTransactionByTradeDate.BillType NOT IN ('央行NCD', '一年以上央行NCD')
    AND BillTransactionByTradeDate.TransactionType NOT IN ('兌償/到期還本', '攤提', '附買回履約', '附買回解約', '附賣回履約', '附賣回解約')
    AND BillTransactionByTradeDate.DataMonthString = [DataMonthParam]
GROUP BY 
    IIf([BillTransactionByTradeDate.Days] >= 0 AND [BillTransactionByTradeDate.Days] <= 30, '0-30天',
    IIf([BillTransactionByTradeDate.Days] > 30 AND [BillTransactionByTradeDate.Days] <= 90, '31-90天',
    IIf([BillTransactionByTradeDate.Days] > 90 AND [BillTransactionByTradeDate.Days] <= 180, '91-180天',
    IIf([BillTransactionByTradeDate.Days] > 180 AND [BillTransactionByTradeDate.Days] <= 270, '181-270天',
    IIf([BillTransactionByTradeDate.Days] > 270 AND [BillTransactionByTradeDate.Days] <= 365, '271-365天', '其他')))));
