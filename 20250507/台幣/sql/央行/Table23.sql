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
    IIf([BillTransactionDetails.Days] >= 0 AND [BillTransactionDetails.Days] <= 30, '0-30天',
    IIf([BillTransactionDetails.Days] > 30 AND [BillTransactionDetails.Days] <= 90, '31-90天',
    IIf([BillTransactionDetails.Days] > 90 AND [BillTransactionDetails.Days] <= 180, '91-180天',
    IIf([BillTransactionDetails.Days] > 180 AND [BillTransactionDetails.Days] <= 270, '181-270天',
    IIf([BillTransactionDetails.Days] > 270 AND [BillTransactionDetails.Days] <= 365, '271-365天', '其他'))))) AS DayPeriod,
    SUM([BillTransactionDetails.FaceValue] * [BillTransactionDetails.TradeYield]) AS 'FaceValue*TradeYield'
FROM 
    BillTransactionDetails
WHERE
    BillTransactionDetails.DataMonthString = [DataMonthParam]
GROUP BY 
    IIf([BillTransactionDetails.Days] >= 0 AND [BillTransactionDetails.Days] <= 30, '0-30天',
    IIf([BillTransactionDetails.Days] > 30 AND [BillTransactionDetails.Days] <= 90, '31-90天',
    IIf([BillTransactionDetails.Days] > 90 AND [BillTransactionDetails.Days] <= 180, '91-180天',
    IIf([BillTransactionDetails.Days] > 180 AND [BillTransactionDetails.Days] <= 270, '181-270天',
    IIf([BillTransactionDetails.Days] > 270 AND [BillTransactionDetails.Days] <= 365, '271-365天', '其他')))));
