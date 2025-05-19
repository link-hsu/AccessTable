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
