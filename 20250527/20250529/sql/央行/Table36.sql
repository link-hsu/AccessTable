票券庫存(日結)
餘額E

- 民營企業1
    - 公債
=SUMIF('債券RPRS到期時序表(債券RS庫存明細表)'!Z:Z,"銀行",'債券RPRS到期時序表(債券RS庫存明細表)'!R:R)/1000
    - 商業本票
=SUMIF('票券庫存(日結)'!H:H,"*票券",'票券庫存(日結)'!U:U)/1000
- 貨 幣 機 構4
    - 公債
=SUMIF('債券RPRS到期時序表(債券RS庫存明細表)'!Z:Z,"票券",'債券RPRS到期時序表(債券RS庫存明細表)'!R:R)/1000
    - 商業本票
=SUMIF('票券庫存(日結)'!H:H,"*銀*",'票券庫存(日結)'!U:U)/1000


C 發票人 left 2 為銀行 加總R帳上成本


RPRS到期時序表 當天到期要勾選否

PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf(Left(BondRSHoldingDetails.Issuer, 2) = "銀行", "銀行",
        IIf(Left(BondRSHoldingDetails.Issuer, 2) = "票券", "票券", "其他")) AS 發行人類別,
    SUM(BondRSHoldingDetails.BookCost) AS 總成本
FROM 
    BondRSHoldingDetails
WHERE
    BondRSHoldingDetails.DataMonthString = [DataMonthParam]
GROUP BY 
    IIf(Left(BondRSHoldingDetails.Issuer, 2) = "銀行", "銀行",
        IIf(Left(BondRSHoldingDetails.Issuer, 2) = "票券", "票券", "其他"));
        
帳上成本 BookCost
發票人 Issuer


PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf(BillHoldingDetails.Counterparty LIKE "*銀", "銀行",
        IIf(BillHoldingDetails.Counterparty LIKE "*票券", "票券", "其他")) AS 交易對手類別,
    SUM(BillHoldingDetails.BookCost) AS 總帳上成本
FROM 
    BillHoldingDetails
WHERE
    BillHoldingDetails.DataMonthString = [DataMonthParam]
GROUP BY 
    IIf(BillHoldingDetails.Counterparty LIKE "*銀", "銀行",
        IIf(BillHoldingDetails.Counterparty LIKE "*票券", "票券", "其他"));
