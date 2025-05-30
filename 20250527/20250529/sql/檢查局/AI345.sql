1.更新匯率及餘額C


你可以在同一個 `SELECT` 裡，利用 Access 的條件式聚合（`IIF`）來分別計算「只加總大於 0 的金額」和「只加總小於 0 的金額」。大致上把原本的


SUM(ab.Amount) AS SubtotalAmount


改成


SUM(IIF(ab.Amount > 0, ab.Amount, 0))  AS CreditBalance,   -- 貸餘（Positive）
SUM(IIF(ab.Amount < 0, ab.Amount, 0))  AS DebitBalance     -- 借餘（Negative）


完整範例如下：


PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementSubType
      & "_" & AccountCodeMap.Category      AS MeasurementCategory,
    SUM(IIF(ab.Amount > 0, ab.Amount, 0))  AS CreditBalance,   -- 貸餘
    SUM(IIF(ab.Amount < 0, ab.Amount, 0))  AS DebitBalance     -- 借餘
FROM AccountCodeMap
INNER JOIN
    (
        SELECT 
            AccountBalance.AccountCode, 
            AccountBalance.Amount
        FROM AccountBalance
        WHERE 
            AccountBalance.DataMonthString = [DataMonthParam]
            AND AccountBalance.BalanceType = '餘額C'
    ) AS ab
    ON AccountCodeMap.AccountCode = ab.AccountCode
WHERE
    AccountCodeMap.GroupFlag       IN ('台幣債','股權','受益憑證','CP','Derivative','Bill')
    AND AccountCodeMap.Category    IN ('Cost','ValuationAdjust','ImpairmentLoss','ImpairmentAllowance')
    AND AccountCodeMap.SingleOrSubtotal = 'Single'
GROUP BY
    AccountCodeMap.AssetMeasurementType,
    AccountCodeMap.AssetMeasurementSubType,
    AccountCodeMap.Category;


**說明：**

* `IIF(ab.Amount > 0, ab.Amount, 0)`：
  只有在 `ab.Amount > 0` 時才回傳該值，否則回傳 `0`，最後 `SUM` 起來就是「所有正數的合計」，也就是你的「貸餘」。

* `IIF(ab.Amount < 0, ab.Amount, 0)`：
  只有在 `ab.Amount < 0` 時才回傳該值，否則回傳 `0`，最後 `SUM` 起來就是「所有負數的合計」，也就是「借餘」。

這樣一張查詢裡就會同時回傳每個分類的「貸餘」與「借餘」兩個欄位。若你想要把「借餘」顯示成正數，可以再包一層 `ABS(...)`：

```sql
SUM(IIF(ab.Amount < 0, ABS(ab.Amount), 0)) AS DebitBalance
```

— 以上即可滿足你的需求。
