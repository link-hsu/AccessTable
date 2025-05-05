- 前端
    - 每次生成報表access資料表中沒有一個獨立可以驗證是哪次生成的欄位，造成儲存的資料無法驗證是哪次生成的資料
    - (已修正) Button Click後執行程序
    - reportName使用class直接取得，不要再宣告新的參數
    - (已修正) 測試dataMonthString用input輸入方式
    - (已修正) FM2 FBA3 有空白欄位無法顯示 
    - (已修正) 前端以日期變數輸入抓access資料
    - (已修正) 前端加入log
    - (已修正) FM10還要加入AC
    - (已修正) CloseRate要篩選出DataDate的最後一天，FM2和表2
    - (已修正) F1F2表需要Create出要讓F2表填的欄位
    - (?修正?) AI602 經過驗算結果與預期一致，但與柏丞結果不一致
    - (已修正) 修改 Access SQL為以param方式抓取資料 修改到ai240
    - mm4901b swift 沒篩完全，剩下2024年開始之後貼，CTBC 中信swift code要再確認是不是不只一個
    - (已確認沒問題) f1_f2 還要修正 bug
    - (已修正) Think 建立一個Swift Table Primary key is swift code for FM2 Report(或許直接改以swift code為primary key就好，這要確認一下SQL代碼) => 原本Primary Key為DataID，只需要直接往下增加資料就可以
    - FM 11 ACCAccount不確定的有還沒加入SQL和程式碼中: FVPL處分利益全部、FVOCI處分利益 金融債券、AC處分利益 金融債券、FVPL處分損失 金融債券、AC處分損失全部、FVPL金融資產利息收入 金融債券、FVPL金融資產評價利益 金融債券、FVPL金融資產評價損失 金融債券
    - FM 13裡面的國家需要人工篩選
    - 
    - (已修正) BankDirectory 中的中國信託swift如果沒抓到，就不會抓到BankCode，銀行名稱就是空白 => 原本Primary Key為DataID，只需要直接往下增加資料就可以

- 其他項目，例如後端
    - (已修正) 整理資料庫匯入程式碼
    - (已修正) 解決開啟多個Excel檔案時會出錯的問題，是因為 ImporterStandard 那邊又定義一次xlApp導致，從外面的Sub傳進去之後就沒問題了。
    - (已修正) 後端增加 CloseRate
    - (已修正) 加入處理 F2 後端
    - (已修正) 加入處理 F2 前端
    - (已完成) 完成剩下後端資料彙入 (run 20240329的OBU_AC5601時，在.txt檔的250行應該是EUR，但是資料錯誤文字是 CNY，導致BUG)
    - (已完成) 匯入這個月的資料庫跑看看有沒有問題
    - (已修正) 將CloseRate併入後端相同檔案
    - (已修正) 核對Swiftcode
    - (已修正) 建立 CloseRate Func
    - (已完成) 測試在共用路徑能不能用
    - (已修正) 建立完整Log管理機制
    - (已完成) 加入錯誤處理
    - (已修正) 清理欄位時，有些資料表會新增空白資料
    - (已修正) 減損後端匯入的時候，會漏掉 AC債務工具投資-公債-中央政府(? 這張表沒有匯入，需將已經匯入的資料庫重新跑一次
    - 將每天的CloseRate匯入




- 代辦順序
    - (已完成) 將抓取匯率的function放進去
    - (已修正) 後端抓 AC債務工具投資-公債-中央政府(? 報表沒有匯入

    - (已修正) MM4901b要往前篩swift code
    - 開始製作新的報表，1.嘉妤那邊的表。 先整理一下全部還有哪些表需要匯入(報表哪邊下載取得，怎麼放到後端method清理)，後端部分先處理
                    2.熊軒那邊的表，新增分頁撰寫處理邏輯

    - class範例





- 資料庫部分
    - ACCAccount 使用餘額C代號去Map科目名稱
    - 另外建立一個資料表，欄位名稱有 DataID DataDate DataMonth DataMonthString BalanceType (A台幣 C全部 D OBU E 台幣+DBU) AccountCode Amount



- 待匯入報表
    - 餘額A C D E
    - 票券交易明細表
    - 票券庫存明細表(日結)
    - 債券風險部位餘額表
    - 票券風險部位餘額表
    - PNCDCAL
    - 債券RP/RS
    - 債券RP/RS 到期時序表
    - OM CBAS-BUY
    - OM CBAS-SELL
    - 票券交易明細表_發行及首購
    - 票券交易明細表_含CP2及央NCD
    - 債券評價表 for 匯出檔案
    - 票券評價表 for 匯出檔案

- 待製作
    -  外幣
        已經有製表說明
        AI822-對大陸地區之授信計算表
        
        台幣
        AI230 利率敏感性資產負債分析表-新臺幣
        AI240 新臺幣到期日期限結構分析表

        AI230 AI240
        全體資負表
        餘額E
        餘額A
        利率敏感性存款統計表
        利率敏感性放款統計表
    - 台幣
        - 央行
            - 底稿 表
                - 華泰商業銀行ＮＣＤ 發行 償還及餘額統計表 
                - 票券交易明細表
                - 票券庫存明細表(日結)
                - 會計資料庫 餘額 A C D E
                - 債券風險部位餘額表
                - 票券風險部位餘額表
            - 表10
                - 需報表資料
                    - 餘額E
            - 表15A
                - 需報表資料
                    - 電子化報表下載月底日之PNCDCAL
            - 表15B
                - 需報表資料
                    - 無資料
            - 表16
                - 需報表資料
                    - 無資料
            - 表20
                - 需報表資料
                    - 餘額E
                    - 債券RP/RS
                    - RP/RS到期時序表
            - 表22
                - 需報表資料
                    - 票券交易明細表(留意E欄交易對手若為銀行,字串可能會有隱藏空格需刪除(使用取代))
            - 表23
                - 需報表資料
                    - 票券交易明細表 承銷交易 (票券別央行NCD相關不勾選、交易類別兌償到期及履約相關不勾選)
            - 表27
                - 需報表資料
                    - 人工確認
            - 表36
                - 需報表資料
                    - 餘額E
                    - 票券庫存明細表(日結)：OM->固定收益->票券->日報表->票券庫存明細表(日結)(H欄交易對手若為銀行,字串可能會有空格需刪除)
                    - 債券RS部位(基本上目前不會有) 債券RPRS到期時序表(債券RS庫存明細表)
            - (B008)金融機構承作新台幣結構型商品季報表
                - 無資料
            - 銀行投資部位連結美國次級房貸統計
                - 無資料

            
        - 檢查局
            - AI233 新臺幣投資明細之利率變動情境分析表
                - 需報表資料
                    - 餘額C
                    - 總務提供 不動產投資部位，留意是否有台幣投資相關新會計
                    - 債券風險部位餘額表
                    - 票券風險部位餘額表
                - 有關情境分析之DV01算法要怎麼由
                  透過損益按公允價值衡量之 投資成本 帳面價值 去計算 本期損益
                  其公式是什麼
            - AI345
                - 需報表資料
                    - TMU違約客戶明細表(取衍商違約金額)
                    - 餘額C
                    - CloseRate
                    - 減損檔案
                    - PIBALL23 會科 減 備抵呆帳-應收衍生性商品違約那頁
            - AI405
                - 需報表資料
                    - OM 債券交易明細表
                    - OM CBAS-BUY
                    - OM CBAS-SELL
                    - 餘額C
            - AI410
                - 需報表資料
                    - 票券交易明細表_發行及首購
            - AI415
                - 需報表資料
                    - 票券交易明細表_含CP2及央NCD
            - AI430
                - 需報表資料
                    - 餘額C
            - AI601
                - 需報表資料
                    - 債券評價表 for 匯出檔案
                    - 票券評價表 for 匯出檔案
                    - 餘額 C
                    - 餘額 D
                    - 外幣債評估表
                    - 手動填入轉投資成本
                    - 總務提供不動產投資餘額
                    - CloseRate
            - AI605 半年報
                - 需報表資料
                    - CloseRate
                    - 餘額A
                    - 餘額C
                    - 外幣債評估表
                    - 持有到期債券風險餘額表
                    - 債券評價表
                    - 減損表
            - AI816
                - 需報表資料
                    - 無資料
            - AI271-季報
                - 需報表資料
                    - 無資料
            - AI272-季報
                - 需報表資料
                    - 無資料
            - AI273-季報
                - 需報表資料
                    - 無資料
            - AI281-季報
                - 需報表資料
                    - 無資料
            - AI282-季報
                - 需報表資料
                    - 無資料

RP新作
RP履約





2. Function sub使用access
3. 使用Excel寫好的Fucntion直接Call Access
4. Access基本使用
6. ref債券名稱要先更新DB
7. 往前拉mm4901b
8. 介面有沒有更好方式
9. ctrl f3
10. 程式碼結構
11. Log檔在哪裡



債券交易明細表

發行前買斷
發行前賣斷
首購
買斷
賣斷
附買回
附買回續作
附賣回
附賣回續作


CBAS BUY
CBAS SELL
僅勾選解約





會計資料庫處理
for loop if sheetName=餘額A 餘額C 餘額D 餘額E
其他的分頁都刪除

將Column D和B刪除
將Column C移到Column A

增加三列
DataDate DataMonth DataMonthString

票券交易明細表 統一編號，補到8位數，票載利率要拿掉百分比



- 20250428 Work
    - Check ai602 fm10 來源報表及計算
    1. 建立台幣Access資料表
        a. 完成DB column mapping 
        
        1.2 清洗資料表資料，欄位%調整和餘額欄位整併，另外還有利率敏感性分析的表
    2. 建立AccountCode Pair ReportType
        2.1.1 建立好配對表，取得原始配對方式之後，分別為不同報表需求建立Query(兩種方式，一種是建立一個帶入參數的共用Query，參數會由前端代碼輸入，另一種是為不同的表建立不同的Query，就後續維護角度是共用參數較為理想)
    3. check 前端資料取得函數Paramater
    4. 修改外幣SQL，可能要增加傳入參數
    5. 台幣製作報表邏輯
    6. 台幣有些欄位資料需要人工輸入，可能放在Setting頁面，VBA操控Control頁面清除必要填入欄位資料，程式碼在剛運行時，檢核相關欄位是否填入，如果沒有填入則強制中止Sub運行
    7. 

    - 修改Valuation MapTable:步驟
        1. 減損及外幣債表加入新增欄位
        2. 修改 DBsColTable 欄位
        3. 修改FilePaths
        4. 修改程式碼
        5. 更新rawFile資料
        6. Create Configuration Data
        7. 點擊更新按鈕

    外幣需要處理的報表
    AI602
    FB1
    FM10
    FM11
    







- 說明整個運作流程
    1. 原始批次、資通、外幣債等檔案在哪邊，資料呈現如何。
    2. 原始檔案放在Relation\RawData\項下，更改檔名為正確格式。
    3. 開啟Access的Configuration Form表單，填入ReportDataDate(輸入原始報表資料日期，日期格式為 YYYY/MM/DD，例如: 2025/2/27)，按Tab後ReportMonth會自動填入為月底日期，RawFilePath和CopyFilePath為原始檔案路徑和複製檔案路徑(複製檔案路徑中的檔案後續會清理為正規格式)
    4. 填妥後滑鼠點擊更新Configuration，會在Configuration資料表中加入一筆設定資料(後續打開Access之後預設值會抓取這筆資料內容)，並將RawFilePath中的檔案複製一份到CopyFilePath
    5. 點擊【上傳至Access DB】後即開始批次清洗Excel檔案為正規格式並匯入Access資料庫中。

    1. 匯入CloseRate，進入CloseRate Form，填入關帳匯率日期及完整檔案路徑，點擊【Update Close Rate】

    1. 資料表欄位統整資料Excel檔，以及批次建立資料表腳本，相關欄位名稱及資料類型等可以參考，如果有需要再調整也可以參考。

    1. Log檔放哪邊

    1. Table BankDirectory 外幣拆款報表，要填報和那些銀行程作多少金額會用到
        DBsColTable 為彙整所有資料表欄位，這個欄位也需要更新，因為會影響到最後匯入資料庫的資料表名稱。
        FilePath就是要設定想要清理和匯入資料庫的報表名稱
        FXDebtReferCountry 為FM13建立的資料表，因為相關table沒辦法篩出外幣債是哪個國家發行，這邊只能手動填入

        MonthlyDeclarationReport 為產製申報報表填報的欄位和值的資料
    - 有關後端路徑部分，需先建立Configuration檔案，設定rawFilePath和copyFilePath，點擊後會依照這兩個路徑由rawFilePath複製一份檔案到copyFilePath，並在Configuration建立一筆路徑和日期等資料，點擊【上傳至Access DB】時會從最新那筆Configuration建立的路徑存取資料，所以一定要先建立一筆Configuration

    2. 前端產生報表，









## 🏦 債券交易明細表 BondTransactionDetails
| 中文欄位 | 英文命名 |
|----------|-----------|
交易類別 | `TransactionType`  
交易處所 | `TradingVenue`  
成交單號 | `TradeID`  
帳務目的 | `AccountingPurpose`  
關係人 | `RelatedParty`  
客戶統編 | `CustomerTaxID`  
客戶名稱 | `CustomerName`  
交易日期 | `TradeDate`  
交割日期 | `SettlementDate`  
到期日期 | `MaturityDate`  
債券代號 | `BondCode`  
債券名稱 | `BondName`  
面額 | `FaceValue`  
成交價格 | `TradePrice`  
應計利息 | `AccruedInterest`  
應收/付金額 | `NetAmount`  
利息累計稅款 | `AccruedTax`  
天數 | `Days`  
利率% | `InterestRate`  
到期金額 | `MaturityAmount`  
承作期間稅款 | `TaxDuringPeriod`  
借券費率 | `BorrowFeeRate`  
借券金額 | `BorrowAmount`  
帳上成本 | `BookCost`  
帳上利息 | `BookInterest`  
利益 | `Profit`  
損失 | `Loss`  
交易員 | `Trader`  
覆核層級 | `ReviewLevel`  

---

英文欄位	Access 資料型態	備註／欄位長度
TransactionType	Short Text (255)	交易類別
TradingVenue	Short Text (255)	交易處所
TradeID	Short Text (255)	主鍵 candidate，可再加索引
AccountingPurpose	Short Text (255)	帳務目的
RelatedParty	Short Text (255)	關係人
CustomerTaxID	Short Text (50)	統一編號
CustomerName	Short Text (255)	客戶名稱
TradeDate	Date/Time	交易日期
SettlementDate	Date/Time	交割日期
MaturityDate	Date/Time	到期日期
BondCode	Short Text (50)	債券代號
BondName	Short Text (255)	債券名稱
FaceValue	Currency	面額
TradePrice	Currency	成交價格
AccruedInterest	Currency	應計利息
NetAmount	Currency	應收/付金額
AccruedTax	Currency	利息累計稅款
Days	Number (Integer)	天數
InterestRate	Number (Double)	利率 %
MaturityAmount	Currency	到期金額
TaxDuringPeriod	Currency	承作期間稅款
BorrowFeeRate	Number (Double)	借券費率
BorrowAmount	Currency	借券金額
BookCost	Currency	帳上成本
BookInterest	Currency	帳上利息
Profit	Currency	利益
Loss	Currency	損失
Trader	Short Text (255)	交易員
ReviewLevel	Number (Integer)	覆核層級


---

Text
Text
Text
Text
Text
Text
Text
Date
Date
Date
Text
Text
Currency
Double
Currency
Currency
Currency
Number
Double
Currency
Currency
Double
Currency
Currency
Currency
Currency
Currency
Text
Number




---
TransactionType
TradingVenue
TradeID
AccountingPurpose
RelatedParty
CustomerTaxID
CustomerName
TradeDate
SettlementDate
MaturityDate
BondCode
BondName
FaceValue
TradePrice
AccruedInterest
NetAmount
AccruedTax
Days
InterestRate
MaturityAmount
TaxDuringPeriod
BorrowFeeRate
BorrowAmount
BookCost
BookInterest
Profit
Loss
Trader
ReviewLevel

---

交易類別
交易處所
成交單號
帳務目的
關係人
客戶統編
客戶名稱
交易日期
交割日期
到期日期
債券代號
債券名稱
面額
成交價格
應計利息
應收/付金額
利息累計稅款
天數
利率%
到期金額
承作期間稅款
借券費率
借券金額
帳上成本
帳上利息
利益
損失
交易員  
覆核層級


## 📊 債券風險部位餘額 BondRiskPositionBalance
| 中文欄位 | 英文命名 |
|----------|-----------|
債券代號 | `BondCode`  
債券簡稱 | `BondShortName`  
到期日 | `MaturityDate`  
保證人 | `Guarantor`  
票載利率 | `CouponRate`  
存續期間 | `Duration`  
市場利率 | `MarketRate`  
百元價格 | `PricePer100`  
庫存收益率 | `InventoryYield`  
庫存面額 | `InventoryFaceValue`  
庫存市價 | `InventoryMarketValue`  
庫存成本 | `InventoryCost`  
損益 | `ProfitLoss`  
應計利息 | `AccruedInterest`  
DV01 | `DV01`  
帳務目的 | `AccountingPurpose`  

---

英文欄位	Access 資料型態	備註
BondCode	Short Text (50)	債券代號
BondShortName	Short Text (255)	債券簡稱
MaturityDate	Date/Time	到期日
Guarantor	Short Text (255)	保證人
CouponRate	Number (Double)	票載利率 %
Duration	Number (Double)	存續期間
MarketRate	Number (Double)	市場利率 %
PricePer100	Number (Double)	百元價格
InventoryYield	Number (Double)	庫存收益率 %
InventoryFaceValue	Currency	庫存面額
InventoryMarketValue	Currency	庫存市價
InventoryCost	Currency	庫存成本
ProfitLoss	Currency	損益
AccruedInterest	Currency	應計利息
DV01	Number (Double)	DV01
AccountingPurpose	Short Text (255)	帳務目的


---

英文欄位	Access 資料型態	備註
Text
Text
Date
Text
Double
Double
MDouble
Double
Double
Currency
Double
Currency
Currency
Currency
Double
Text

---
BondCode
BondShortName
MaturityDate
Guarantor
CouponRate
Duration
MarketRate
PricePer100
InventoryYield
InventoryFaceValue
InventoryMarketValue
InventoryCost
ProfitLoss
AccruedInterest
DV01
AccountingPurpose


---

債券代號
債券簡稱
到期日
保證人
票載利率
存續期間
市場利率
百元價格
庫存收益率
庫存面額
庫存市價
庫存成本
損益
應計利息
DV01  
帳務目的


---

## 📈 票券風險部位餘額 BillRiskPositionBalance
| 中文欄位 | 英文命名 |
|----------|-----------|
部門代號 | `DepartmentCode`  
部門名稱 | `DepartmentName`  
評價日 | `ValuationDate`  
幣別 | `Currency`  
排序 | `SortOrder`  
帳務目的 | `AccountingPurpose`  
票別 | `BillType`  
票券批號 | `BillBatchNumber`  
票券批號流水號 | `BillBatchSerial`  
發票人 | `Issuer`  
庫存面額 | `InventoryFaceValue`  
發票日 | `IssueDate`  
到期日 | `MaturityDate`  
交割日 | `SettlementDate`  
成交利率 | `TradeYield`  
市場利率 | `MarketYield`  
萬元單價 | `PricePerTenThousand`  
距到期日/天 | `DaysToMaturity`  
庫存成本 | `InventoryCost`  
現值 | `PresentValue`  
市價 | `MarketValue`  
評價損益 | `ValuationGainLoss`  
DVO1 | `DV01`  

---

英文欄位	Access 資料型態	備註
DepartmentCode	Short Text (50)	部門代號
DepartmentName	Short Text (255)	部門名稱
ValuationDate	Date/Time	評價日
Currency	Short Text (50)	幣別
SortOrder	Number (Integer)	排序
AccountingPurpose	Short Text (255)	帳務目的
BillType	Short Text (50)	票別
BillBatchNumber	Short Text (100)	票券批號
BillBatchSerial	Short Text (100)	票券批號流水號
Issuer	Short Text (255)	發票人
InventoryFaceValue	Currency	庫存面額
IssueDate	Date/Time	發票日
MaturityDate	Date/Time	到期日
SettlementDate	Date/Time	交割日
TradeYield	Number (Double)	成交利率 %
MarketYield	Number (Double)	市場利率 %
PricePerTenThousand	Number (Double)	萬元單價
DaysToMaturity	Number (Integer)	距到期日/天
InventoryCost	Currency	庫存成本
PresentValue	Currency	現值
MarketValue	Currency	市價
ValuationGainLoss	Currency	評價損益
DV01	Number (Double)	DVO1


---

英文欄位	Access 資料型態	備註
Text
Text
Date
Text
Number
Text
Text
Text
Text
Text
Currency
Date
Date
Date
Double
Double
Double
Number
Currency
Currency
Currency
Currency
Double

---

DepartmentCode
DepartmentName
ValuationDate
Currency
SortOrder
AccountingPurpose
BillType
BillBatchNumber
BillBatchSerial
Issuer
InventoryFaceValue
IssueDate
MaturityDate
SettlementDate
TradeYield
MarketYield
PricePerTenThousand
DaysToMaturity
InventoryCost
PresentValue
MarketValue
ValuationGainLoss
DV01


---


部門代號
部門名稱
評價日
幣別
排序
帳務目的
票別
票券批號
票券批號流水號
發票人
庫存面額
發票日
到期日
交割日
成交利率
市場利率
萬元單價
距到期日/天
庫存成本
現值
市價
評價損益
DVO1

## 🛒 CBAS BUY CbasBuy
| 中文欄位 | 英文命名 |
|----------|-----------|
成交單號 | `TradeID`  
契約編號 | `ContractID`  
交易日期 | `TradeDate`  
交割日期 | `SettlementDate`  
債券代號 | `BondCode`  
標的名稱 | `UnderlyingName`  
申請額度 | `RequestedLimit`  
已動用額度 | `UsedLimit`  
客戶名稱 | `CustomerName`  
交易對象額度 | `CounterpartyLimit`  
已動用額度 | `UsedCounterpartyLimit`  
名目本金 | `NotionalPrincipal`  
可轉債到期日 | `ConvertibleBondMaturityDate`  
選擇權到期日 | `OptionMaturityDate`  
交換利率(%) | `SwapRate`  
計息期間 | `InterestPeriod`  
付息日 | `InterestPaymentDate1`...`InterestPaymentDate4`  
可轉債形式 | `ConvertibleBondType`  
備註 | `Remarks`  

---
英文欄位	Access 資料型態	備註
TradeID	Short Text (255)	成交單號
ContractID	Short Text (255)	契約編號
TradeDate	Date/Time	交易日期
SettlementDate	Date/Time	交割日期
BondCode	Short Text (50)	債券代號
UnderlyingName	Short Text (255)	標的名稱
RequestedLimit	Currency	申請額度
UsedLimit	Currency	已動用額度
CustomerName	Short Text (255)	客戶名稱
CounterpartyLimit	Currency	交易對象額度
UsedCounterpartyLimit	Currency	已動用對手額度
NotionalPrincipal	Currency	名目本金
ConvertibleBondMaturityDate	Date/Time	可轉債到期日
OptionMaturityDate	Date/Time	選擇權到期日
SwapRate	Number (Double)	交換利率 (%)
InterestPeriod	Short Text (255)	計息期間
InterestPaymentDate1…4	Date/Time	付息日 (最多四次)
ConvertibleBondType	Short Text (255)	可轉債形式
Remarks	Long Text	備註 (多行)

---

英文欄位	Access 資料型態	備註
Text
Text
Date
Date
Text
Text
Currency
Currency
Text
Currency
Currency
Currency
Date
Date
Double
Text
Date
Text
Text

---

TradeID
ContractID
TradeDate
SettlementDate
BondCode
UnderlyingName
RequestedLimit
UsedLimit
CustomerName
CounterpartyLimit
UsedCounterpartyLimit
NotionalPrincipal
ConvertibleBondMaturityDate
OptionMaturityDate
SwapRate
InterestPeriod
InterestPaymentDate1
ConvertibleBondType
Remarks


---

成交單號
契約編號
交易日期
交割日期
債券代號
標的名稱
申請額度
已動用額度
客戶名稱
交易對象額度
已動用額度
名目本金
可轉債到期日
選擇權到期日
交換利率(%)
計息期間
付息日
可轉債形式
備註

## 🧾 CBAS SELL CbasSell
| 中文欄位 | 英文命名 |
|----------|-----------|
成交單號 | `TradeID`  
契約編號 | `ContractID`  
解約成交單號 | `TerminationTradeID`  
交易日期 | `TradeDate`  
交割日期 | `SettlementDate`  
標的代號 | `UnderlyingCode`  
標的名稱 | `UnderlyingName`  
客戶名稱 | `CustomerName`  
可轉債到期日 | `ConvertibleBondMaturityDate`  
名目本金 | `NotionalPrincipal`  
買回本金 | `RepurchasePrincipal`  
交換利率(%) | `SwapRate`  
解約利率(%) | `TerminationRate`  
計息起日 | `InterestStartDate`  
計息迄日 | `InterestEndDate`  
利息 | `InterestAmount`  
違約金 | `Penalty`  
到期補償金 | `MaturityCompensation`  
應收付金額 | `NetSettlementAmount`  
可轉債形式 | `ConvertibleBondType`  
備註 | `Remarks`  

---
英文欄位	Access 資料型態	備註
TradeID	Short Text (255)	成交單號
ContractID	Short Text (255)	契約編號
TerminationTradeID	Short Text (255)	解約成交單號
TradeDate	Date/Time	交易日期
SettlementDate	Date/Time	交割日期
UnderlyingCode	Short Text (50)	標的代號
UnderlyingName	Short Text (255)	標的名稱
CustomerName	Short Text (255)	客戶名稱
ConvertibleBondMaturityDate	Date/Time	可轉債到期日
NotionalPrincipal	Currency	名目本金
RepurchasePrincipal	Currency	買回本金
SwapRate	Number (Double)	交換利率 (%)
TerminationRate	Number (Double)	解約利率 (%)
InterestStartDate	Date/Time	計息起日
InterestEndDate	Date/Time	計息迄日
InterestAmount	Currency	利息
Penalty	Currency	違約金
MaturityCompensation	Currency	到期補償金
NetSettlementAmount	Currency	應收付金額
ConvertibleBondType	Short Text (255)	可轉債形式
Remarks	Long Text	備註 (多行)

---

英文欄位	Access 資料型態	備註
Text
Text
Text
Date
Date
Text
Text
Text
Date
Currency
Currency
Double
Double
Date
Date
Currency
Currency
Currency
Currency
Text
Text

---
TradeID
ContractID
TerminationTradeID
TradeDate
SettlementDate
UnderlyingCode
UnderlyingName
CustomerName
ConvertibleBondMaturityDate
NotionalPrincipal
RepurchasePrincipal
SwapRate
TerminationRate
InterestStartDate
InterestEndDate
InterestAmount
Penalty
MaturityCompensation
NetSettlementAmount
ConvertibleBondType
Remarks


---


成交單號
契約編號
解約成交單號
交易日期
交割日期
標的代號
標的名稱
客戶名稱
可轉債到期日
名目本金
買回本金
交換利率(%)
解約利率(%)
計息起日
計息迄日
利息
違約金
到期補償金
應收付金額
可轉債形式
備註

## 💵 票券交易明細表 BillTransactionDetails
| 中文欄位 | 英文命名 |
|----------|-----------|
交易日期 | `TradeDate`  
交割日期 | `SettlementDate`  
成交單編號 | `TradeID`  
帳務目的 | `AccountingPurpose`  
交易對手 | `Counterparty`  
統一編號 | `TaxID`  
交易員 | `Trader`  
覆核層級 | `ReviewLevel`  
交易別 | `TransactionType`  
訊息類別 | `MessageType`  
票券批號-子號 | `BillBatchSubNumber`  
票類 | `BillType`  
發票人 | `Issuer`  
保證人 | `Guarantor`  
統一編號 | `IssuerTaxID`, `GuarantorTaxID`, `CounterpartyTaxID`  
發票日 | `IssueDate`  
到期日 | `MaturityDate`  
面額 | `FaceValue`  
萬元單價/票載利率 | `PricePerTenThousandOrCoupon`  
成交利率 | `TradeYield`  
天數 | `Days`  
成交金額 | `TradeAmount`  
帳上成本 | `BookCost`  
利息累計稅款 | `AccruedTax`  
利息金額 | `InterestAmount`  
約定到期金額 | `AgreedMaturityAmount`  
約定到期日 | `AgreedMaturityDate`  
損益 | `ProfitLoss`  
帳上稅款 | `BookTax`  

---

英文欄位	Access 資料型態	備註
TradeDate	Date/Time	交易日期
SettlementDate	Date/Time	交割日期
TradeID	Short Text (255)	成交單編號
AccountingPurpose	Short Text (255)	帳務目的
Counterparty	Short Text (255)	交易對手
TaxID	Short Text (50)	統一編號
Trader	Short Text (255)	交易員
ReviewLevel	Number (Integer)	覆核層級
TransactionType	Short Text (255)	交易別
MessageType	Short Text (255)	訊息類別
BillBatchSubNumber	Short Text (100)	票券批號-子號
BillType	Short Text (50)	票類
Issuer	Short Text (255)	發票人
Guarantor	Short Text (255)	保證人
IssuerTaxID, GuarantorTaxID, CounterpartyTaxID	Short Text (50) each	各方統一編號
IssueDate	Date/Time	發票日
MaturityDate	Date/Time	到期日
FaceValue	Currency	面額
PricePerTenThousandOrCoupon	Number (Double)	萬元單價/票載利率
TradeYield	Number (Double)	成交利率 %
Days	Number (Integer)	天數
TradeAmount	Currency	成交金額
BookCost	Currency	帳上成本
AccruedTax	Currency	利息累計稅款
InterestAmount	Currency	利息金額
AgreedMaturityAmount	Currency	約定到期金額
AgreedMaturityDate	Date/Time	約定到期日
ProfitLoss	Currency	損益
BookTax	Currency	帳上稅款

---

英文欄位	Access 資料型態	備註
Date
Date
Text
Text
Text
Text
Text
Text
Text
Text
Text
Text
Text
Text
Text
Date
Date
Currency
Double
Double
Number
Currency
Currency
Currency
Currency
Currency
Date
Currency
Currency

---

TradeDate
SettlementDate
TradeID
AccountingPurpose
Counterparty
TaxID
Trader
ReviewLevel
TransactionType
MessageType
BillBatchSubNumber
BillType
Issuer
Guarantor
IssuerTaxID`, `GuarantorTaxID`, `CounterpartyTaxID
IssueDate
MaturityDate
FaceValue
PricePerTenThousandOrCoupon
TradeYield
Days
TradeAmount
BookCost
AccruedTax
InterestAmount
AgreedMaturityAmount
AgreedMaturityDate
ProfitLoss
BookTax


---


交易日期
交割日期
成交單編號
帳務目的
交易對手
統一編號
交易員
覆核層級
交易別
訊息類別
票券批號-子號
票類
發票人
保證人
統一編號
發票日
到期日
面額
萬元單價/票載利率
成交利率
天數
成交金額
帳上成本
利息累計稅款
利息金額
約定到期金額
約定到期日
損益
帳上稅款

## 📋 債券評價表 for 匯出檔案 BondValuationExport
| 中文欄位 | 英文命名 |
|----------|-----------|
多/空部位 | `PositionType` (`LongShortFlag`)  
部門別 | `Department`  
帳務日 | `AccountingDate`  
計提日 | `AccrualDate`  
債券名稱 | `BondName`  
債券代碼 | `BondCode`  
帳務目的 | `AccountingPurpose`  
債券產品別 | `BondProductType`  
發票日 | `IssueDate`  
到期日 | `MaturityDate`  
票載利率% | `CouponRate`  
首買利率% | `InitialPurchaseRate`  
市場價格(利率%) | `MarketRate`  
庫存面額 | `InventoryFaceValue`  
帳上成本 | `BookCost`  
庫存市價 | `MarketValue`  
應計利息 | `AccruedInterest`  
放空帳上利息 | `ShortInterest`  
應付代收稅款 | `WithholdingTax`  
放空帳上稅款 | `ShortTax`  
評價損益 | `ValuationGainLoss`  
放空成本 | `ShortCost`  

---

英文欄位	Access 資料型態	備註
PositionType (LongShortFlag)	Short Text (50)	多/空部位
Department	Short Text (255)	部門別
AccountingDate	Date/Time	帳務日
AccrualDate	Date/Time	計提日
BondName	Short Text (255)	債券名稱
BondCode	Short Text (50)	債券代碼
AccountingPurpose	Short Text (255)	帳務目的
BondProductType	Short Text (255)	債券產品別
IssueDate	Date/Time	發票日
MaturityDate	Date/Time	到期日
CouponRate	Number (Double)	票載利率 %
InitialPurchaseRate	Number (Double)	首買利率 %
MarketRate	Number (Double)	市場價格 (利率 %)
InventoryFaceValue	Currency	庫存面額
BookCost	Currency	帳上成本
MarketValue	Currency	庫存市價
AccruedInterest	Currency	應計利息
ShortInterest	Currency	放空帳上利息
WithholdingTax	Currency	應付代收稅款
ShortTax	Currency	放空帳上稅款
ValuationGainLoss	Currency	評價損益
ShortCost	Currency	放空成本

---
英文欄位	Access 資料型態	備註
Text
Text
Date
Date
Text
Text
Text
Text
Date
Date
Double
Double
Double
Currency
Currency
Currency
Currency
Currency
Currency
Currency
Currency
Currency

---

PositionType` (`LongShortFlag
Department
AccountingDate
AccrualDate
BondName
BondCode
AccountingPurpose
BondProductType
IssueDate
MaturityDate
CouponRate
InitialPurchaseRate
MarketRate
InventoryFaceValue
BookCost
MarketValue
AccruedInterest
ShortInterest
WithholdingTax
ShortTax
ValuationGainLoss
ShortCost

---

多/空部位
部門別
帳務日
計提日
債券名稱
債券代碼
帳務目的
債券產品別
發票日
到期日
票載利率%
首買利率%
市場價格(利率%)
庫存面額
帳上成本
庫存市價
應計利息
放空帳上利息
應付代收稅款
放空帳上稅款
評價損益
放空成本 

## 🧮 票券評價表 for 匯出檔案 BillValuationExport
| 中文欄位 | 英文命名 |
|----------|-----------|
部門別 | `Department`  
評價日 | `ValuationDate`  
帳務目的 | `AccountingPurpose`  
票別 | `BillType`  
票券批號 | `BillBatchNumber`  
發票人 | `Issuer`  
庫存面額 | `InventoryFaceValue`  
發票日 | `IssueDate`  
到期日 | `MaturityDate`  
交割日 | `SettlementDate`  
成交利率 | `TradeYield`  
市場利率 | `MarketYield`  
萬元單價 | `PricePerTenThousand`  
距到期日天數 | `DaysToMaturity`  
庫存成本 | `InventoryCost`  
現值 | `PresentValue`  
市價 | `MarketValue`  
評價損益 | `ValuationGainLoss`  

---


英文欄位	Access 資料型態	備註
Department	Short Text (255)	部門別
ValuationDate	Date/Time	評價日
AccountingPurpose	Short Text (255)	帳務目的
BillType	Short Text (50)	票別
BillBatchNumber	Short Text (100)	票券批號
Issuer	Short Text (255)	發票人
InventoryFaceValue	Currency	庫存面額
IssueDate	Date/Time	發票日
MaturityDate	Date/Time	到期日
SettlementDate	Date/Time	交割日
TradeYield	Number (Double)	成交利率 %
MarketYield	Number (Double)	市場利率 %
PricePerTenThousand	Number (Double)	萬元單價
DaysToMaturity	Number (Integer)	距到期日天數
InventoryCost	Currency	庫存成本
PresentValue	Currency	現值
MarketValue	Currency	市價
ValuationGainLoss	Currency	評價損益

---
英文欄位	Access 資料型態	備註
Text
Date
Text
Text
Text
Text
Currency
Date
Date
Date
Double
Double
Double
Number
Currency
Currency
Currency
Currency

---

Department
ValuationDate
AccountingPurpose
BillType
BillBatchNumber
Issuer
InventoryFaceValue
IssueDate
MaturityDate 
SettlementDate
TradeYield
MarketYield
PricePerTenThousand
DaysToMaturity
InventoryCost
PresentValue
MarketValue
ValuationGainLoss

---


部門別
評價日
帳務目的
票別
票券批號
發票人
庫存面額
發票日
到期日
交割日
成交利率
市場利率
萬元單價
距到期日天數
庫存成本
現值
市價
評價損益

## 🧾 票券庫存明細 BillInventoryDetails
| 中文欄位 | 英文命名 |
|----------|-----------|
組織架構 | `OrganizationStructure`  
帳務目的 | `AccountingPurpose`  
庫存類別 | `InventoryType`  
票券批號-子號 | `BillBatchSubNumber`  
票類 | `BillType`  
發票人 | `Issuer`  
保證人 | `Guarantor`  
交易對手 | `Counterparty`  
發票人統一編號 | `IssuerTaxID`  
保證人統一編號 | `GuarantorTaxID`  
交易對手統一編號 | `CounterpartyTaxID`  
發票日 | `IssueDate`  
到期日 | `MaturityDate`  
買入日期 | `PurchaseDate`  
RS到期日 | `RSMaturityDate`  
距到期日天數 | `DaysToMaturity`  
原購利率 | `OriginalPurchaseRate`  
萬元單價票載利率 | `PricePerTenThousandOrCoupon`  
帳上面額 | `BookFaceValue`  
RP在外面額 | `OutstandingRPAmount`  
帳上成本 | `BookCost`  
帳上稅款 | `BookTax`  
面額（一） | `FaceValue1`  
張數（一） | `Units1`  
面額（二） | `FaceValue2`  
張數（二） | `Units2`  
面額（三） | `FaceValue3`  
張數（三） | `Units3`  
稅後實得額 | `NetAfterTaxAmount`  
免稅款 | `TaxExemptAmount`  
收益率% | `Yield`  

---

英文欄位	Access 資料型態	備註
OrganizationStructure	Short Text (255)	組織架構
AccountingPurpose	Short Text (255)	帳務目的
InventoryType	Short Text (50)	庫存類別
BillBatchSubNumber	Short Text (100)	票券批號-子號
BillType	Short Text (50)	票類
Issuer	Short Text (255)	發票人
Guarantor	Short Text (255)	保證人
Counterparty	Short Text (255)	交易對手
IssuerTaxID	Short Text (50)	發票人統一編號
GuarantorTaxID	Short Text (50)	保證人統一編號
CounterpartyTaxID	Short Text (50)	交易對手統一編號
IssueDate	Date/Time	發票日
MaturityDate	Date/Time	到期日
PurchaseDate	Date/Time	買入日期
RSMaturityDate	Date/Time	RS 到期日
DaysToMaturity	Number (Integer)	距到期日天數
OriginalPurchaseRate	Number (Double)	原購利率 %
PricePerTenThousandOrCoupon	Number (Double)	萬元單價/票載利率
BookFaceValue	Currency	帳上面額
OutstandingRPAmount	Currency	RP 在外面額
BookCost	Currency	帳上成本
BookTax	Currency	帳上稅款
FaceValue1, FaceValue2, FaceValue3	Currency each	面額（一/二/三）
Units1, Units2, Units3	Number (Integer)	張數（一/二/三）
NetAfterTaxAmount	Currency	稅後實得額
TaxExemptAmount	Currency	免稅款
Yield	Number (Double)	收益率 %

---
英文欄位	Access 資料型態	備註
Text
Text
Text
Text
Text
Text
Text
Text
Text
Text
Text
Date
Date
Date
Date
Number
Double
Double
Currency
Currency
Currency
Currency
Currency
Number
Currency
Currency
Double

---

OrganizationStructure
AccountingPurpose
InventoryType
BillBatchSubNumber
BillType
Issuer
Guarantor
Counterparty
IssuerTaxID
GuarantorTaxID
CounterpartyTaxID
IssueDate
MaturityDate
PurchaseDate
RSMaturityDate
DaysToMaturity
OriginalPurchaseRate
PricePerTenThousandOrCoupon
BookFaceValue
OutstandingRPAmount
BookCost
BookTax
FaceValue1
Units1
FaceValue2
Units2
FaceValue3
Units3
NetAfterTaxAmount
TaxExemptAmount
Yield

---

組織架構
帳務目的
庫存類別
票券批號-子號
票類
發票人
保證人
交易對手
發票人統一編號
保證人統一編號
交易對手統一編號
發票日
到期日
買入日期
RS到期日
距到期日天數
原購利率
萬元單價票載利率
帳上面額
RP在外面額
帳上成本
帳上稅款
面額（一）
張數（一）
面額（二）
張數（二）
面額（三）
張數（三）
稅後實得額
免稅款
收益率%



票券交易明細表
E欄位有空格要刪除

票券庫存明細表
H欄位有空格要刪除

債券交易明細表
統一編號 F

債券風險部位餘額
NO

CBAS BUY
NO

CBAS SELL
NO













說明整個運作流程

原始批次、資通、外幣債等檔案在哪邊，資料呈現如何。

原始檔案放在Relation\RawData\項下，更改檔名為正確格式。

開啟Access的Configuration Form表單，填入ReportDataDate(輸入原始報表資料日期，日期格式為 YYYY/MM/DD，例如: 2025/2/27)，按Tab後ReportMonth會自動填入為月底日期，RawFilePath和CopyFilePath為原始檔案路徑和複製檔案路徑(複製檔案路徑中的檔案後續會清理為正規格式)

填妥後滑鼠點擊更新Configuration，會在Configuration資料表中加入一筆設定資料(後續打開Access之後預設值會抓取這筆資料內容)，並將RawFilePath中的檔案複製一份到CopyFilePath

點擊【上傳至Access DB】後即開始批次清洗Excel檔案為正規格式並匯入Access資料庫中。

匯入CloseRate，進入CloseRate Form，填入關帳匯率日期及完整檔案路徑，點擊【Update Close Rate】

資料表欄位統整資料Excel檔，以及批次建立資料表腳本，相關欄位名稱及資料類型等可以參考，如果有需要再調整也可以參考。







- 每月匯入資料庫(Access)使用流程(Access檔案路徑為:\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\DB_MonthlyReport.accdb):
    - 下載製作報表所需原始資料:批次檔案、資通檔案、外幣債評估表、減損表(檔案置放於共用資料夾路徑:\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\RelationData\RawData)。
    - 修改檔案名稱為符合固定格式:報表名稱_yyyymmdd.副檔名(其中yyyymmdd為月底最後一工作日)(資通下載報表及以民國命名報表須修改)。
    - 開啟共用資料夾Access檔案(預設檔案開啟頁面為Configuration表單，可於Access左側找到該表單)，預設可輸入ReportDataDate或ReportMonth切換方式:
        1. ReportDataDate: 輸入原始資料日期yyyy/mm/dd，即月底工作日(例如:2025/02/27)(需與流程2檔名年月日一致)，點擊Tab鍵ReportMonth會自動填入月底日期。
        2. ReportMonth: 輸入年月份yyyy/mm(例如:2025/02)，點擊Tab鍵ReportDataDate會自動填入月底工作日。
    補充:(1)後續篩選資料時間區別係以DataMonthString欄位判斷。
    - 填妥後點擊【更新Configuration】按鈕，此時會在Configuration資料表建立一筆設定資料(開啟Configuration表單預設填入資料會抓取最新那筆設定資料)，並將RawFilePath中的檔案複製一份到CopyFilePath(RawFilePath檔案不會異動，CopyFilePath檔案會清理為正規格式)。
    - 先行確認FilePath資料表包含所有要匯入之報表，後點擊【上傳至Access DB】即開始批次清洗Excel檔案為正規格式並匯入Access資料庫中，出現執行完畢視窗後，確認資料庫表單資料是否匯入。

- 關帳匯率CloseRate匯入資料庫使用流程:
    - 點擊Access左側【CloseRate表單】，輸入要匯入之關帳匯率西元年月日及完整檔案路徑後，點擊【Update Close Rate】，視窗顯示XXXX即更新完成。

- 前端製作報表(Excel)使用流程(Excel檔案路徑為:\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\RelationData\RawData\MonthlyReport\ControlDB_Panel.xlsm):
    - 複製共用資料夾路徑中的MonthlyReport資料夾一份於本地端，開啟Excel檔案，切換分頁為Controll分頁，點擊執行按鈕，於視窗輸入申報西元年月份(月份需補0，例如:2025/01)。
    - EmptyReport資料夾為空白申報檔，OutputReport資料夾為執行後產生之申報檔；Control頁面為相關配置設定，包含Access資料庫檔名、空白申報檔路徑等等；一個分頁代表申報的一個檔案，執行後會抓取Access已建立查找報表(Query)顯示於Excel，紀錄相關需填報分頁及欄位名稱與填報數值，後將相關填報報表名稱、欄位及其值新增至【MonthlyDeclarationReport】資料表中。


- 比較重要的表
BankDirectory: 外幣部分有關拆款報表，需填寫銀行Code，由原始報表中的SWIFT CODE關聯拉取資料。
DBsColTable: 彙整所有資料表及欄位資訊的表，在匯入資料庫前需拉取這邊資料資料表欄位名稱(順序會影響資料表是否正確匯入，IMPORT程序會檢核Excel Row1名稱是否和資料表欄位名稱一致)。
FilePath: 設定需匯入資料庫之報表名稱、檔名及副檔名。
FXDebtReferCountry: 為FM13所建立資料表，因為原始資料table沒辦法篩出外幣債是哪個國家發行，這邊只能手動填入。


- Log檔案路徑、
快速帶一下前端的原始碼，帶一下如果後續要改要怎麼改，例如有新增欄位Excel欄位設定、class欄位設定、修改Process報表名稱內容、有必要時修改Access Query；例如新增報表: Excel增加分頁、Class新增Case設定欄位、Array增加報表、新增Process報表名稱Sub(報表製作處理邏輯)、新增Access Query。等等...資料庫匯入那邊要在繼續寫

- 彙整的對照表(比較Account Code寫死在SQL中和使用另一張表集中管理差異比較，才不會改死人，又容易漏掉；優點有一個人集中管理可以同時修改所有相關表單資料，可能比較熟的經辦直接修改那邊的資料，那個對照關係可能關連到不同人的表單，有好有壞，有些人習慣都自己控，這樣可能會對自己手上表單有沒辦法掌控的異動。我現在就是先整理一個版本，可能也會有錯，還需要更多討論)

- 有寫好建立資料表的腳本


- 待辦事項:
    1. 修改台幣債資料庫匯入事項
    2. 統一使用統整表SQL問題
    3. 共用Function拉取彙整報表資料、關帳匯率
    4. 確認修改後的SQL沒問題
    5. 確認再那台主機上可以使用
    6. 確認整個流程沒問題，Access放在共用資料夾，Excel放在本地端
    7. 有關彙整表，是否可以建立一個介面，現有選單是下拉式選單，選取之後可以建立SQL可否達成，或是方便觀看也可以(可能中間還有一堆判斷篩選邏輯)；新增項目的功能。可能有些人習慣自己用自己的，或是統一由比較有經驗的人去控管
    8. 針對控管部分一個是沒遇過的就直接擋掉；另一種是設定增幅控，同樣是否可以製作一個表單去控制
    9. 台幣Import尚未處理 餘額D copypaste問題以及統一編號問題和check資料是否都有處理過
