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
    - 有可能程序問題是ImporterStandard那邊的xlapp導致的，可以試看看




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





1. 要不要使用msgbox，還是只剩下log就好
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

    2. 前端產生報表，




這是說在家裡面做的
- AccountCodeMap 修改原來 AssetMeasurementSubType 為 AssetMeasurementType 和 增加 GroupFlag 欄位
- 已經改好外幣修改SQL
- x還沒改好相對應修正前端代碼，可能要稍微核對一下名稱和欄位有沒有變動
- 有關台幣清理部分已經完成









每月匯入資料庫(Access)使用流程(Access檔案路徑為:):
1.下載製作報表所需原始資料:批次檔案、資通檔案、外幣債評估表、減損表(檔案置放於共用資料夾路徑:Relation\RawData\)。
2.修改檔案名稱為符合固定格式:報表名稱_yyyymmdd.副檔名(其中yyyymmdd為月底最後一工作日)(資通下載報表及以民國命名報表須修改)。
3.開啟共用資料夾Access檔案(預設檔案開啟頁面為Configuration表單，可於Access左側找到該表單)，預設可輸入ReportDataDate或ReportMonth切換方式:
    a.ReportDataDate: 輸入原始資料日期yyyy/mm/dd，即月底工作日(例如:2025/02/27)(需與流程2檔名年月日一致)，點擊Tab鍵ReportMonth會自動填入月底日期。
    b.ReportMonth: 輸入年月份yyyy/mm(例如:2025/02)，點擊Tab鍵ReportDataDate會自動填入月底工作日。
補充:(1)後續篩選資料時間區別係以DataMonthString欄位判斷。
4.填妥後點擊【更新Configuration】按鈕，此時會在Configuration資料表建立一筆設定資料(開啟Configuration表單預設填入資料會抓取最新那筆設定資料)，並將RawFilePath中的檔案複製一份到CopyFilePath(RawFilePath檔案不會異動，CopyFilePath檔案會清理為正規格式)。
5.先行確認FilePath資料表包含所有要匯入之報表，後點擊【上傳至Access DB】即開始批次清洗Excel檔案為正規格式並匯入Access資料庫中，出現執行完畢視窗後，確認資料庫表單資料是否匯入。

關帳匯率CloseRate匯入資料庫使用流程:
1.點擊Access左側【CloseRate表單】，輸入要匯入之關帳匯率西元年月日及完整檔案路徑後，點擊【Update Close Rate】，視窗顯示XXXX即更新完成。

前端製作報表(Excel)使用流程(Excel檔案路徑為:):
1.複製共用資料夾路徑中的XX檔案一份於本地端，開啟Excel檔案，切換分頁為Controll分頁，點擊執行按鈕，於視窗輸入申報西元年月份(月份需補0，例如:2025/01)。
2.EmptyReport資料夾為空白申報檔，OutputReport資料夾為執行後產生之申報檔；Control頁面為相關配置設定，包含Access資料庫檔名、空白申報檔路徑等等；一個分頁代表申報的一個檔案，執行後會抓取Access已建立查找報表(Query)顯示於Excel，紀錄相關需填報分頁及欄位名稱與填報數值，後將相關填報報表名稱、欄位及其值新增至【MonthlyDeclarationReport】資料表中。

UI部分只是方便先這樣配置沒有多想，有更好的想法可以再修改

資料檢核部分還沒做

還需要更多測試

比較重要的表
BankDirectory: 外幣部分有關拆款報表，需填寫銀行Code，由原始報表中的SWIFT CODE關聯拉取資料。
DBsColTable: 彙整所有資料表及欄位資訊的表，在匯入資料庫前需拉取這邊資料資料表欄位名稱(順序會影響資料表是否正確匯入，IMPORT程序會檢核Excel Row1名稱是否和資料表欄位名稱一致)。
FilePath: 設定需匯入資料庫之報表名稱、檔名及副檔名。
FXDebtReferCountry: 為FM13所建立資料表，因為原始資料table沒辦法篩出外幣債是哪個國家發行，這邊只能手動填入。


Log檔案路徑、
快速帶一下前端的原始碼，帶一下如果後續要改要怎麼改，例如有新增欄位Excel欄位設定、class欄位設定、修改Process報表名稱內容、有必要時修改Access Query；例如新增報表: Excel增加分頁、Class新增Case設定欄位、Array增加報表、新增Process報表名稱Sub(報表製作處理邏輯)、新增Access Query。等等...資料庫匯入那邊要在繼續寫

彙整的對照表(比較Account Code寫死在SQL中和使用另一張表集中管理差異比較，才不會改死人，又容易漏掉；優點有一個人集中管理可以同時修改所有相關表單資料，可能比較熟的經辦直接修改那邊的資料，那個對照關係可能關連到不同人的表單，有好有壞，有些人習慣都自己控，這樣可能會對自己手上表單有沒辦法掌控的異動。我現在就是先整理一個版本，可能也會有錯，還需要更多討論)



待辦事項:
1.修改台幣債資料庫匯入事項
2.統一使用統整表SQL問題
3.共用Function拉取彙整報表資料、關帳匯率
4.確認修改後的SQL沒問題
5.確認再那台主機上可以使用
6.確認整個流程沒問題，Access放在共用資料夾，Excel放在本地端
7.有關彙整表，是否可以建立一個介面，現有選單是下拉式選單，選取之後可以建立SQL可否達成，或是方便觀看也可以(可能中間還有一堆判斷篩選邏輯)；新增項目的功能。可能有些人習慣自己用自己的，或是統一由比較有經驗的人去控管
8.針對控管部分一個是沒遇過的就直接擋掉；另一種是設定增幅控，同樣是否可以製作一個表單去控制
9.有關檢核一個月報表區分不同次產生，可以新增一個index，每次去抓取之前當年月的最大那個index再+1