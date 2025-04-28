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
    - (已嘗試改善) 解決開啟多個Excel檔案時會出錯的問題
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
    - 減損後端匯入的時候，會漏掉 AC債務工具投資-公債-中央政府(? 這張表沒有匯入，需將已經匯入的資料庫重新跑一次
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
    1. 建立台幣Access資料表
        1.1 完成DB column mapping
        1.2 清洗資料表資料，欄位%調整和餘額欄位整併，另外還有利率敏感性分析的表
    2.1 建立AccountCode Pair ReportType
        2.1.1 建立好配對表，取得原始配對方式之後，分別為不同報表需求建立Query(兩種方式，一種是建立一個帶入參數的共用Query，參數會由前端代碼輸入，另一種是為不同的表建立不同的Query，就後續維護角度是共用參數較為理想)
    2.2 check 前端資料取得函數Paramater
    2.3 修改外幣SQL，可能要增加傳入參數
    3.1 台幣製作報表邏輯
    3.2 台幣有些欄位資料需要人工輸入，可能放在Setting頁面，VBA操控Control頁面清除必要填入欄位資料，程式碼在剛運行時，檢核相關欄位是否填入，如果沒有填入則強制中止Sub運行
    3.3


    處理外幣債表ValuationType欄位，寫一個for迴圈，然後看是不是要寫一個dictionary去替換名稱
    減損的表也需要處理，寫for迴圈去逐一處理
    至於要使用的主要名稱以餘額C名稱為主


    外幣需要處理的報表
    AI602
    FB1
    FM10
    FM11
    










FVPL_GovBond_Foreign
FVPL_GovBond_Foreign
FVPL_CompanyBond_Foreign
FVPL_CompanyBond_Foreign
FVPL_FinancialBond_Foreign
FVOCI_GovBond_Foreign
FVOCI_CompanyBond_Foreign
FVOCI_CompanyBond_Foreign
FVOCI_FinancialBond_Foreign
AC_GovBond_Foreign
AC_CompanyBond_Foreign
AC_CompanyBond_Foreign
AC_FinancialBond_Foreign


外幣債表
AH

減損
L
