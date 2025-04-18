- 前端
    - 每次生成報表access資料表中沒有一個獨立可以驗證是哪次生成的欄位，造成儲存的資料無法驗證是哪次生成的資料
    - (已修正) Button Click後執行程序
    - reportName使用class直接取得，不要再宣告新的參數
    - (已修正) 測試dataMonthString用input輸入方式
    - (已修正) FM2 FBA3 有空白欄位無法顯示 
    - (已修正) 前端以日期變數輸入抓access資料
    - log想說要不要存入資料庫 (需要再想一想要怎麼做)
    - (已修正) 前端加入log
    - (已修正) FM10還要加入AC
    - (已修正) CloseRate要篩選出DataDate的最後一天，FM2和表2
    - (已修正) F1F2表需要Create出要讓F2表填的欄位
    
    - 修改 Access SQL為以param方式抓取資料 修改到ai240
    - mm4901b swift 沒篩完全，剩下2024年開始之後貼，CTBC 中信swift code要再確認是不是不只一個
    - f1_f2 還要修正 bug
    - Think 建立一個Swift Table Primary key is swift code for FM2 Report



- 其他項目，例如後端
    - (已修正) 整理資料庫匯入程式碼
    - (已嘗試改善) 解決開啟多個Excel檔案時會出錯的問題
    - (已修正) 後端增加 CloseRate
    - (已修正) 加入處理 F2 後端
    - (已修正) 加入處理 F2 前端
    - (已完成) 完成剩下後端資料彙入 (run 20240329的OBU_AC5601時，在.txt檔的250行應該是EUR，但是資料錯誤文字是 CNY，導致BUG)
    - 匯入這個月的資料庫跑看看有沒有問題
    - (已修正) 將CloseRate併入後端相同檔案
    - (已修正) 核對Swiftcode
    - 建立 CloseRate Func
    - 測試在共用路徑能不能用
    - (已修正) 建立完整Log管理機制
    - 加入錯誤處理
    - (已修正) 清理欄位時，有些資料表會新增空白資料








