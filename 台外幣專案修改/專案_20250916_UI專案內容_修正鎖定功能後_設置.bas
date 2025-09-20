' ====
' Temp
' ====



好的 — 我把 完整專案規劃（含所有操作細節與觸發方式） 整理成一份可直接照做的操作手冊。
內容包含：必要前置、資料夾/檔案/CSV 範例、VBA 模組放哪、UI（Excel）版面與每個按鈕要綁哪個 macro、執行順序（使用者從開啟、選報表、編輯、儲存、執行報表一整套流程）、自動啟動設定、除錯紀錄位置，以及 Dir（萬用字元）行為說明。文末還有快速驗證步驟供你立刻測試。

我把共用部分（版本 A / B 共用）先列出，
然後針對兩個版本的差異（右側是否新增 Report tab）說明要如何設定與觸發。請照順序完成設定並測試。

⸻

一、專案總覽（快速看圖）
	•	設定來源：以 CSV 放在 config\ 資料夾（可版本控制）。
	•	Excel UI：工作簿包含設定表（tblReports、tblUpdateSheet、tblExportPDF、Mappings）與操作介面 UI_Main。
	•	使用流程（使用者）：開啟 Workbook → UI 初始化（InitializeUI）→ 左側選 Report → 右側檢視 → 點 Edit 進入編輯模式 → 編輯右側（只允許右側範圍）→ Save 或 Cancel → 若要跑報表按 Run All（或單一報表 Run）→ 檢查 logs\RunLog_YYYYMMDD.txt。
	•	程式觸發方式：按鈕 (Assign Macro)、或 Workbook_Open 自動呼叫 InitializeUI、或直接在 VBA 按 F5 執行宏（開發測試）。

⸻

二、必要前置（請先做）
	1.	建立資料夾（放在 Excel 檔同一個資料夾）：

<workbook_path>\
  config\
  config\backup\
  logs\
  SAVE_PDF\

（EnsureConfigFolders 程式也會自動建立，手動建立可先行）

	2.	必要工作表（sheet 名稱非常重要，請確保名稱完全相符）：
	•	tblReports（主設定）
	•	tblUpdateSheet
	•	tblExportPDF
	•	Mappings
	•	UI_Main（使用者介面）
	•	另：你原本的處理程式所需之輔助函式、模板檔、範本檔要放好（如 GetMonths 等 helper 函式必須存在）
	3.	將我之前給你的三個 Module（ConfigIO, Helpers, UIHandlers）貼到 VBA 編輯器。
	•	版本 A：使用 UIHandlers（A 版本的 code）。
	•	版本 B：使用 UIHandlers（B 版本的 code，含 Report tab）。
	4.	在 ThisWorkbook 的 code（Workbook 物件）加入 Workbook_Open 呼叫（下方有程式片段）以自動初始化（可選）。
	5.	確保 ThisWorkbook 有命名範圍 YearMonth（若你用到時間自動填 / ProcessAllReports 會需要）。

⸻

三、CSV 格式（最小可運作範例）

把這些範例存在 config\ 資料夾：

config\tblReports.csv（一行 header，至少有 ReportID）

ReportID,TplPathPattern,TplPathTimeFormat,DeclPathPattern,DeclPathTimeFormat,IsDeleteTplPattern,HeaderTimeSheetRange,HeaderTimeFormat,ProcessingMacro,PDFParentFolder,ReportType
F1_F2,F1F2\F1F2_NEW_v巨集底稿YYYYMM.xlsm,NEW_AD_YYYYMM,F1F2\申報_vYYYYMM.xlsm,NEW_AD_YYYYMM,TRUE,Table1!A2|Table2!A5,NEW_AD_YYYYMM,Module1.MainSub,,月報

config\tblUpdateSheet.csv

ReportID,UpdateSheet,ClearRange,ImportPathPattern,ImportSheets,ImportProcessType,ImportPathTimeFormat,FilterSpec
F1_F2,F1,"A1:Z500","批次報表\*DBU_DL6850*","Sheet1","Copy",NEW_AD_YYYYMM,
F1_F2,F2,"A1:Z500","F1F2\DBU_CM2810_31_AC_E_YYYYMM.xls","Sheet1","Copy",NEW_AD_YYYYMM,交易別:首購;續發|票類:CP1;CP2

config\tblExportPDF.csv

ReportID,PDFSheets,ParentFolder
F1_F2,"F1,F2",

config\Mappings.csv

ReportID,SrcSheet,SrcRange,DstSheet,DstRange
F1_F2,F1,A1:C20,申報表1,A1:C20

注意：FilterSpec 字串使用範例格式 欄位:選項1;選項2|欄位2:選項A;選項B（程式會需要解析，若需我可加入 parser）。

⸻

四、VBA 模組與放置位置（要貼入哪裡）
	1.	在 VBA 編輯器（Alt+F11） → 插入 Module，貼入 ConfigIO（含 LoadAllCSVToSheets, LoadCSVToSheet, SaveSheetToCSV, ParseCSVLine 等）。
	2.	新增 Module，貼入 Helpers（含 LogInfo/LogError/MkDirRecursive）。
	3.	新增 Module，貼入對應版本的 UIHandlers（版本 A 或 B）。
	4.	UI_Main 工作表代碼：右鍵 UI_Main → 檢視程式碼，貼上 Worksheet_SelectionChange 事件（我在前面兩版本都給了代碼）。
	5.	（可選）在 ThisWorkbook 的程式碼加：

Private Sub Workbook_Open()
    On Error Resume Next
    InitializeUI
End Sub

（若不想自動啟動，可不加入；手動執行 InitializeUI 就行）

⸻

五、UI（Excel 介面）詳細佈局與按鈕（建議格位與 Macro 對應）

我用 UI_Main 的格子定位，方便你用 Excel 插入 shape 並對齊。

建議格位（可手動調整）
	•	左側（報表清單）：
	•	範圍 A1:C20（或 A1:C100）：放 tblReports 的 header 與 data（RefreshLeftPanel 會把 tblReports 複製到這）
	•	右側（顯示區）：
	•	起點 E3 開始：header 和資料會被貼到 E3（RefreshRightPanel 利用 E3）
	•	Navbar（右側上方）：
	•	在 UI_Main 放 4 個 Shape（矩形）當作 tab（版本 A 為 3 個）；放在 E1:H1（或你喜歡的位置）
	•	按鈕放置（在 UI_Main）
	•	Edit（編輯）→ 建議放在 B22，Assign Macro: EnterEditMode
	•	Save（儲存）→ 建議放在 C22，Assign Macro: SaveEdits
	•	Cancel（取消）→ 建議放在 D22，Assign Macro: CancelEdits
	•	Run All（執行全部報表）→ 建議放在 A22（或右上），Assign Macro: RunAll_FromUI（如需我可貼此 wrapper；下方有說明）
	•	Nav 按鈕 macro 對應：
	•	Nav_ShowUpdateSheet -> 連結到「UpdateSheet」tab
	•	Nav_ShowExportPDF -> ExportPDF
	•	Nav_ShowMappings -> Mappings
	•	（版本 B）Nav_ShowReport -> Report

如何新增並 assign macro
	1.	Excel 插入 → 形狀 → 選矩形 → 在 UI_Main 位置放置。
	2.	右鍵形狀 → 指派巨集 → 選擇對應 macro（例如 Nav_ShowUpdateSheet）。
	3.	重複為其他按鈕 assign macro。

⸻

六、執行流程（使用者一步步操作說明）

以下為使用者從開啟到執行報表完整步驟（包含觸發點）：

A. 初次啟動（第一次設定時）
	1.	將工作簿與 config 資料夾及 CSV 檔放到同一個資料夾。
	2.	在 VBA 中貼上三個 Module（ConfigIO、Helpers、UIHandlers）與 UI_Main 的 SelectionChange 事件。
	3.	（可選）在 ThisWorkbook 添加 Workbook_Open 呼叫 InitializeUI，以便開啟檔案自動初始化。
	4.	開啟 Excel 檔案或在 VBA 執行 InitializeUI（F5） — 這會：
	•	讀取 config\*.csv 並把資料貼到相應 sheet（tblReports 等）。
	•	在 UI_Main 顯示左側 tblReports（A1）且右側預設為空或 header（E3）。
	•	將設定 sheet 及 UI sheet 設為保護（不可編輯），右側 staging 區亦鎖定（只能用 Edit 解鎖）。

Trigger: InitializeUI（手動或 Workbook_Open 自動）

B. 日常操作（選、編輯、儲存）
	1.	在 UI_Main 左側（A 列）點選某筆 ReportID → Worksheet_SelectionChange 事件會觸發 → 內部會呼叫 RefreshRightPanel(reportID, currentActiveTab)。
	•	Trigger: Worksheet_SelectionChange → RefreshRightPanel
	2.	右側會顯示對應的 tblUpdateSheet / tblExportPDF / Mappings（取決於你目前在哪個 navbar tab）。
	3.	若要修改設定，按 Edit 按鈕（Assign: EnterEditMode）：
	•	程式會把 currentInEditMode = True，解除鎖定 僅右側 staging 範圍，讓使用者編輯。
	•	Trigger: EnterEditMode
	4.	編輯完成後按 Save（Assign: SaveEdits）：
	•	程式會驗證（ReportID 一致等），把右側編輯結果寫回對應底層 sheet（tblUpdateSheet / tblExportPDF / Mappings /（版本 B）tblReports）。
	•	接著呼叫 SaveSheetToCSV 存回 config\*.csv（會做備份）。
	•	最後重新鎖定 UI 並 reload 所有 CSV（LoadAllCSVToSheets）。
	•	Trigger: SaveEdits → SaveSheetToCSV → LoadAllCSVToSheets
	5.	若按 Cancel（Assign: CancelEdits）：
	•	採取不存檔、重新載入右側（把右側範圍回復為儲存檔），解除編輯模式。
	•	Trigger: CancelEdits

C. 執行報表（Run）
	•	你可以新增一個 Run All 按鈕（形狀），Assign macro（建議兩種做法）：
	1.	直接呼叫原本你已有的 ProcessAllReports_New（若你已整合我 earlier 的 ProcessAllReports_New 版本）。
	•	Trigger: ProcessAllReports_New
	2.	建議寫一個小 wrapper RunAll_FromUI（放於 ConfigIO 或 UIHandlers）：

Public Sub RunAll_FromUI()
    ' 先把 UI 的修改寫回 CSV（確保保存）
    LoadAllCSVToSheets  ' 確保讀取最新
    ' 可 optional: Save currently open UI edits
    ' 執行處理
    ProcessAllReports_New
End Sub

	•	如果你希望在執行前自動儲存設定，可在 wrapper 中先呼叫 SaveSheetToCSV（或強制呼叫 SaveEdits 流程）。

D. 單一報表測試（Run single）
	•	我之前提供 Run_Test_F1F2 作為測試，assign 一個按鈕 Run_Test_F1F2 即可執行單筆測試（通常會用 F1_F2）。

⸻

七、Startup 自動化（建議）

把下列程式放到 ThisWorkbook：

Private Sub Workbook_Open()
    On Error Resume Next
    InitializeUI
End Sub

這樣使用者開啟檔案就會自動載入 CSV 並初始化 UI。若不想自動，可以把它註解或不要加入。

⸻

八、錯誤與除錯（常見問題與處理）
	1.	沒有找到工作表（Subscript out of range）
	•	檢查 tblReports、tblUpdateSheet、tblExportPDF、Mappings、UI_Main 是否存在且拼字完全一致（大小寫不敏感，但空格或錯字會造成問題）。
	2.	CSV 無法載入 / 空表
	•	檢查 config\*.csv 是否存在，欄位分隔符是否為逗號（CSV），編碼建議 UTF-8 with BOM。執行 LoadAllCSVToSheets 可在即時偵錯。
	3.	拒絕存檔（SaveSheetToCSV 失敗）
	•	可能是檔案被其他程序鎖定（或沒有寫入權限），檢查 config\ 是否可以寫入。
	4.	Run 時找不到某個範本檔或輸出路徑
	•	ProcessAllReports 會以 ThisWorkbook.Path 為 base，檢查 TplPathPattern 的相對路徑是否正確（例如 F1F2\...）以及那些檔案是否存在。
	5.	查看 log：<workbook_path>\logs\RunLog_YYYYMMDD.txt（常見錯誤會記在這裡）。
	6.	如果有 Runtime Error，請用 VBA 即時監看（Debug -> 堆疊）並把錯誤訊息貼給我，我可以幫你定位原因。

⸻

九、Dir 與萬用字元（*、?）行為說明（簡明）
	•	Dir(pathPattern) 用來回傳符合 pathPattern 的第一個檔名（不含路徑）。
	•	範例：Dir("C:\MyFolder\*cm2610*") 會回傳第一個檔名，像 DBU_CM2610_202506.xls。
	•	呼叫範例一般用法：

Dim f As String
f = Dir(basePath & "\" & pattern) ' pattern 可能含 * 或 ?
If f <> "" Then
    ' f 是找到的第一個檔名 (ex: "DBU_CM2610_202506.xls")
    fullPath = basePath & "\" & f
End If


	•	若要取得所有符合項目，使用迴圈：

f = Dir(basePath & "\" & pattern)
Do While f <> ""
    Debug.Print f
    f = Dir() ' 連續呼叫 Dir() 回傳下一個符合項
Loop


	•	在我們的專案中，ResolveImportFiles 已實作：當路徑 pattern 含 * 時會呼叫 Dir 抓到第一個相符的檔名，並把 pattern 中的檔名片段替換為實際檔名（只匹配第一個，這正是你之前要求的行為）。

⸻

十、快速驗證（5 分鐘上手）
	1.	把我給的 Module（ConfigIO,Helpers,UIHandlers）貼上，並新增 UI_Main sheet。
	2.	把範例 CSV 放到 config\。
	3.	在 ThisWorkbook 加 Workbook_Open 呼叫 InitializeUI（或在 VBA 直接執行 InitializeUI）。
	4.	開啟 UI_Main，左側選 F1_F2，右側會顯示子表資料（E3 開始）。
	5.	按 Edit（EnterEditMode），在右側修改一個值（例如 ImportPathPattern），按 Save（SaveEdits）。
	6.	檢查 config\tblUpdateSheet.csv 是否已變更（有備份在 config\backup\）。
	7.	點 Run All（或直接執行 ProcessAllReports_New），觀察 logs\RunLog_YYYYMMDD.txt 產生紀錄；若有 PDF 輸出，會在 SAVE_PDF\{newMon}\{ReportID}\。

⸻

十一、版本 A / B 差異回顧（觸發點）
	•	兩版本的觸發方式完全相同（因為 UIHandlers 提供同名宏）：
	•	InitializeUI（啟動）
	•	Worksheet_SelectionChange（左側選取觸發）
	•	EnterEditMode, SaveEdits, CancelEdits（編輯流程）
	•	Nav_ShowUpdateSheet, Nav_ShowExportPDF, Nav_ShowMappings（切換 tab）
	•	版本 B 多 Nav_ShowReport（新增 tab），User 點 Report tab 會顯示 tblReports 的該筆 row；Save 時會把 row 覆寫回 tblReports.csv。
	•	ProcessAllReports_New（或 RunAll_FromUI wrapper）為執行報表的觸發 macro。

⸻

十二、我建議你現在做的次序（實作順序）
	1.	在測試用資料夾建立 config\，並放上示例 CSV。
	2.	貼入 ConfigIO、Helpers。執行 LoadAllCSVToSheets 檢查 CSV 載入正確。
	3.	選 A 或 B 版本，把相應 UIHandlers 貼入，並把 UI_Main 的 Worksheet_SelectionChange 貼上。
	4.	在 UI_Main 放好按鈕並 assign macro（照上面對應表）。
	5.	執行 InitializeUI，驗證 UI 顯示與鎖定狀態。
	6.	試 Edit / Save / Cancel。
	7.	試 Run（ProcessAllReports_New 或 RunAll_FromUI），檢查 logs 與 PDF 目錄。

⸻

如果你要我現在幫你做下一步，我可以立刻（選一項）：
	1.	把 Workbook_Open 與 RunAll_FromUI wrapper 的完整程式碼段貼給你（方便直接貼入 ThisWorkbook）。
	2.	或把 FilterSpec 字串 parser 加到 ImportCsvWithFilter 的呼叫流程中（把 FilterSpec 轉成 filterFields/filterValues）。
	3.	或直接把 UI_Main 的「按鈕建置配置圖（格子位置與精確大小）」貼出，讓你可以直接用 Excel 建置。

你選其中一項我就接著立即貼出相對應的程式碼/圖示。




' =====
' NavBar 設置
' =====


好問題 👍
前面專案規劃我有寫到「右方的關聯資料表會用 NavBar 呈現，點擊切換會刷新顯示不同表格」，但確實沒有說清楚 怎麼觸發切換。我來把這部分補完整，並列出實作方式。

⸻

NavBar 切換機制設計

1. NavBar 呈現方式
	•	建議做法：在 UI 頁面 (例如 UI_Main) 左上方放一排 表單控制元件的按鈕 (Form Control Button)，每個按鈕對應一張關聯表：
	•	btnUpdateSheet → 切換到 tblUpdateSheet
	•	btnExportPDF → 切換到 tblExportPDF
	•	btnMappings → 切換到 Mappings
這些按鈕會像 NavBar 一樣，排成水平或垂直選單。
	•	替代做法：用一個 下拉式清單 (Data Validation List) 來選擇當前 Focus 表，VBA 在 Worksheet_Change 偵測選項，刷新顯示右側的內容。

⸻

2. 事件觸發與程式碼呼叫
	•	Form Control Button 的情境
每個按鈕綁定對應的 VBA 巨集，例如：

Sub ShowUpdateSheet()
    Call SwitchNav("tblUpdateSheet")
End Sub

Sub ShowExportPDF()
    Call SwitchNav("tblExportPDF")
End Sub

Sub ShowMappings()
    Call SwitchNav("Mappings")
End Sub


	•	SwitchNav 的主控程式

Sub SwitchNav(targetTable As String)
    ' 更新全域變數，記錄當前 focus 的表
    gCurrentNav = targetTable

    ' 如果目前有選擇 tblReports 的某筆資料
    If gCurrentReportID <> "" Then
        ' 重新載入對應資料表內容到右側顯示區
        Call RefreshRightPane(gCurrentReportID, gCurrentNav)
    End If
End Sub


	•	RefreshRightPane

Sub RefreshRightPane(reportID As String, navTable As String)
    Select Case navTable
        Case "tblUpdateSheet"
            Call LoadUpdateSheetUI(reportID)
        Case "tblExportPDF"
            Call LoadExportPDFUI(reportID)
        Case "Mappings"
            Call LoadMappingsUI(reportID)
    End Select
End Sub



⸻

3. UI 操作流程
	1.	使用者點選左側的某筆 Report (例如 ReportA) → gCurrentReportID = "ReportA"
預設右側顯示 tblUpdateSheet (或任何你定義的預設 Nav)。
	2.	使用者切換 NavBar (例如按下「ExportPDF」按鈕) → 觸發 SwitchNav("tblExportPDF")。
	3.	系統檢查目前 gCurrentReportID 是否有值，有的話 → 重新載入右側對應的表格內容。
	4.	如果使用者尚未啟動 編輯模式 → 所有右側表格鎖定為唯讀狀態。
	5.	若使用者點選「開啟編輯模式」 → 解除右側該表格的保護，允許編輯。

⸻

4. 是否需要額外元件？
	•	需要：至少要有 Form Control 按鈕（或 ActiveX 按鈕、或下拉清單）。
	•	不建議：完全依靠 Worksheet_SelectionChange 來判斷，因為那樣會太難用（例如使用者誤點儲存格就觸發切換）。

⸻

要不要我幫你把這個 NavBar 設計整合進 A 版 / B 版完整專案流程，並標註哪邊是「新增元件」和「綁定程式」的細節？



' ==============
' 有關矩形按鈕設置
' ==============

好 👌 我來解釋清楚：

1. 為什麼用 Shape（矩形）當 Tab

在 Excel 裡，Shape（圖形）（例如矩形、圓角矩形）可以放在工作表上當作視覺化的 Tab / 按鈕。
	•	這樣 UI 會比較像「頁籤切換」效果。
	•	可以用不同顏色表示目前選中的 Tab（例如選中變藍色，其他是灰色）。
	•	每個 Shape 都可以綁定一個 VBA 巨集 → 點擊時觸發程式。

這就是我說的在 UI_Main 放四個 Shape，分別代表：
	•	tblUpdateSheet
	•	tblExportPDF
	•	Mappings
	•	tblReports（額外視窗用，B 版才會用到）

⸻

2. 怎麼新增 Shape（矩形）當按鈕
	1.	到 UI_Main 分頁 → 功能區選 插入 → 圖形 → 矩形。
	2.	拉出四個矩形，排成水平（像 NavBar）或垂直（像側邊欄）。
	3.	在每個矩形上輸入對應文字，例如「UpdateSheet」「ExportPDF」。
	4.	選擇矩形 → 右鍵 → 指派巨集 → 選一個 VBA Sub（例如 ShowUpdateSheet）。

⸻

3. VBA 端程式結構

你需要寫四個小 Sub，分別綁到四個矩形：

Sub ShowUpdateSheet()
    Call SwitchNav("tblUpdateSheet")
End Sub

Sub ShowExportPDF()
    Call SwitchNav("tblExportPDF")
End Sub

Sub ShowMappings()
    Call SwitchNav("Mappings")
End Sub

Sub ShowReports()
    Call SwitchNav("tblReports")
End Sub

再搭配前面說的 SwitchNav 主控程式：

Sub SwitchNav(targetTable As String)
    gCurrentNav = targetTable
    
    ' 更新 Tab 外觀（顏色變化）
    Call UpdateNavBarUI(targetTable)
    
    ' 刷新右側內容
    If gCurrentReportID <> "" Then
        Call RefreshRightPane(gCurrentReportID, gCurrentNav)
    End If
End Sub


⸻

4. Tab 外觀切換效果（選中變色）

你可以在 SwitchNav 內呼叫 UpdateNavBarUI，改變 Shape 填滿顏色：

Sub UpdateNavBarUI(activeTab As String)
    Dim s As Shape
    With Worksheets("UI_Main")
        ' Reset all tabs to灰色
        .Shapes("Tab_UpdateSheet").Fill.ForeColor.RGB = RGB(200, 200, 200)
        .Shapes("Tab_ExportPDF").Fill.ForeColor.RGB = RGB(200, 200, 200)
        .Shapes("Tab_Mappings").Fill.ForeColor.RGB = RGB(200, 200, 200)
        .Shapes("Tab_Reports").Fill.ForeColor.RGB = RGB(200, 200, 200)
        
        ' Active one to藍色
        Select Case activeTab
            Case "tblUpdateSheet"
                .Shapes("Tab_UpdateSheet").Fill.ForeColor.RGB = RGB(0, 112, 192)
            Case "tblExportPDF"
                .Shapes("Tab_ExportPDF").Fill.ForeColor.RGB = RGB(0, 112, 192)
            Case "Mappings"
                .Shapes("Tab_Mappings").Fill.ForeColor.RGB = RGB(0, 112, 192)
            Case "tblReports"
                .Shapes("Tab_Reports").Fill.ForeColor.RGB = RGB(0, 112, 192)
        End Select
    End With
End Sub


⸻

✅ 這樣你就有一個「自製 NavBar」：
	•	四個矩形 = 四個 Tab
	•	點擊矩形 → 觸發巨集 → 切換右側內容 + 更新顏色

⸻

要不要我幫你把這個「Shape NavBar 設計」完整整合進 A 版專案流程，直接標示哪些是 新增的 UI 物件、哪些是 程式碼新增點？



' ===========
' 矩形按鈕美工
' ===========

可以的 👍
雖然 Excel 的矩形（Shape）預設滑鼠移上去顯示的是「十字游標」，但只要你 有對它指派巨集，它就能當作按鈕使用。

你會看到的情況
	•	預設：游標移到 Shape 上 → 出現十字游標（因為是圖形物件）。
	•	一旦 右鍵 → 指派巨集 → 綁定到某個 Sub，
那麼 點擊 Shape 就會觸發巨集。

👉 換句話說，游標樣式不影響功能，點下去就能執行程式。

⸻

讓它更有「按鈕感」的方法

如果你希望更像「真正的按鈕」，可以：
	1.	格式化矩形：加上漸層或陰影，看起來像立體按鈕。
	2.	滑鼠移入效果：
	•	沒有原生 hover 事件，但你可以設計兩個矩形（普通態、選中態），切換可見性。
	•	或者簡單點，僅在「選中」的 Tab 改顏色，其餘維持灰色。
	3.	改用 ActiveX 或表單控制項：
	•	如果要游標變成「手型 🖐️」，那就用「表單控制項 → 按鈕」或 ActiveX 控制項。
	•	但這樣 UI 就比較生硬，不像 NavBar。

⸻

所以結論：
✔️ 矩形絕對可以當作按鈕用，只是游標樣式是 Excel 的預設，點擊後會觸發你指定的巨集。

要不要我幫你寫一個簡單的 完整範例檔架構（含 4 個矩形 → 切換顏色 → 刷新右側內容），讓你可以直接貼進 Excel 測試？




' =======
' 測試編碼
' =======

好 — 我直接給你一段可貼到新 Module 的測試程式，會把同一個 CSV 用幾種常見編碼（UTF-8、Big5、Windows-1252）讀入，並把每種讀取結果放在同一張工作表的不同欄，方便你肉眼比對哪個編碼結果是正確的（快速定位是哪種編碼 / 讀法造成亂碼）。

請按以下步驟操作：
	1.	開啟你的 Excel，按 Alt+F11 → Insert → Module，把下面整段 VBA 貼上。
	2.	存好你的測試 CSV（任何路徑都行）。
	3.	在 VBA 中執行 Test_ReadCSV_Compare（F5），或回到 Excel 執行巨集 Test_ReadCSV_Compare。
	4.	會跳出檔案選擇視窗，選你的 CSV。完成後檢視工作表 TestEncodings：每一欄代表一種編碼的讀取結果（UTF-8 / Big5 / Windows-1252），比對哪一欄中文字正確。

Option Explicit

' 測試用：用不同 Charset 用 ADODB.Stream 讀同一檔案，並把結果放在 TestEncodings 表中（便於比對）
' 支援的編碼順序可調：預設嘗試 "utf-8", "big5", "windows-1252"

Public Sub Test_ReadCSV_Compare()
    Dim f As Variant
    f = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV to test")
    If VarType(f) = vbBoolean And f = False Then
        MsgBox "已取消", vbInformation
        Exit Sub
    End If
    Dim filePath As String: filePath = CStr(f)

    Dim charsets As Variant
    charsets = Array("utf-8", "big5", "windows-1252") ' 你可以在這裡新增或改順序

    Dim results As Object: Set results = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(charsets) To UBound(charsets)
        Dim cs As String: cs = charsets(i)
        Dim lines As Variant
        lines = ReadTextFileByCharset_SplitLines(filePath, cs)
        results.Add cs, lines
    Next i

    WriteResultsToSheet results, filePath

    MsgBox "讀檔完成。請打開工作表 'TestEncodings' 比對各欄結果。", vbInformation
End Sub

' 以指定 charset 讀整個檔案，回傳行陣列（1-based zero-like variant）
Public Function ReadTextFileByCharset_SplitLines(filePath As String, charset As String) As Variant
    Dim stm As Object
    On Error GoTo ErrHandler
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Mode = 3
    stm.Charset = charset
    stm.Open
    stm.LoadFromFile filePath
    Dim txt As String: txt = stm.ReadText
    stm.Close
    Set stm = Nothing

    ' Normalize line endings -> 使用 vbLf 為分隔依據
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    ' 如果檔案以 BOM 開頭 (某些情況)，有時 ADODB 會自動處理，但保險起見：
    If Len(txt) > 0 Then
        If AscW(Left$(txt, 1)) = 65279 Then ' U+FEFF BOM
            txt = Mid$(txt, 2)
        End If
    End If

    Dim arr As Variant
    If txt = vbNullString Then
        ReDim arr(0 To 0)
        arr(0) = ""
    Else
        arr = Split(txt, vbLf)
    End If

    ReadTextFileByCharset_SplitLines = arr
    Exit Function
ErrHandler:
    ' 若 ADODB 不可用或讀取錯誤，回傳一個包含錯誤訊息的陣列
    Dim errArr(0 To 0) As String
    errArr(0) = "ERROR reading with charset='" & charset & "': " & Err.Number & " " & Err.Description
    ReadTextFileByCharset_SplitLines = errArr
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    Set stm = Nothing
End Function

' 把 dictionary 的每個 charset 結果寫到工作表 TestEncodings（每個 charset 一欄）
Private Sub WriteResultsToSheet(results As Object, sourceFilePath As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TestEncodings")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "TestEncodings"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ws.Range("A1").Value = "Source file:"
    ws.Range("B1").Value = sourceFilePath

    Dim col As Long: col = 1
    Dim keys As Variant: keys = results.Keys
    Dim maxRows As Long: maxRows = 0
    Dim k As Long
    For k = 0 To results.Count - 1
        Dim cs As String: cs = keys(k)
        ws.Cells(2, col).Value = cs
        Dim arr As Variant: arr = results(cs)
        Dim r As Long
        For r = LBound(arr) To UBound(arr)
            ws.Cells(r + 3, col).Value = arr(r)
        Next r
        If UBound(arr) - LBound(arr) + 1 > maxRows Then maxRows = UBound(arr) - LBound(arr) + 1
        col = col + 1
    Next k

    ' 自動調整寬度、凍結窗格
    ws.Columns("A:Z").AutoFit
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True

    ' 在第一列加提示
    ws.Cells(1, 4).Value = "說明：每一欄為用該 charset 讀出的結果。比對哪一欄中文正確。"
End Sub


⸻

補充說明（看完會更快找到問題）
	•	若「UTF-8 欄」的中文正確，表示 CSV 實際是 UTF-8 → 建議後續用 ADODB.Stream 用 charset = "utf-8" 來讀即可。
	•	若「Big5 欄」中文才正常，表示 CSV 實際是 Big5（在繁體 Windows 常見），你可把 InitializeUI 改以 big5 去讀，或把所有 CSV 轉成 UTF-8。
	•	若三欄都亂，可能代表 CSV 在建立時就有資料被破壞（例如含非文字二進位），或檔案是由其他系統用不同分隔符或控制碼產生（需進一步檢查檔頭）。
	•	ADODB.Stream 在大部分 Windows/Office 環境可用；若出錯（例：系統沒有 MDAC/ADODB），上述函式會回傳錯誤訊息在表格中（可看到第一欄會顯示 ERROR）。

⸻

如果你跑完測試並把 TestEncodings 的截圖或描述貼給我，我可以直接告訴你下一步該怎麼修改 InitializeUI（例如改成固定用 UTF-8，或改成嘗試 Big5 fallback，或直接用 Open For Binary + 手動轉碼）。
