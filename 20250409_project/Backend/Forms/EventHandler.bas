Private Sub Form_Load()
    Dim rs As DAO.Recordset
    Dim sqlString As String
    Dim lastWorkday As Date
    Dim reportDataDate As Date

    ' 查詢最新一筆資料
    sqlString = "SELECT TOP 1 * FROM Configuration ORDER BY DataID DESC"
    Set rs = CurrentDb.OpenRecordset(sqlString)

    If Not rs.EOF Then
        Me.inputReportDataDate = rs!ReportDataDate
        reportDataDate = rs!ReportDataDate
        Me.inputReportMonth = rs!ReportMonth
        Me.inputRawFilePath = rs!RawFilePath
        Me.inputCopyFilePath = rs!CopyFilePath
    End If
    rs.Close
    Set rs = Nothing

    Me.inputReportMonth.Enabled = False

    'Get LastWorkday
    lastWorkday = GetLastWorkday(Year(Date), Month(Date)-1)

    If IsDate(reportDataDate) Then
        If DateAdd("m", 1, DateSerial(Year(reportDataDate), Month(reportDataDate), 1)) < DateSerial(Year(Date), Month(Date), 1) Then
            Me.lblCheckUpdateLabel.Caption = "請更新ReportDataDate(報表資料日期)" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
        Else
            Me.lblCheckUpdateLabel.Caption = " "
        End If

        ' 驗證 ReportDataDate 是否為該月的月底工作日
        If reportDataDate <> lastWorkday Then
            Me.lblCheckReportDateLabel.Caption = "最近申報時間" & Year(Date) & "年" & (Month(Date)-1) & "月之月底工作日為 " & Format(lastWorkday, "mm/dd/yyyy")
        Else
            Me.lblCheckReportDateLabel.Caption = " "
        End If
    Else
        MsgBox "Configuration資料表ReportDataDAte欄位資料格式錯誤: ReportDataDate不是日期格式" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)", vbExclamation
    End If
End Sub

' Compute Last Workday of each month
Public Function GetLastWorkday(ByVal yr As Integer, ByVal mth As Integer) As Date
    Dim lastDay As Date
    Dim isHoliday As Boolean
    Dim rs As DAO.Recordset
    Dim sqlString As String

    'Set first day of month
    lastDay = DateSerial(yr, mth + 1, 0)
    
    Do    
        isHoliday = False
        sqlString = "SELECT COUNT(Date) AS cnt FROM Holidays WHERE Date = #" & lastDay & "#"
        Set rs = CurrentDb.OpenRecordset(sqlString)
        If Not rs.EOF Then isHoliday = (rs!cnt > 0)
        rs.Close
        Set rs = Nothing
        'If is holiday then day + 1
        If isHoliday Then lastDay = lastDay - 1
    Loop While isHoliday  ' Is holiday then continue loop
    
    GetLastWorkday = lastDay
End Function

'TextBox event beforeUpdate handler for ReportDataDate
Private Sub inputReportDataDate_BeforeUpdate(Cancel As Integer)
    Dim reportDataDate As Date
    Dim lastWorkday As Date

    reportDataDate = Me.inputReportDataDate
    lastWorkday = GetLastWorkday(Year(Date), Month(Date)-1)

    ' Check inputReportDataDate is Date
    If IsDate(reportDataDate) Then
        If reportDataDate > Date Then
            Me.lblCheckUpdateLabel.Caption = "輸入ReportDataDate(報表資料日期)無資料，請輸入有效日期" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
            Cancel = True
        Else
            Me.lblCheckUpdateLabel.Caption = " "
        End If

        ' 驗證 ReportDataDate 是否為該月的月底工作日
        If reportDataDate <> lastWorkday Then
            Me.lblCheckReportDateLabel.Caption = "最近申報時間" & Year(Date) & "年" & (Month(Date)-1) & "月之月底工作日為 " & Format(lastWorkday, "mm/dd/yyyy")
        Else
            Me.lblCheckReportDateLabel.Caption = " "
        End If
    Else
        MsgBox "Configuration資料表ReportDataDAte欄位資料格式錯誤: ReportDataDate不是日期格式" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)", vbExclamation
        Cancel = True
    End If
End Sub

'TextBox event AfterUpdate handler for ReportDataDate
Private Sub inputReportDataDate_AfterUpdate()
    Dim reportDataDate As Date
    reportDataDate = Me.inputReportDataDate

    ' Check inputReportDataDate is Date
    If IsDate(reportDataDate) Then Me.inputReportMonth = GetLastWorkday(Year(reportDataDate), (Month(reportDataDate)-1))
End Sub

'SWitch Date(ReportData/Month) Mode
'Button event click handler for ReportMonth
Private Sub btnSWitchDateTextBox_Click()
    Me.inputReportMonth.Enabled = Not Me.inputReportMonth.Enabled
    Me.inputReportDataDate.Enabled = Not Me.inputReportDataDate.Enabled
End Sub

'TextBox event beforeUpdate handler for ReportMonth
Private Sub inputReportMonth_BeforeUpdate(Cancel As Integer)
    Dim dt As Date
    Dim lastWorkday As Date

    dt = Me.inputReportMonth
    lastWorkday = GetLastWorkday(Year(Date), Month(Date)-1)

    ' Check inputReportMonth is Date
    If IsDate(dt) Then 
        If dt > Date Then
            Me.lblCheckUpdateLabel.Caption = "輸入ReportDataDate(報表資料日期)無資料，請輸入有效日期" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
            Cancel = True
        Else
            Me.lblCheckUpdateLabel.Caption = " "
        End If

        ' 驗證 ReportDataDate 是否為該月的月底工作日
        If dt <> lastWorkday Then
            Me.lblCheckReportDateLabel.Caption = "最近申報時間" & Year(Date) & "年" & (Month(Date)-1) & "月之月底工作日為 " & Format(lastWorkday, "mm/dd/yyyy")
        Else
            Me.lblCheckReportDateLabel.Caption = " "
        End If
    Else
        MsgBox "Configuration資料表ReportDataDAte欄位資料格式錯誤: ReportDataDate不是日期格式" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub inputReportMonth_AfterUpdate()
    Dim dt As Date
    Dim lastWorkday As Date

    dt = Me.inputReportMonth
    lastWorkday = GetLastWorkday(Year(dt), Month(dt))
    
    ' Check Date
    If IsDate(dt) And IsDate(lastWorkday) Then
        Me.inputReportDataDate = lastWorkday
        Me.inputReportMonth = DateSerial(Year(dt), Month(dt) + 1, 0)
        MsgBox "ReportDataDate 更新為申報月份月底工作日:" & vbCrLf & "         " & Format(lastWorkday, "yyyy/mm/dd")
    Else
        MsgBox "lastWorkday資料取得錯誤，請更新ReportDataDate", vbExclamation
    End If
End Sub

Private Sub btnSaveConfiguration_Click()
    Dim db As DAO.Database
    Dim sqlString As String
    Dim reportDataDate As Date, reportMonth As Date
    Dim reportMonthString As String
    Dim rawFilePath As String, copyFilePath As String

    Set db = CurrentDb
    reportDataDate = Me.inputReportDataDate
    reportMonth = Me.inputReportMonth
    reportMonthString = Format(reportMonth, "yyyy/mm")
    rawFilePath = Me.inputRawFilePath
    copyFilePath = Me.inputCopyFilePath


    ' 更新資料表
    sqlString = "INSERT INTO Configuration (ReportDataDate, ReportMonth, ReportMonthString, RawFilePath, CopyFilePath, CaseCreatedAt) " & _
          "VALUES (#" & reportDataDate & "#, #" & reportMonth & "#, '" & reportMonthString & "', '" & rawFilePath & "', '" & copyFilePath & "', Now());"

    db.Execute sqlString, dbFailOnError
    MsgBox "Update data successfully！", vbInformation

    db.Close
    Set db = Nothing

    call CopyRawFileData(rawFilePath, copyFilePath)
End Sub




Private Sub btnCleanAndImportData_Click()
    call ProcessAllReports
End Sub

'--------
'參考資料 ex chatgpt說明
'--------
' ### **步驟詳解**
' 在 Access 中建立表單（例如命名為 `frmConfiguration`），並添加以下控件：
' - **TextBox 控件**：
'     - `inputReportMonth`（綁定 `ReportMonth`）
'     - `inputReportDataDate`（綁定 `ReportDataDate`）
'     - `inputRawFilePath`（綁定 `RawFilePath`）
'     - `inputCopyFilePath`（綁定 `CopyFilePath`）
' - **Label 控件**：
'     - `lblCheckUpdateLabel`（顯示 ReportMonth 的檢查結果）
'     - `lblCheckReportDateLabel`（顯示 ReportDataDate 的檢查結果）
' - **Button 控件**：
'     - `btnSave`（儲存按鈕）

' ### **2. 設定表單啟動時自動開啟**

' 1. **將表單 `frmConfiguration` 設定為 Access 開啟時的預設表單**
'     - 在「檔案」->「選項」->「目前資料庫」->「顯示表單」-> 選擇 `frmConfiguration`。
' 2. **載入最新的 CaseCreatedAt 資料**
'     - 在 `frmConfiguration` 的 **On Load** 事件中撰寫 VBA 代碼：
' ### **1. 建立表單 UI**
' OK


'btn save click時，需要處理report month資料，看要當作string or date處理，如果當作date處理，則在前端將資料送出時，需要處理日期資料，或是在sql中要另外處理
' ### **4. 更新資料**

' 當用戶修改 `inputReportMonth`、`inputReportDataDate`、`inputRawFilePath`、`inputCopyFilePath` 後，按下 `btnSave` 按鈕時，將資料存入 `Configuration` 表：
'ok

' ### **5. 總結**

' 1. **表單 UI**
'     - 建立 `frmConfiguration` 表單，包含輸入框、標籤、按鈕。
' 2. **自動載入最新資料**
'     - `Form_Load()` 讀取 `Configuration` 表最新資料，並檢查 `ReportMonth` 和 `ReportDataDate`。
' 3. **計算第一個工作日**
'     - `GetLastWorkday()` 函式考慮周末和國定假日，計算當月第一個工作日。
' 4. **更新資料**
'     - `btnSave_Click()` 按鈕將修改後的資料儲存至 `Configuration` 表。

' 這樣你的 Access 報表自動化專案就能按照需求運作了！有需要調整的部分可以再討論。 🚀
