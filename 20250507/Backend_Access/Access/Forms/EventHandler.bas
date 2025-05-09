Option Compare Database

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
        WriteLog "Configuration資料表ReportDataDAte欄位資料格式錯誤: ReportDataDate不是日期格式" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
    End If
End Sub

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
        WriteLog "Configuration資料表ReportDataDAte欄位資料格式錯誤: ReportDataDate不是日期格式" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
        Cancel = True
    End If
End Sub

'TextBox event AfterUpdate handler for ReportDataDate
Private Sub inputReportDataDate_AfterUpdate()
    Dim reportDataDate As Date
    reportDataDate = Me.inputReportDataDate

    ' Check inputReportDataDate is Date
    If IsDate(reportDataDate) Then Me.inputReportMonth = DateSerial(Year(reportDataDate), Month(reportDataDate) + 1, 0)
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
        WriteLog "Configuration資料表ReportDataDAte欄位資料格式錯誤: ReportDataDate不是日期格式" & vbCrLf & "(日期格式為 YYYY/MM/DD，例如: 2025/02/27)"
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
        WriteLog "ReportDataDate 更新為申報月份月底工作日:" & vbCrLf & "         " & Format(lastWorkday, "yyyy/mm/dd")
    Else
        MsgBox "lastWorkday資料取得錯誤，請更新ReportDataDate", vbExclamation
        WriteLog "lastWorkday資料取得錯誤，請更新ReportDataDate"
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
    WriteLog "Update data successfully！"

    db.Close
    Set db = Nothing

    call CopyRawFileData(rawFilePath, copyFilePath)
End Sub




Private Sub btnCleanAndImportData_Click()
    call ProcessAllReports
End Sub
