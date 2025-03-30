

'-----------
'|Question|
'-----------

' Insert Data to Access DB
Public Sub InsertIntoTable(ByVal DBPath As String, _
                           ByVal tableName As String, _
                           ByVal dataMonthString As String, _
                           ByVal reportTitle As String, _
                           ByVal fieldCode As String, _
                           ByVal content As Double, _
                           ByVal description As String)
    Dim conn As Object
    Dim cmd As Object

    ' On Error GoTo ErrHandler

    ' 建立 ADODB 連線
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    ' 建立 SQL 指令
    Dim sql As String
    sql = "INSERT INTO MonthlyDeclarationReport (DataMonthString, ReportTitle, FieldCode, Content, Description, CaseCreatedAt) " & _
    "VALUES ('" & dataMonthString & "', '" & reportTitle & "', '" & fieldCode & "', '" & content & "', '" & description & "', Now());"


    ' 執行 SQL
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Execute

    ' 關閉連線
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Exit Sub

    ' ErrHandler:
    '     MsgBox "發生錯誤: " & Err.Description, vbCritical
    '     If Not conn Is Nothing Then conn.Close
    '     Set conn = Nothing
    '     Set cmd = Nothing
End Sub

' Check Input Format fir "yyyy/mm"
Public Function IsValidDataMonth(ByVal userInput As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "^\d{4}/(0[1-9]|1[0-2])$"
        .IgnoreCase = True
        .Global = False
    End With
    
    IsValidDataMonth = regex.Test(Trim(userInput))
End Function

Public Function GetAccessDataAsArray(ByVal DBPath As String, _
                                     ByVal QueryName As String, _
                                     Optional ByVal dataMonthString As String = vbNullString) As Variant
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim dataArr As Variant
    Dim colCount As Integer, rowCount As Integer
    Dim headerArr() As String
    
    On Error GoTo ErrHandler

    ' Build ADODB Connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath

    ' Build Command Object
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = QueryName
    cmd.CommandType = 4 ' adCmdStoredProc (表示這是儲存查詢)

    ' If dataMonthString, then pass to access
    If dataMonthString <> vbNullString Then
        cmd.Parameters.Append cmd.CreateParameter("DataMonthParam", 200, 1, 255, dataMonthString)
    End If

    ' Execute, use adUseClient to support GetRows()
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient（讓 GetRows 可用）
    rs.Open cmd

    ' Make sure Recordset is not null
    If rs Is Nothing Or rs.EOF Then
        MsgBox "查詢結果為空，請檢查資料庫與查詢條件。", vbExclamation
        GetAccessDataAsArray = Array()
        Exit Function
    End If
    ' Fetch Columns Name
    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    ' Save Column to headerArr
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next i
    ' Fetch data by GetRows()
    dataArr = rs.GetRows()
    rowCount = UBound(dataArr, 2) + 1 ' 資料列數
    ' Create resultArr()
    Dim resultArr() As Variant
    ReDim resultArr(0 To rowCount, 0 To colCount - 1)
    ' Set Column
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next i
    ' Set data
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next j
    Next i
    ' Close connection
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    GetAccessDataAsArray = resultArr
    Exit Function
ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
    GetAccessDataAsArray = Array()
End Function
Option Explicit
Public dataMonthString As String
Public Sub RunTotal()
    Dim isInputValid As Boolean
    isInputValid = False
    ' via InputBox for Users to enter yyyy/mm
    Do
        dataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(dataMonthString) Then
            isInputValid = True
        ElseIf Trim(dataMonthString) = "" Then
            MsgBox "請輸入需取得報表年度/月份" & vbCrLf & "(例如: 2024/01 )", vbExclamation, "Nul Error"
        Else
            MsgBox "輸入格式有誤，請輸入正確格式(yyyy/mm)" & vbCrLf & "(例如: 2024/01 )", vbExclamation, "Format Error"
        End If
    Loop Until isInputValid
    Call CNY1
End Sub
Public Sub CNY1()
    Dim controlSheet As Worksheet
    Dim xlsht As Worksheet

    Dim DBsPath As String
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    Dim dataArr As Variant

    Dim reportTitle As String
    Dim queryTable As String
    
    Dim rngs As Range
    Dim rng As Range

    Dim fxReceive As Double
    Dim fxPay As Double

    reportTitle = "CNY1"
    queryTable = "CNY1_DBU_AC5601"

    Set controlSheet = ThisWorkbook.Sheets("ControlPanel")

    ' 取得 Access 資料庫路徑
    DBsPath = ThisWorkbook.Path & "\" & controlSheet.Range("DBsPathFileName")

    '***Set xlsht = ThisWorkbook.Sheets("CNY1")中的CNY1要用一個字串，當作參數傳入
    ' 設定目標 Excel 分頁
    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    xlsht.Range("A:E").ClearContents

    dataArr = GetAccessDataAsArray(DBsPath, queryTable)
    
    For j = 0 To UBound(dataArr, 2)
        For i = 0 To UBound(dataArr, 1)
            xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
        Next i
    Next j
    

    ' Compute 期收/付遠匯款-換匯遠期
    fxReceive = 0
    fxPay = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C3:C" & lastRow)

    For Each rng In rngs
        If rng.Value = "155930402" Then
            fxReceive = fxReceive + rng.Offset(0, 2).Value
        ElseIf rng.Value = "255930402" Then
            fxPay = fxPay + rng.Offset(0, 2).Value
        End If
    Next rng


    fxReceive = Round(fxReceive / 1000, 0)
    fxPay = Round(fxPay / 1000, 0)
    
    xlsht.Range("其他金融資產_淨額").Value = fxReceive
    xlsht.Range("其他").Value = fxReceive
    xlsht.Range("CNY1_資產總計").Value = fxReceive

    xlsht.Range("其他金融負債").Value = fxPay
    xlsht.Range("其他什項金融負債").Value = fxPay
    xlsht.Range("CNY1_負債總計").Value = fxPay
    
    'Set Number Format
    xlsht.Cells(2, "T").Resize(2, 1).NumberFormat = "#,##,00"
    
    MsgBox "Finish " & reportTitle & "分頁取得資料庫資料"
    
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融資產_淨額", xlsht.Range("其他金融資產_淨額").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他", xlsht.Range("其他").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_資產總計", xlsht.Range("CNY1_資產總計").Value, "期收遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他金融負債", xlsht.Range("其他金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "其他什項金融負債", xlsht.Range("其他什項金融負債").Value, "期付遠匯款-換匯遠期"
    InsertIntoTable DBsPath, "MonthlyDeclarationReport", dataMonthString, reportTitle, "CNY1_負債總計", xlsht.Range("CNY1_負債總計").Value, "期付遠匯款-換匯遠期"
    MsgBox "Finish " & reportTitle & "申報資料寫入資料庫"


    ' Release Resources
    Set rng = Nothing
    Set rngs = Nothing
    Set xlsht = Nothing
    Set controlSheet = Nothing   
End Sub

以上是我現在excel vba前端的程式碼，有一個自動化報表的專案使用microsoft access為資料庫，前端使用excel抓取後端access建立的Query，
上面是前端的部分完成好的程式碼，抓取後端Query的資料後進行處理，因為我有十幾張報表，我每一張報表都要建立類似這樣的class

' === clsCNY1 Class Module ===
Option Explicit

Private pFieldValues As Object    ' 存放欄位值 (key: 欄位名稱, value: 數值)
Private pFieldPositions As Object ' 存放 Excel 儲存格位置 (key: 欄位名稱, value: 儲存格地址)

Private Sub Class_Initialize()
    Set pFieldValues = CreateObject("Scripting.Dictionary")
    Set pFieldPositions = CreateObject("Scripting.Dictionary")
    
    ' 預先設定欄位名稱與 Excel 儲存格地址
    pFieldPositions.Add "其他金融資產_淨額", "B2"
    pFieldPositions.Add "其他", "C2"
    pFieldPositions.Add "CNY1_資產總計", "D2"
    pFieldPositions.Add "其他金融負債", "E2"
    pFieldPositions.Add "其他什項金融負債", "F2"
    pFieldPositions.Add "CNY1_負債總計", "G2"
    
    ' 初始化所有欄位值為 Null
    Dim key As Variant
    For Each key In pFieldPositions.Keys
        pFieldValues.Add key, Null
    Next key
End Sub

' 設定欄位值
Public Sub SetField(ByVal fieldName As String, ByVal fieldValue As Variant)
    If pFieldValues.Exists(fieldName) Then
        pFieldValues(fieldName) = fieldValue
    Else
        Err.Raise 1001, , "欄位名稱 [" & fieldName & "] 不存在！"
    End If
End Sub

' 取得所有欄位的值
Public Function GetAllFieldValues() As Object
    Set GetAllFieldValues = pFieldValues
End Function

' 取得所有欄位的儲存格位置
Public Function GetAllFieldPositions() As Object
    Set GetAllFieldPositions = pFieldPositions
End Function

' 驗證欄位是否有值（Null 代表未填）
Public Function ValidateFields() As Boolean
    Dim key As Variant
    Dim msg As String
    msg = ""
    
    For Each key In pFieldValues.Keys
        If IsNull(pFieldValues(key)) Then
            msg = msg & key & vbCrLf
        End If
    Next key
    
    If msg <> "" Then
        MsgBox "以下欄位未填入有效資料:" & vbCrLf & msg, vbExclamation
        ValidateFields = False
    Else
        ValidateFields = True
    End If
End Function

' 將 class 內的值填入 Excel 報表
Public Sub ApplyToWorksheet(ByVal ws As Worksheet)
    Dim key As Variant
    For Each key In pFieldValues.Keys
        If Not IsNull(pFieldValues(key)) Then
            ws.Range(pFieldPositions(key)).Value = pFieldValues(key)
        End If
    Next key
End Sub


請針對我下面想要做的事情，仔細針對每一句的每個步驟，幫我建立vba專案程式碼
1.每個class會先定義好要填入的欄位名稱，欄位初始化數值，和實際申報報表的儲存格位置，
2.在entry進入時，我會先宣告一些config設定值和會使用到的class 報表名稱，
3.先假設有 CNY1 MM4901B Ac5601 AC5602 三張報表，因為我有很多張報表，我要將這些報表使用迴圈逐一初始化，
然後逐一將每張表的每個欄位、欄位初始數值、欄位儲存格位置、報表名稱等必要資訊在access資料表中預先建立資料庫的初始資料，
4.接下來會建立分別的Sub(例如我上面使用RunTotal call CNY1)，需要從資料庫中取得query資料，進行資料處理後將欄位資料set value到分別的報表instance中，
然後更新資料庫的資料，並進行相關必要的處理
5.接下來我要將更新完成的class instance中的資料，逐一檢查是否完成資料的更新(這部分請看在每個報表的sub中處理或是到這邊一次處理)
6.然後我要針對原先已經設置好的申報報表路徑，將已經更新好的class instance欄位數據更新到，申報報表路徑中的每一個excel檔案中的相關欄位，然後另存新檔


'--------
'|Answer|
'--------


下面提供一份較完整的 VBA 專案範例，滿足你各項需求。專案分為兩個主要部分：

1. **共用 Class 模組：clsReport**  
   每個報表物件都用同一個 class（你可以依需要繼承或複製），內含下列內容：
   - 報表名稱
   - 欄位名稱、初始值（預設為 Null）與對應 Excel 儲存格位置（直接寫死在 class 內）
   - 方法：設定欄位值、檢查欄位是否填值、將數據寫入 Excel 工作表

2. **主模組 (Module1)：**  
   包含 config 設定、連線 Access 的相關函式（例如 InsertIntoTable、GetAccessDataAsArray、IsValidDataMonth）以及下列主要流程：
   - (1) 初始化所有報表（假設有 "CNY1"、"MM4901B"、"AC5601"、"AC5602"）– 透過迴圈建立各個 clsReport 物件，並利用每個報表內定義的欄位資訊，在 Access 資料表中預先建立初始資料。
   - (2) 依據每張報表專用的 Sub（例如 CNY1_Sub）來從 Access 取得 Query 資料、處理數據、更新各個報表物件的欄位數值，再更新資料庫（用 INSERT 或 UPDATE，這裡示範 INSERT）。
   - (3) 最後依序讀取各份申報報表 Excel 檔案（路徑在 config 中已設定），將更新後的 class 內數據依照對應儲存格位置更新到各檔案中，並另存新檔。

以下為完整範例程式碼：

---

### 【1】 Class 模組：clsReport

將此程式碼放入一個 Class 模組，命名為 **clsReport**。

```vba
Option Explicit

Private pReportName As String            ' 報表名稱
Private pFieldValues As Object           ' Dictionary：key = 欄位名稱，value = 數值
Private pFieldPositions As Object        ' Dictionary：key = 欄位名稱，value = Excel 儲存格位置

' --- 初始化報表 (根據報表名稱建立欄位定義) ---
Public Sub Init(ByVal reportName As String)
    pReportName = reportName
    Set pFieldValues = CreateObject("Scripting.Dictionary")
    Set pFieldPositions = CreateObject("Scripting.Dictionary")
    
    ' 根據報表名稱，加入各個欄位、設定儲存格位置，初始值皆為 Null
    Select Case reportName
        Case "CNY1"
            AddField "其他金融資產_淨額", "B2", Null
            AddField "其他", "C2", Null
            AddField "CNY1_資產總計", "D2", Null
            AddField "其他金融負債", "E2", Null
            AddField "其他什項金融負債", "F2", Null
            AddField "CNY1_負債總計", "G2", Null
        Case "MM4901B"
            ' 請根據 MM4901B 報表實際情形設定欄位與儲存格
            AddField "短期負債", "B3", Null
            AddField "長期負債", "C3", Null
        Case "AC5601"
            AddField "資產總計", "B4", Null
            AddField "負債總計", "C4", Null
        Case "AC5602"
            AddField "總收入", "B5", Null
            AddField "總支出", "C5", Null
        ' 如有其他報表，可再加入
    End Select
End Sub

' --- 私有：加入欄位定義 ---
Private Sub AddField(ByVal fieldName As String, ByVal cellAddress As String, ByVal initValue As Variant)
    pFieldPositions.Add fieldName, cellAddress
    pFieldValues.Add fieldName, initValue
End Sub

' --- 設定某欄位數值 ---
Public Sub SetField(ByVal fieldName As String, ByVal value As Variant)
    If pFieldValues.Exists(fieldName) Then
        pFieldValues(fieldName) = value
    Else
        Err.Raise 1001, , "欄位名稱 [" & fieldName & "] 不存在於報表 " & pReportName
    End If
End Sub

' --- 取得所有欄位數值 (Dictionary) ---
Public Function GetAllFieldValues() As Object
    Set GetAllFieldValues = pFieldValues
End Function

' --- 取得所有欄位的儲存格位置 (Dictionary) ---
Public Function GetAllFieldPositions() As Object
    Set GetAllFieldPositions = pFieldPositions
End Function

' --- 驗證報表中是否每個欄位都有填入數值 (Null 表示未填) ---
Public Function ValidateFields() As Boolean
    Dim key As Variant, msg As String
    msg = ""
    For Each key In pFieldValues.Keys
        If IsNull(pFieldValues(key)) Then
            msg = msg & key & vbCrLf
        End If
    Next key
    
    If msg <> "" Then
        MsgBox "報表 [" & pReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg, vbExclamation
        ValidateFields = False
    Else
        ValidateFields = True
    End If
End Function

' --- 將 class 內的數值根據儲存格位置寫入指定的 Excel 工作表 ---
Public Sub ApplyToWorksheet(ByVal ws As Worksheet)
    Dim key As Variant
    For Each key In pFieldValues.Keys
        If Not IsNull(pFieldValues(key)) Then
            ws.Range(pFieldPositions(key)).Value = pFieldValues(key)
        End If
    Next key
End Sub

' --- 報表名稱屬性 ---
Public Property Get ReportName() As String
    ReportName = pReportName
End Property
```

---

### 【2】 標準模組：Module1

將下列程式碼放入一個標準模組 (例如 Module1) 中。此模組包含連線 Access 的函式、config 設定，以及各主要步驟流程。

```vba
Option Explicit

'=== 連線 Access & 資料庫相關函式 ===

' (1) 插入資料至 Access DB (可用於初始資料建立或後續更新)
Public Sub InsertIntoTable(ByVal DBPath As String, _
                           ByVal tableName As String, _
                           ByVal dataMonthString As String, _
                           ByVal reportTitle As String, _
                           ByVal fieldCode As String, _
                           ByVal content As Double, _
                           ByVal description As String)
    Dim conn As Object, cmd As Object
    Dim sql As String
    
    ' 建立 ADODB 連線
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 建立 SQL 指令
    sql = "INSERT INTO " & tableName & " (DataMonthString, ReportTitle, FieldCode, Content, Description, CaseCreatedAt) " & _
          "VALUES ('" & dataMonthString & "', '" & reportTitle & "', '" & fieldCode & "', " & content & ", '" & description & "', Now());"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
End Sub

' (2) 從 Access 取得 Query 資料，回傳二維陣列
Public Function GetAccessDataAsArray(ByVal DBPath As String, _
                                     ByVal QueryName As String, _
                                     Optional ByVal dataMonthString As String = vbNullString) As Variant
    Dim conn As Object, cmd As Object, rs As Object
    Dim dataArr As Variant
    Dim colCount As Integer, rowCount As Integer
    Dim headerArr() As String, i As Integer, j As Integer
    
    On Error GoTo ErrHandler
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = QueryName
    cmd.CommandType = 4 ' 儲存查詢
    
    If dataMonthString <> vbNullString Then
        cmd.Parameters.Append cmd.CreateParameter("DataMonthParam", 200, 1, 255, dataMonthString)
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open cmd
    
    If rs Is Nothing Or rs.EOF Then
        MsgBox "查詢結果為空，請檢查資料庫與查詢條件。", vbExclamation
        GetAccessDataAsArray = Array()
        Exit Function
    End If
    
    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next i
    
    dataArr = rs.GetRows()
    rowCount = UBound(dataArr, 2) + 1
    
    Dim resultArr() As Variant
    ReDim resultArr(0 To rowCount, 0 To colCount - 1)
    
    ' 第一列存放欄位名稱
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next i
    ' 後續存放資料
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next j
    Next i
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    GetAccessDataAsArray = resultArr
    Exit Function
ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
    GetAccessDataAsArray = Array()
End Function

' (3) 驗證使用者輸入之資料月份格式 (yyyy/mm)
Public Function IsValidDataMonth(ByVal userInput As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "^\d{4}/(0[1-9]|1[0-2])$"
        .IgnoreCase = True
        .Global = False
    End With
    
    IsValidDataMonth = regex.Test(Trim(userInput))
End Function

'=== 全域設定與變數 ===

Public gDataMonthString As String    ' 資料月份，由使用者輸入
Public gDBPath As String               ' 資料庫路徑
Public gReportFolder As String         ' 申報報表 Excel 檔案所在資料夾（原始檔案）
Public gOutputFolder As String         ' 更新後另存新檔的資料夾
Public gReportNames As Variant         ' 報表名稱陣列
Public gReports As Collection          ' 存放所有報表的 clsReport 物件

'=== 主流程入口 ===

Public Sub Main()
    ' (a) 取得使用者輸入之資料月份
    Dim isInputValid As Boolean
    isInputValid = False
    Do
        gDataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(gDataMonthString) Then
            isInputValid = True
        ElseIf Trim(gDataMonthString) = "" Then
            MsgBox "請輸入報表資料所屬的年度/月份 (例如: 2024/01)", vbExclamation, "輸入錯誤"
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
        End If
    Loop Until isInputValid
    
    ' (b) 設定其他 config 參數
    gDBPath = ThisWorkbook.Path & "\" & Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = "C:\申報報表\原始檔\"      ' 請依實際路徑調整
    gOutputFolder = "C:\申報報表\Processed\"      ' 請依實際路徑調整
    gReportNames = Array("CNY1", "MM4901B", "AC5601", "AC5602")
    
    ' (c) 先初始化所有報表，並將各報表的初始資料寫入 Access (預先建立資料庫記錄)
    Call InitializeReports
    
    ' (d) 依各報表取得 Query 資料、處理數據並更新報表物件及資料庫
    Call ProcessReports
    
    ' (e) 最後依據更新後的報表物件，更新各申報 Excel 檔案並另存新檔
    Call UpdateExcelReports
    
    MsgBox "全部流程處理完成！"
End Sub

'=== (c) 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
    Dim rpt As clsReport
    Dim rptName As Variant
    Dim key As Variant
    Dim initDataMonth As String
    Dim i As Integer
    
    ' 假設初始資料月份可用 "Init" 表示，或自行指定
    initDataMonth = "Init"
    
    Set gReports = New Collection
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName
        gReports.Add rpt, rptName
        
        ' 將每個報表的每個欄位初始記錄（欄位名稱、初始值、儲存格位置）寫入 Access 的設定資料表 (例如 ReportConfig)
        For Each key In rpt.GetAllFieldValues().Keys
            ' 這裡假設資料表名稱為 ReportConfig，請根據實際情況調整
            InsertIntoTable gDBPath, "ReportConfig", initDataMonth, rptName, key, 0, _
                "初始設定，儲存格位置: " & rpt.GetAllFieldPositions()(key)
        Next key
    Next rptName
    MsgBox "報表初始化及初始資料建立完成！"
End Sub

'=== (d) 依各報表取得 Query 資料、處理數據並更新報表物件及資料庫 ===
Public Sub ProcessReports()
    Dim rpt As clsReport
    Dim rptName As Variant
    Dim dataArr As Variant
    Dim i As Integer
    Dim key As Variant
    Dim tempValue As Double
    
    ' 範例中，假設各報表都有一組 Query (名稱為 "報表名稱_Query")
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        ' 取得 Access Query 資料 (請根據實際 Query 名稱調整)
        dataArr = GetAccessDataAsArray(gDBPath, rptName & "_Query", gDataMonthString)
        
        ' 示範：假設取得的 Query 資料中，第一筆資料(第一列, 第二欄)為某個欄位的數值，
        ' 實際上你會依據每個報表的運算邏輯處理資料，然後更新報表各欄位
        ' 以下僅以 CNY1 為例：
        If rpt.ReportName = "CNY1" Then
            ' 假設 dataArr(1,0) 為 "其他金融資產_淨額" 的值
            tempValue = CDbl(dataArr(1, 0))
            rpt.SetField "其他金融資產_淨額", tempValue
            ' 其他欄位依此類推……
            ' 例如：
            rpt.SetField "其他", tempValue
            rpt.SetField "CNY1_資產總計", tempValue
            ' 假設 dataArr(1,1) 為 "其他金融負債" 值：
            tempValue = CDbl(dataArr(1, 1))
            rpt.SetField "其他金融負債", tempValue
            rpt.SetField "其他什項金融負債", tempValue
            rpt.SetField "CNY1_負債總計", tempValue
        ElseIf rpt.ReportName = "MM4901B" Then
            ' 請依該報表實際邏輯處理
            rpt.SetField "短期負債", 100    ' 範例數值
            rpt.SetField "長期負債", 200
        ElseIf rpt.ReportName = "AC5601" Then
            rpt.SetField "資產總計", 300
            rpt.SetField "負債總計", 150
        ElseIf rpt.ReportName = "AC5602" Then
            rpt.SetField "總收入", 500
            rpt.SetField "總支出", 400
        End If
        
        ' 驗證該報表資料是否完整，若完整則更新至資料庫
        If rpt.ValidateFields() Then
            For Each key In rpt.GetAllFieldValues().Keys
                ' 這裡以 InsertIntoTable 為例，更新資料庫中報表的各個欄位值，
                ' 請根據實際需求調整（若記錄已存在則改用 UPDATE 語法）
                InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rpt.ReportName, key, _
                    CDbl(rpt.GetAllFieldValues()(key)), "由 Query 處理更新"
            Next key
        Else
            MsgBox "報表 [" & rpt.ReportName & "] 資料不完整，請檢查！", vbCritical
        End If
    Next rptName
    MsgBox "各報表 Query 資料處理並更新資料庫完成！"
End Sub

'=== (e) 更新各申報 Excel 檔案，將報表物件數值寫入指定儲存格，另存新檔 ===
Public Sub UpdateExcelReports()
    Dim rpt As clsReport
    Dim rptName As Variant
    Dim wb As Workbook, ws As Worksheet
    Dim reportFilePath As String, outputFilePath As String
    
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        
        ' 開啟原始報表 Excel 檔（檔名以報表名稱命名）
        reportFilePath = gReportFolder & rptName & ".xlsx"
        On Error Resume Next
        Set wb = Workbooks.Open(reportFilePath)
        If wb Is Nothing Then
            MsgBox "無法開啟檔案: " & reportFilePath, vbExclamation
            On Error GoTo 0
            GoTo NextRpt
        End If
        On Error GoTo 0
        
        ' 假設資料更新至第一個工作表
        Set ws = wb.Sheets(1)
        ' 將報表物件數值寫入工作表指定儲存格
        rpt.ApplyToWorksheet ws
        
        ' 另存新檔到輸出資料夾
        outputFilePath = gOutputFolder & rptName & "_Processed.xlsx"
        wb.SaveAs Filename:=outputFilePath
        wb.Close SaveChanges:=False
        
NextRpt:
    Next rptName
    MsgBox "所有 Excel 申報報表已更新並另存！"
End Sub
```

---

### 使用說明

1. **先建立好 Class 模組 clsReport。**  
   根據各報表實際欄位及儲存格位置，修改 `Select Case` 中各 Case 區段。

2. **在 ControlPanel 工作表中設定資料庫檔名（DBsPathFileName）。**

3. **修改 Module1 中 config 路徑變數：**  
   - `gReportFolder`：原始申報報表 Excel 檔所在的資料夾  
   - `gOutputFolder`：處理後另存新檔的資料夾

4. **執行 Main Sub：**  
   程式會依序要求輸入資料月份、初始化報表（並預先在 Access 中建立初始資料）、從 Access 取得 Query 資料並處理更新、最後讀取 Excel 申報報表檔案更新資料後另存新檔。

這樣的設計可以讓你針對數十張報表，統一利用迴圈進行初始化、資料更新、資料庫寫入與 Excel 更新等自動化處理工作。請根據實際情況調整各部分細節（例如各報表欄位、查詢邏輯與資料庫更新方式）。









'----------
'|Question|
'----------

有關ProcessReports部分，因為每張報表要處理的程序可能很不一樣，要寫在不同的sub中，請針對這個部分再幫我修改一下，請延續上一個回答做修改，並給我完整的版本



'--------
'|Answer|
'--------


以下提供一個完整的範例，此版本延續前面的設計，但將各報表的資料處理邏輯拆分成獨立的 Sub，方便各報表依各自需求進行處理。整個專案包含兩大部分：

1. **共用 Class 模組：clsReport**  
   此 class 內定義了報表名稱、欄位名稱、初始值（預設為 Null）與對應的 Excel 儲存格位置，並提供設定欄位、驗證資料、將數值寫入工作表等方法。

2. **主模組（Module1）**  
   包含連線 Access 的函式、全域 config 設定與主要流程：
   - 取得使用者輸入的資料月份與其他設定參數。
   - 以迴圈依照報表名稱（例如 "CNY1"、"MM4901B"、"AC5601"、"AC5602"）建立各報表的 clsReport 物件，並預先寫入 Access（例如存入初始設定）。
   - 各報表各自的處理邏輯分別寫在獨立的 Sub（例如 ProcessCNY1、ProcessMM4901B……），這些 Sub 會從 Access 取得 Query 資料、處理數據後更新報表物件的欄位，並將結果更新到資料庫。
   - 最後，再依據更新完成的報表物件，依設定的路徑讀取對應的 Excel 申報檔，將數值寫入指定儲存格後另存新檔。

下面的程式碼即為完整版本：

–––––【1】 clsReport Class 模組 –––––
將此程式碼放入一個 Class 模組（名稱請設定為 clsReport）：

```vba
Option Explicit

Private pReportName As String            ' 報表名稱
Private pFieldValues As Object           ' Dictionary：key = 欄位名稱，value = 數值
Private pFieldPositions As Object        ' Dictionary：key = 欄位名稱，value = Excel 儲存格位置

'=== 初始化報表 (根據報表名稱建立欄位定義) ===
Public Sub Init(ByVal reportName As String)
    pReportName = reportName
    Set pFieldValues = CreateObject("Scripting.Dictionary")
    Set pFieldPositions = CreateObject("Scripting.Dictionary")
    
    ' 根據報表名稱，加入各個欄位、設定儲存格位置，初始值皆為 Null
    Select Case reportName
        Case "CNY1"
            AddField "其他金融資產_淨額", "B2", Null
            AddField "其他", "C2", Null
            AddField "CNY1_資產總計", "D2", Null
            AddField "其他金融負債", "E2", Null
            AddField "其他什項金融負債", "F2", Null
            AddField "CNY1_負債總計", "G2", Null
        Case "MM4901B"
            ' 請根據 MM4901B 報表實際欄位設定
            AddField "短期負債", "B3", Null
            AddField "長期負債", "C3", Null
        Case "AC5601"
            AddField "資產總計", "B4", Null
            AddField "負債總計", "C4", Null
        Case "AC5602"
            AddField "總收入", "B5", Null
            AddField "總支出", "C5", Null
        ' 如有其他報表，請加入對應設定
    End Select
End Sub

'--- 私有方法：加入欄位定義 ---
Private Sub AddField(ByVal fieldName As String, ByVal cellAddress As String, ByVal initValue As Variant)
    pFieldPositions.Add fieldName, cellAddress
    pFieldValues.Add fieldName, initValue
End Sub

'=== 設定某欄位數值 ===
Public Sub SetField(ByVal fieldName As String, ByVal value As Variant)
    If pFieldValues.Exists(fieldName) Then
        pFieldValues(fieldName) = value
    Else
        Err.Raise 1001, , "欄位名稱 [" & fieldName & "] 不存在於報表 " & pReportName
    End If
End Sub

'=== 取得所有欄位數值 (Dictionary) ===
Public Function GetAllFieldValues() As Object
    Set GetAllFieldValues = pFieldValues
End Function

'=== 取得所有欄位儲存格位置 (Dictionary) ===
Public Function GetAllFieldPositions() As Object
    Set GetAllFieldPositions = pFieldPositions
End Function

'=== 驗證報表中是否每個欄位都有填入數值 (Null 表示未填) ===
Public Function ValidateFields() As Boolean
    Dim key As Variant, msg As String
    msg = ""
    For Each key In pFieldValues.Keys
        If IsNull(pFieldValues(key)) Then
            msg = msg & key & vbCrLf
        End If
    Next key
    
    If msg <> "" Then
        MsgBox "報表 [" & pReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg, vbExclamation
        ValidateFields = False
    Else
        ValidateFields = True
    End If
End Function

'=== 將 class 內的數值依照儲存格位置寫入指定的 Excel 工作表 ===
Public Sub ApplyToWorksheet(ByVal ws As Worksheet)
    Dim key As Variant
    For Each key In pFieldValues.Keys
        If Not IsNull(pFieldValues(key)) Then
            ws.Range(pFieldPositions(key)).Value = pFieldValues(key)
        End If
    Next key
End Sub

'=== 報表名稱屬性 ===
Public Property Get ReportName() As String
    ReportName = pReportName
End Property
```

–––––【2】 Module1 –––––  
將以下程式碼放入標準模組（例如 Module1），此模組包含連線 Access 的函式、全域設定與主要流程。

```vba
Option Explicit

'=== 連線 Access & 資料庫相關函式 ===

' (1) 插入資料至 Access DB（可用於初始資料建立或後續更新）
Public Sub InsertIntoTable(ByVal DBPath As String, _
                           ByVal tableName As String, _
                           ByVal dataMonthString As String, _
                           ByVal reportTitle As String, _
                           ByVal fieldCode As String, _
                           ByVal content As Double, _
                           ByVal description As String)
    Dim conn As Object, cmd As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    sql = "INSERT INTO " & tableName & " (DataMonthString, ReportTitle, FieldCode, Content, Description, CaseCreatedAt) " & _
          "VALUES ('" & dataMonthString & "', '" & reportTitle & "', '" & fieldCode & "', " & content & ", '" & description & "', Now());"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
End Sub

' (2) 從 Access 取得 Query 資料，回傳二維陣列
Public Function GetAccessDataAsArray(ByVal DBPath As String, _
                                     ByVal QueryName As String, _
                                     Optional ByVal dataMonthString As String = vbNullString) As Variant
    Dim conn As Object, cmd As Object, rs As Object
    Dim dataArr As Variant
    Dim colCount As Integer, rowCount As Integer
    Dim headerArr() As String, i As Integer, j As Integer
    
    On Error GoTo ErrHandler
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = QueryName
    cmd.CommandType = 4 ' 儲存查詢
    
    If dataMonthString <> vbNullString Then
        cmd.Parameters.Append cmd.CreateParameter("DataMonthParam", 200, 1, 255, dataMonthString)
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open cmd
    
    If rs Is Nothing Or rs.EOF Then
        MsgBox "查詢結果為空，請檢查資料庫與查詢條件。", vbExclamation
        GetAccessDataAsArray = Array()
        Exit Function
    End If
    
    colCount = rs.Fields.Count
    ReDim headerArr(0 To colCount - 1)
    For i = 0 To colCount - 1
        headerArr(i) = rs.Fields(i).Name
    Next i
    
    dataArr = rs.GetRows()
    rowCount = UBound(dataArr, 2) + 1
    
    Dim resultArr() As Variant
    ReDim resultArr(0 To rowCount, 0 To colCount - 1)
    
    ' 第一列存放欄位名稱
    For i = 0 To colCount - 1
        resultArr(0, i) = headerArr(i)
    Next i
    ' 後續存放資料
    For i = 0 To colCount - 1
        For j = 1 To rowCount
            resultArr(j, i) = dataArr(i, j - 1)
        Next j
    Next i
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    GetAccessDataAsArray = resultArr
    Exit Function
ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
    GetAccessDataAsArray = Array()
End Function

' (3) 驗證使用者輸入之資料月份格式 (yyyy/mm)
Public Function IsValidDataMonth(ByVal userInput As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "^\d{4}/(0[1-9]|1[0-2])$"
        .IgnoreCase = True
        .Global = False
    End With
    
    IsValidDataMonth = regex.Test(Trim(userInput))
End Function

'=== 全域設定與變數 ===

Public gDataMonthString As String    ' 由使用者輸入的資料月份
Public gDBPath As String               ' 資料庫路徑
Public gReportFolder As String         ' 原始申報報表 Excel 檔所在資料夾
Public gOutputFolder As String         ' 更新後另存新檔的資料夾
Public gReportNames As Variant         ' 報表名稱陣列
Public gReports As Collection          ' 存放所有報表 (clsReport) 的 Collection

'=== 主流程入口 ===

Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False
    Do
        gDataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(gDataMonthString) Then
            isInputValid = True
        ElseIf Trim(gDataMonthString) = "" Then
            MsgBox "請輸入報表資料所屬的年度/月份 (例如: 2024/01)", vbExclamation, "輸入錯誤"
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
        End If
    Loop Until isInputValid
    
    ' 設定其他 config 參數
    gDBPath = ThisWorkbook.Path & "\" & Sheets("ControlPanel").Range("DBsPathFileName").Value
    gReportFolder = "C:\申報報表\原始檔\"      ' 請依實際路徑調整
    gOutputFolder = "C:\申報報表\Processed\"      ' 請依實際路徑調整
    gReportNames = Array("CNY1", "MM4901B", "AC5601", "AC5602")
    
    ' (a) 先初始化所有報表，並將初始資料寫入 Access
    Call InitializeReports
    
    ' (b) 各報表分別進行資料處理 (各自邏輯分離)
    Call ProcessCNY1
    Call ProcessMM4901B
    Call ProcessAC5601
    Call ProcessAC5602
    
    ' (c) 最後更新申報 Excel 檔案並另存新檔
    Call UpdateExcelReports
    
    MsgBox "全部流程處理完成！"
End Sub

'=== (a) 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
    Dim rpt As clsReport
    Dim rptName As Variant
    Dim key As Variant
    Dim initDataMonth As String
    
    initDataMonth = "Init"   ' 初始資料月份，可自行調整
    Set gReports = New Collection
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName
        gReports.Add rpt, rptName
        
        ' 將該報表所有欄位初始設定寫入 Access（資料表名稱例如 ReportConfig）
        For Each key In rpt.GetAllFieldValues().Keys
            InsertIntoTable gDBPath, "ReportConfig", initDataMonth, rptName, key, 0, _
                "初始設定，儲存格位置: " & rpt.GetAllFieldPositions()(key)
        Next key
    Next rptName
    MsgBox "報表初始化及初始資料建立完成！"
End Sub

'=== (b) 各報表獨立處理邏輯 ===

'【CNY1】資料處理
Public Sub ProcessCNY1()
    Dim rpt As clsReport
    Dim dataArr As Variant
    Dim tempValue As Double
    Dim key As Variant
    
    Set rpt = gReports("CNY1")
    ' 取得 Query 資料 (查詢名稱為 "CNY1_Query")
    dataArr = GetAccessDataAsArray(gDBPath, "CNY1_Query", gDataMonthString)
    If UBound(dataArr) < 1 Then
        MsgBox "CNY1 查詢資料不完整！", vbExclamation
        Exit Sub
    End If
    
    ' 假設資料陣列中第一筆資料各欄位依序排列
    tempValue = CDbl(dataArr(1, 0)) ' 例如：其他金融資產_淨額
    rpt.SetField "其他金融資產_淨額", tempValue
    rpt.SetField "其他", tempValue
    rpt.SetField "CNY1_資產總計", tempValue
    
    tempValue = CDbl(dataArr(1, 1)) ' 例如：其他金融負債
    rpt.SetField "其他金融負債", tempValue
    rpt.SetField "其他什項金融負債", tempValue
    rpt.SetField "CNY1_負債總計", tempValue
    
    ' 驗證並更新資料庫（以 Insert 作示範，若記錄已存在則應使用 UPDATE）
    If rpt.ValidateFields() Then
        For Each key In rpt.GetAllFieldValues().Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rpt.ReportName, key, _
                CDbl(rpt.GetAllFieldValues()(key)), "由 CNY1 Query 處理更新"
        Next key
    End If
End Sub

'【MM4901B】資料處理
Public Sub ProcessMM4901B()
    Dim rpt As clsReport
    Set rpt = gReports("MM4901B")
    ' 取得 Query 資料 (查詢名稱 "MM4901B_Query")
    Dim dataArr As Variant
    dataArr = GetAccessDataAsArray(gDBPath, "MM4901B_Query", gDataMonthString)
    
    ' 依據該報表邏輯進行資料處理，此處以示範數值代入：
    rpt.SetField "短期負債", 100
    rpt.SetField "長期負債", 200
    
    Dim key As Variant
    If rpt.ValidateFields() Then
        For Each key In rpt.GetAllFieldValues().Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rpt.ReportName, key, _
                CDbl(rpt.GetAllFieldValues()(key)), "由 MM4901B Query 處理更新"
        Next key
    End If
End Sub

'【AC5601】資料處理
Public Sub ProcessAC5601()
    Dim rpt As clsReport
    Set rpt = gReports("AC5601")
    Dim dataArr As Variant
    dataArr = GetAccessDataAsArray(gDBPath, "AC5601_Query", gDataMonthString)
    
    ' 依據實際運算邏輯處理，示範：
    rpt.SetField "資產總計", 300
    rpt.SetField "負債總計", 150
    
    Dim key As Variant
    If rpt.ValidateFields() Then
        For Each key In rpt.GetAllFieldValues().Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rpt.ReportName, key, _
                CDbl(rpt.GetAllFieldValues()(key)), "由 AC5601 Query 處理更新"
        Next key
    End If
End Sub

'【AC5602】資料處理
Public Sub ProcessAC5602()
    Dim rpt As clsReport
    Set rpt = gReports("AC5602")
    Dim dataArr As Variant
    dataArr = GetAccessDataAsArray(gDBPath, "AC5602_Query", gDataMonthString)
    
    ' 依據實際運算邏輯處理，示範：
    rpt.SetField "總收入", 500
    rpt.SetField "總支出", 400
    
    Dim key As Variant
    If rpt.ValidateFields() Then
        For Each key In rpt.GetAllFieldValues().Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rpt.ReportName, key, _
                CDbl(rpt.GetAllFieldValues()(key)), "由 AC5602 Query 處理更新"
        Next key
    End If
End Sub

'=== (c) 更新申報 Excel 檔案，將各報表物件數值寫入對應儲存格，並另存新檔 ===
Public Sub UpdateExcelReports()
    Dim rpt As clsReport
    Dim rptName As Variant
    Dim wb As Workbook, ws As Worksheet
    Dim reportFilePath As String, outputFilePath As String
    
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        
        ' 開啟原始 Excel 檔（檔名以報表名稱命名）
        reportFilePath = gReportFolder & rptName & ".xlsx"
        On Error Resume Next
        Set wb = Workbooks.Open(reportFilePath)
        If wb Is Nothing Then
            MsgBox "無法開啟檔案: " & reportFilePath, vbExclamation
            On Error GoTo 0
            GoTo NextRpt
        End If
        On Error GoTo 0
        
        ' 假設數據更新至第一個工作表
        Set ws = wb.Sheets(1)
        rpt.ApplyToWorksheet ws
        
        outputFilePath = gOutputFolder & rptName & "_Processed.xlsx"
        wb.SaveAs Filename:=outputFilePath
        wb.Close SaveChanges:=False
NextRpt:
    Next rptName
    MsgBox "所有 Excel 申報報表已更新並另存！"
End Sub
```

–––––【使用說明】–––––
1. 請先建立好 Class 模組（clsReport），根據各報表實際欄位與儲存格位置調整 Select Case 區段。  
2. 在 ControlPanel 工作表中設定資料庫檔名（例如 DBsPathFileName）。  
3. 根據環境修改 Module1 中的路徑參數（gReportFolder 與 gOutputFolder）。  
4. 每個報表的資料處理邏輯已分別寫在 ProcessCNY1、ProcessMM4901B、ProcessAC5601、ProcessAC5602 中，如有需要請依實際業務邏輯進行修改。  
5. 執行 Main Sub，系統依序會：  
   - 請使用者輸入資料月份  
   - 初始化各報表並將初始設定寫入 Access（例如存入 ReportConfig 資料表）  
   - 分別呼叫各報表獨立處理的 Sub，從 Access 取得 Query 資料、計算後更新各報表物件，並更新資料庫  
   - 最後根據更新後的報表物件，讀取各 Excel 申報檔，將資料寫入對應儲存格後另存新檔  

這個完整範例可讓你針對多張報表，統一利用迴圈進行初始化、分別處理各報表邏輯、更新資料庫以及更新 Excel 申報檔。請依照實際需求進行進一步調整與擴充。
