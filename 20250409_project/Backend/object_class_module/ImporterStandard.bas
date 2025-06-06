Option Compare Database

Implements IImporter

Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
                                ByVal accessDBPath As String, _
                                ByVal tableName As String)
    Dim cn As Object
    Dim xlApp As Object
    Dim xlbk As Object
    Dim sqlString As String
    Dim sheetName As String
    Dim i As Integer, j As Integer

    Dim tableColumns As Variant

    Dim fieldList As String
    Dim selectList As String

    ' 使用 ADODB 連接 Access 資料庫
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

    ' 使用 Excel 來開啟檔案，取得所有分頁名稱
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Excel 不顯示
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)

    ' 取得 Access 資料表的欄位名稱
    tableColumns = GetTableColumns(tableName) ' 假設這個函式回傳欄位名稱陣列

    ' 確保 tableColumns 至少有 2 個欄位（避免 Primary Key 之外沒有欄位）
    If UBound(tableColumns) < 1 Then
    MsgBox "資料表 " & tableName & " 至少需要 2 個欄位（Primary Key + 其他欄位）。", vbCritical
    Exit Sub
    End If

    ' 動態構建 `INSERT INTO` 及 `SELECT` 語法（忽略 Primary Key）
    fieldList = ""
    selectList = ""

    For i = 1 To UBound(tableColumns) ' 從 tableColumns(1) 開始，略過 Primary Key
        fieldList = fieldList & "[" & tableColumns(i) & "],"
        selectList = selectList & "[" & tableColumns(i) & "],"
    Next i

    ' clear last comma
    fieldList = Left(fieldList, Len(fieldList) - 1)
    selectList = Left(selectList, Len(selectList) - 1)

    For i = 1 To xlbk.Sheets.count
        sheetName = xlbk.Sheets(i).Name
        ' Dynamic structure SQL Query language，skip Primary Key
        sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
        "SELECT " & selectList & " FROM [Excel 12.0 Xml;HDR=YES;Database=" & fullFilePath & "].[" & sheetName & "$]"
        ' 執行 SQL
        cn.Execute sqlString
    Next i

    ' 關閉 Excel 檔案和釋放物件
    xlbk.Close False
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    ' 關閉 ADODB 連接
    cn.Close
    Set cn = Nothing

    MsgBox "完成 " & tableName & " 資料表匯入作業"
    Debug.Print "完成 " & tableName & " 資料表匯入作業"

End Sub


'old_version
' Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
'                                 ByVal accessDBPath As String, _
'                                 ByVal tableName As String)
'     Dim cn As Object
'     Dim xlApp As Object
'     Dim xlbk As Object
'     Dim sqlString As String
'     Dim sheetName As String
'     Dim i As Integer
'     Dim tableColumns As Variant

'     ' 使用 ADODB 連接 Access 資料庫
'     Set cn = CreateObject("ADODB.Connection")
'     cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

'     ' 使用 Excel 來開啟檔案，取得所有分頁名稱
'     Set xlApp = CreateObject("Excel.Application")
'     xlApp.Visible = False ' Excel 不顯示
'     Set xlbk = xlApp.Workbooks.Open(fullFilePath)

'     tableColumns = GetTableColumns(tableName)
    
'     ' 遍歷所有分頁並匯入每一個分頁
'     For i = 1 To xlbk.Sheets.Count
'         sheetName = xlbk.Sheets(i).Name
'         ' 動態建構 SQL 查詢語句，針對每一個分頁匯入資料
'         sqlString = "INSERT INTO " & tableName & " SELECT * FROM [Excel 12.0 Xml;HDR=YES;Database=" & fullFilePath & "].[" & sheetName & "$]"
'         cn.Execute sqlString
'     Next i

'     ' 關閉 Excel 檔案和釋放物件
'     xlbk.Close False
'     Set xlbk = Nothing
'     Set xlApp = Nothing

'     ' 關閉 ADODB 連接
'     cn.Close
'     Set cn = Nothing
' End Sub
