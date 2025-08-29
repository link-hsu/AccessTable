Option Compare Database

Implements IImporter

Public Sub IImporter_ImportData(ByVal fullFilePath As String, _
                                ByVal accessDBPath As String, _
                                ByVal tableName As String, _
                                ByVal xlApp As Excel.Application)
    Dim cn As Object
    Dim xlbk As Object
    Dim sqlString As String
    Dim i As Integer

    Dim tableColumns As Variant

    Dim fieldList As String
    Dim selectList As String

    Dim fso As Object
    Dim ext As String

    ' 建立 FSO 並取得副檔名
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(fullFilePath))

    If ext = "txt" Then
        fullFilePath = Left(fullFilePath, Len(fullFilePath) - Len(ext)) & "csv"
        ext = "csv"
    End If

    ' 使用 ADODB 連接 Access 資料庫
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDBPath & ";"

    ' 取得 Access 資料表的欄位名稱
    tableColumns = GetTableColumns(tableName) ' 假設這個函式回傳欄位名稱陣列

    ' 確保 tableColumns 至少有 2 個欄位（避免 Primary Key 之外沒有欄位）
    If UBound(tableColumns) < 1 Then
        MsgBox "資料表 " & tableName & " 至少需要 2 個欄位（Primary Key + 其他欄位）。", vbCritical
        WriteLog "資料表 " & tableName & " 至少需要 2 個欄位（Primary Key + 其他欄位）。"
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

    If ext = "csv" Then
        sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
                    "SELECT " & selectList & " FROM " & _
                    "[Text;FMT=Delimited;HDR=YES;Database=" & Left(fullFilePath, InStrRev(fullFilePath, "\") - 1) & "].[" & Mid(fullFilePath, InStrRev(fullFilePath, "\") + 1) & "]"
        cn.Execute sqlString
    Else
        Set xlbk = xlApp.Workbooks.Open(fullFilePath)
        For i = 1 To xlbk.Sheets.Count
            sqlString = "INSERT INTO " & tableName & " (" & fieldList & ") " & _
            "SELECT " & selectList & " FROM [Excel 12.0 Xml;HDR=YES;Database=" & fullFilePath & "].[" & xlbk.Sheets(i).Name & "$]"
            cn.Execute sqlString
        Next i
        xlbk.Close False
        Set xlbk = Nothing
    End If

    ' 關閉 ADODB 連接
    cn.Close
    Set cn = Nothing

    MsgBox "完成 " & tableName & " 資料表匯入作業"
    WriteLog "完成 " & tableName & " 資料表匯入作業"
    ' Debug.Print "完成 " & tableName & " 資料表匯入作業"

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