Option Explicit

'建立Access資料表
Private Sub btnCreateTable_Click()
    Dim accessApp As Object
    Dim wb As Workbook
    Dim ws As Worksheet

    Dim dbFilePath As String
    Dim lastRow As Integer
    Dim i As Integer

    Dim tableName As String
    Dim sqlString As String
    Dim columnString As String
    dim referenceTable As String
    dim constraintStr As String
    
    '設定檔案路徑
    Do
        dbFilePath = InputBox("請輸入DB資料庫檔案路徑: ", "輸入檔案路徑")
        If dbFilePath = "" Then
            MsgBox "未輸入路徑，退出程序"
            Exit sub
        End If 
        If Dir(dbFilePath) = "" Then
            MsgBox "檔案不存在，請重新輸入有效路徑" 
        End If
    loop while Dir(dbFilePath) = ""

    'Start Access AP
    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase dbFilePath

    'Fetch create tables setting
    Set ws = ThisWorkbook.Worksheets("DataTables")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    tableName = ""
    columnString = ""
    referenceTable = ""
    constraintStr = ""
    
    ' 讀取 Excel 資料
    For i = 2 To lastRow
        'handle firs table
        If tableName = "" Then
         tableName = ws.Cells(i, "B").Value
        End If
        
        'handle last row for each: create table
        If (ws.Cells(i + 1, "B").Value <> tableName) And columnString <> "" Then
            
            columnString = columnString & ", " & "[" & ws.Cells(i, "C").Value & "] " & ws.Cells(i, "D").Value

            if referenceTable <> "" then
                sqlString = "CREATE TABLE " & tableName & " (" & columnString & referenceTable & ")"
            else
                sqlString = "CREATE TABLE " & tableName & " (" & columnString & ")"
            end if

            Debug.Print "sqlString: " & sqlString
            accessApp.CurrentDb.Execute sqlString            
            
            tableName = ws.Cells(i + 1, "B").Value
            columnString = ""
            referenceTable = ""
        Else
            If columnString = "" Then
                If ws.Cells(i, "G").Value = "Autoincrement" Then
                    'columnString = "" & ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value
                    columnString = "[" & ws.Cells(i, "C").Value & "] AUTOINCREMENT PRIMARY KEY "
                else
                    columnString = "[" & ws.Cells(i, "C").Value & "] " & ws.cells(i, "D") & " PRIMARY KEY "
                end if
            Else
                'columnString = columnString & ", " & "" & ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value
                columnString = columnString & ", " & "[" & ws.Cells(i, 3).Value & "] " & ws.Cells(i, 4).Value
            End If
            
            if not isEmpty(ws.cells(i, "E").value) then
                constraintStr = ws.cells(i, "E").value
                constraintStr = Replace(constraintStr, "([", "_")
                constraintStr = Replace(constraintStr, "])", "") & tableName
                referenceTable = ", CONSTRAINT FK_" & constraintStr & " FOREIGN KEY ([" & ws.cells(i, "C").value & "]) REFERENCES " & ws.cells(i, "E").value
            end if
        End If
    Next i

    ' 關閉 Access
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing
    MsgBox "資料表建立完成"
End Sub
