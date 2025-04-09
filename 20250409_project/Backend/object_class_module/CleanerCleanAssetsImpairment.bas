Option Compare Database

Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet    

    ' Dim xlbk As Workbook
    ' Dim xlsht As Worksheet

    Dim lastRow As Long
    Dim i As Integer
    Dim colArray As Variant
    Dim valueType As Variant
    Dim eachType As Variant
    'collection save different type of row
    Dim securityIndex As Collection
    Dim securityName As Collection

    Dim startRow As Integer
    Dim endRow As Integer
    Dim innerLastRow As Integer
    Dim sheetName As String

    Dim toDelete As Boolean

    'fullFilePath = "D:\DavidHsu\testFile\vba\test\金融資產減損 1140123.xlsx"
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
    End If
    
    Set xlApp = Excel.Application
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    ' Set xlbk = Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets("減損")
        
    colArray = Array("Security_id", _
                     "issuer", _
                     "成本", _
                     "應收利息", _
                     "信評", _
                     "PD", _
                     "LGD", _
                     "上期減損數(成本)", _
                     "本期減損數(成本)", _
                     "上期減損數(利息)", _
                     "本期減損數(利息)")
                     
    '分頁名稱不能夠有?，會報錯誤
    valueType = Array("強制FVPL金融資產-公債-中央政府", "強制FVPL金融資產-公債-地方政府(我國)", _
                        "強制FVPL金融資產-普通公司債(公", "強制FVPL金融資產-普通公司債(民", _
                        "強制FVPL金融資產-商業本票", "FVOCI債務工具-央行NCD", _
                        "FVOCI債務工具-公債-中央政府(我", _
                        "FVOCI債務工具-公債-地方政府(我國)", _
                        "FVOCI債務工具-普通公司債（公營", _
                        "FVOCI債務工具-普通公司債（民營", _
                        "AC債務工具-央行NCD", _
                        "AC債務工具投資-公債-中央政府(?", _
                        "AC債務工具投資-公債-地方政府(?", _
                        "AC債務工具投資-普通公司債(公營", _
                        "AC債務工具投資-普通公司債(民營", _
                        "強制FVPL金融資產-公債-中央政府(外國)", _
                        "強制FVPL金融資產-普通公司債(公營)-海外", _
                        "強制FVPL金融資產-普通公司債(民營)-海外", _
                        "FVOCI債務工具-公債-中央政府(外國)", _
                        "FVOCI債務工具-普通公司債(公營)-海外", _
                        "FVOCI債務工具-普通公司債(民營)-海外", _
                        "FVOCI債務工具-金融債券-海外", _
                        "AC債務工具投資-公債-中央政府(外國)", _
                        "AC債務工具投資-普通公司債(公營)-海外", _
                        "AC債務工具投資-普通公司債(民營)-海外", _
                        "AC債務工具投資-金融債券-海外")
                        
    lastRow = xlsht.Cells(xlsht.Rows.Count, 3).End(xlUp).Row
    
    '開始建立分頁和處理資料
    Set securityName = New Collection
    Set securityIndex = New Collection
    tableColumns = GetTableColumns(cleaningType)

    For i = lastRow To 1 Step -1
        If IsEmpty(xlsht.Cells(i, 3).Value) Then
            xlsht.Rows(i).Delete
        End If
        If Left(xlsht.Cells(i, "I").Value, 5) = "利息備抵數" Then
            xlsht.Rows(i & ":" & lastRow).Delete
        End If
    Next i
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, 3).End(xlUp).Row
    
    For i = 1 To lastRow
        For Each eachType In valueType
            If Trim(xlsht.Cells(i, 3).Value) = eachType Then
                securityIndex.Add i
                securityName.Add eachType
            End If
        Next eachType
    Next i
    
    For i = 1 To securityIndex.Count
        If i + 1 <= securityIndex.Count Then
            If securityIndex(i) + 1 = securityIndex(i + 1) Then
                GoTo ContinueLoop
            Else
                startRow = securityIndex(i) + 1
                endRow = securityIndex(i + 1) - 1
            End If
        Else
            startRow = securityIndex(i) + 1
            endRow = lastRow
        End If

        If InStr(securityName(i), "?") > 0 Then
            sheetName = Replace(securityName(i), "?", "")
        Else
            sheetName = securityName(i)
        End If

        xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)

        With ActiveSheet
            .Name = sheetName
            xlsht.Range(xlsht.Cells(startRow, "C"), xlsht.Cells(endRow, "M")).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
            innerLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            .Range("L2:L" & innerLastRow).Value = securityName(i)
            'Set Table Columns
            ' .Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
        End With
ContinueLoop:
    Next i
    
    For i = xlbk.Sheets.Count To 1 Step -1
        'Default cancel
        toDelete = True
        'Check wehther in valueType or not
        For Each eachType In valueType
            If xlbk.Sheets(i).Name = eachType Then
                toDelete = False
                Exit For
            End If
        Next eachType
        'Delete if not in list
        If toDelete Then
            xlbk.Sheets(i).Delete
        End If
    Next i

    xlApp.CutCopyMode = False
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    xlApp.AskToUpdateLinks = True
    
    Set securityIndex = Nothing
    Set securityName = Nothing

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing


    ' On Error Resume Next
    ' Set xlsht = Nothing
    ' Set xlbk = Nothing
    ' Set securityIndex = Nothing
    ' Set securityName = Nothing
    ' On Error GoTo 0
    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    'implement operations here
End Sub
