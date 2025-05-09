Option Compare Database

Implements ICleaner

Private clsColsToHandle As Variant

Private clsHasFile As Boolean
Private clsHasData As Boolean

Public Function ICleaner_HasFile() As Boolean
    ICleaner_HasFile = clsHasFile
End Function

Public Function ICleaner_HasData() As Boolean
    ICleaner_HasData = clsHasData
End Function

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant, _
                               Optional ByVal colsToHandle As Variant)        
    clsColsToHandle = colsToHandle
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)
    Dim fso As Object
    Dim inputFile As Object
    Dim outputFile As Object
    Dim inputLines As Variant
    Dim outputLines() As String
    Dim colsDict As Object
    Dim i As Long
    Dim lineParts As Variant
    Dim colIndex As Variant    

    ' 檔案檢查
    If Dir(fullFilePath) = "" Then
        clsHasFile = False
        MsgBox "File not found: " & fullFilePath
        Exit Sub
    Else
        clsHasFile = True
    End If

    ' 讀入所有行
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inputFile = fso.OpenTextFile(fullFilePath, 1, False)
    inputLines = Split(inputFile.ReadAll, vbLf)
    inputFile.Close

    If clsHasData = False Then
        If UBound(inputLines) > 0 Then clsHasData = True
    End If

    ' 如果沒有要清理的欄位，跳出
    If Not IsArray(clsColsToHandle) Or _
       IsEmpty(clsColsToHandle) Or _
       IsNull(clsColsToHandle) Or _
       (Not IsError(LBound(clsColsToHandle)) And LBound(clsColsToHandle) > UBound(clsColsToHandle)) Then
        Exit Sub
    End If

    ' 把 clsColsToHandle 轉成 Dictionary：key = cleaningType, value = array of colIndex
    Set colsDict = CreateObject("Scripting.Dictionary")
    Dim key As String, val As Variant, existingArray As Variant
    For i = 0 To UBound(clsColsToHandle) Step 2
        key = clsColsToHandle(i)
        val = clsColsToHandle(i + 1)
        If colsDict.Exists(key) Then
            existingArray = colsDict(key)
            ReDim Preserve existingArray(UBound(existingArray) + 1)
            existingArray(UBound(existingArray)) = val
            colsDict(key) = existingArray
        Else
            colsDict.Add key, Array(val)
        End If
    Next i

    ReDim outputLines(UBound(inputLines))

    ' 逐行處理（保留 header 原樣）
    For i = 0 To UBound(inputLines)
        If Trim(inputLines(i)) = "" Then GoTo SkipLine
        lineParts = Split(inputLines(i), ",")
        ' 跳過第一列 header
        If i > 0 Then   
            ' RemovePercentCols
            If colsDict.Exists("RemovePercentCols") Then
                For Each colIndex In colsDict("RemovePercentCols")
                    If colIndex - 1 <= UBound(lineParts) Then
                        lineParts(colIndex) = Replace(lineParts(colIndex), "%", "")
                    End If
                Next
            End If
            ' RemoveSpaceCols
            If colsDict.Exists("RemoveSpaceCols") Then
                For Each colIndex In colsDict("RemoveSpaceCols")
                    If colIndex - 1 <= UBound(lineParts) Then
                        lineParts(colIndex) = Replace(lineParts(colIndex), " ", "")
                    End If
                Next
            End If
            ' FormatDateCols （移除 .0 並轉 yyyy-mm-dd）
            If colsDict.Exists("FormatDateCols") Then
                For Each colIndex In colsDict("FormatDateCols")
                    If colIndex - 1 <= UBound(lineParts) Then
                        lineParts(colIndex) = FormatDateTimeFix(lineParts(colIndex))
                    End If
                Next
            End If
        End If

        outputLines(i) = Join(lineParts, ",")
SkipLine:
    Next i

    ' 覆寫回原檔
    Set outputFile = fso.OpenTextFile(fullFilePath, 2, True)
    For i = 0 To UBound(outputLines)
        outputFile.WriteLine outputLines(i)
    Next i
    outputFile.Close

    If clsHasFile And clsHasData Then
        MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath
    End If

    ' ===========================================
    ' ===========================================

'     Dim xlbk As Excel.Workbook
'     Dim xlsht As Excel.Worksheet
'     Dim i As Long, lastRow As Long
'     Dim colsDict As Object
'     Dim colIndex As Variant

'     If Dir(fullFilePath) = "" Then
'         clsHasFile = False
'         MsgBox "File not found: " & fullFilePath
'         Exit Sub
'     Else
'         clsHasFile = True
'     End If

'     Set xlbk = xlApp.Workbooks.Open(fullFilePath)
'     Set xlsht = xlbk.Sheets(1)

'     lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

'     If clsHasData = False Then
'         If (lastRow > 1) Then clsHasData = True
'     End If

'     If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then
'         GoTo CleanUp
'     End If

'     ' 因為Dict在Interface中使用底層有較嚴格的判斷類型，所以建議不使用
'     Set colsDict = CreateObject("Scripting.Dictionary")
'     Dim key As String我我
'     Dim val As Variant
'     Dim existingArray As Variant

'     For i = 0 To UBound(clsColsToHandle) Step 2
'         key = clsColsToHandle(i)
'         val = clsColsToHandle(i + 1)

'         If colsDict.Exists(key) Then
'             existingArray = colsDict(key)
'             ReDim Preserve existingArray(UBound(existingArray) + 1)
'             existingArray(UBound(existingArray)) = val
'             colsDict(key) = existingArray
'         Else
'             colsDict.Add key, Array(val)
'         End If
'     Next i

'     For i = lastRow To 2 Step -1
'         If colsDict.Exists("FormatCols") Then
'             For Each colIndex In colsDict("FormatCols")
'                 With xlsht.Cells(i, colIndex)
'                     .NumberFormat = "@"
'                     .Value = Format(.Value, "00000000")
'                 End With
'                 ' xlsht.Cells(i, colIndex).Value = Format(xlsht.Cells(i, colIndex).Value, "00000000")
'             Next
'         End If

'         If colsDict.Exists("RemovePercentCols") Then
'             For Each colIndex In colsDict("RemovePercentCols")
'                 xlsht.Cells(i, colIndex).Value = Replace(xlsht.Cells(i, colIndex).Value, "%", "")
'             Next
'         End If

'         If colsDict.Exists("RemoveSpaceCols") Then
'             For Each colIndex In colsDict("RemoveSpaceCols")
'                 xlsht.Cells(i, colIndex).Value = Replace(xlsht.Cells(i, colIndex).Value, " ", "")
'             Next
'         End If
'     Next

'     xlbk.Save
' CleanUp:
'     xlbk.Close False
'     Set xlsht = Nothing
'     Set xlbk = Nothing

'     If clsHasFile And clsHasData Then
'         MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath
'     End If
End Sub

' 空的 additionalClean（可留）
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' No additional actions
End Sub
