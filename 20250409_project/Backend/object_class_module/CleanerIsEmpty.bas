'Fit for 報表OBU_DL6320、OBU_CF6320、DBU_CF6850
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
    Dim i As Integer
    
    Dim isNotDataFoundExist As Boolean
    Dim isNotFoundExistFC7700 As Boolean
    Dim isSubToalExist As Boolean
    Dim isTotalExist As Boolean

    Dim isSubToalEmpty As Boolean
    Dim isTotalEmpty As Boolean

    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
    End If
    
    Set xlApp = Excel.Application
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(1)
    
    isNotDataFoundExist = False
    isNotFoundExistFC7700 = False
    isSubToalExist = False
    isTotalExist = False

    isSubToalEmpty = False
    isTotalEmpty = False

    For i = xlsht.Cells(xlsht.Rows.Count, "A").End(xlUp).row To 2 Step -1
        If Left(xlsht.Cells(i, "A").Value, 2) = "小計" Then
            isSubToalExist = True
            If IsEmpty(xlsht.Cells(i, "B").Value) Then
                isSubToalEmpty = True
            Else
                isSubToalEmpty = False
            End If
        ElseIf Left(xlsht.Cells(i, "A").Value, 2) = "總計" Then
            isTotalExist = True
            If IsEmpty(xlsht.Cells(i, "B").Value) Then
                isTotalEmpty = True
            Else
                isTotalEmpty = False
            End If
        ElseIf xlsht.Cells(i, "A").Value = "No Redord Found!" Then
            isNotDataFoundExist = True
        ElseIf xlsht.Cells(i, "A").Value = "無 資 料" Then
            isNotFoundExistFC7700 = True
        End If
    Next i

    If isNotDataFoundExist And isTotalExist Then
        If isTotalEmpty Then
            Debug.print "報表OBU_DL6320無資料"
        Else
            MsgBox "注意!報表OBU_DL6320格式異常"
            Exit Sub
        End If
    ElseIf isSubToalExist And isTotalExist Then
        If isSubToalEmpty And isTotalEmpty Then
            Debug.print "報表OBU_CF6320或DBU_CF6850無資料"
        Else
            MsgBox "注意!報表OBU_CF6320或DBU_CF6850格式異常"
            Exit Sub
        End If
    ElseIf isNotFoundExistFC7700 Then
        Debug.print "報表OBU_FC7700B無資料"
    ElseIf IsEmpty(xlsht.Cells(2, "A").Value) And _
           IsEmpty(xlsht.Cells(3, "A").Value) And _
           IsEmpty(xlsht.Cells(4, "A").Value) Then
        Debug.print "報表OBU_FC9450B無資料"
    Else
        MsgBox "注意!報表格式異常"
        Exit Sub
    End If
    
    xlApp.DisplayAlerts = False
    xlbk.Close
    xlApp.DisplayAlerts = True

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    
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
