Option Compare Database

Public Sub Initialize(Optional ByVal sheetName As Variant = 1, _
                      Optional ByVal loopColumn As Integer = 1, _
                      Optional ByVal leftToDelete As Integer = 2, _
                      Optional ByVal rightToDelete As Integer = 3, _
                      Optional ByVal rowsToDelete As Variant, _
                      Optional ByVal colsToDelete As Variant)
    'implement operations here
End Sub

Public Sub CleanReport(ByVal fullFilePath As String, _
                       ByVal cleaningType As String)
    'implement operations here
End Sub

Public Sub additionalClean(ByVal fullFilePath As String, _
                           ByVal cleaningType As String, _
                           ByVal dataDate As Date, _
                           ByVal dataMonth As Date, _
                           ByVal dataMonthString As String)
    'implement operations here
End Sub
