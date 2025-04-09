Option Compare Database

Public sheetName As Variant
Public loopColumn As Integer
Public leftToDelete As Integer
Public rightToDelete As Integer
Public rowsToDelete As Variant
Public colsToDelete As Variant

' 初始化方法
Public Sub Init(Optional sheetName As Variant = 1, _
                Optional loopColumn As Integer = 1, _
                Optional leftToDelete As Integer = 2, _
                Optional rightToDelete As Integer = 3, _
                Optional rowsToDelete As Variant = Null, _
                Optional colsToDelete As Variant = Null)
    
    ' 設定屬性
    sheetName = sheetName
    loopColumn = loopColumn
    leftToDelete = leftToDelete
    rightToDelete = rightToDelete
    rowsToDelete = IIf(IsEmpty(rowsToDelete), Array(), rowsToDelete)
    colsToDelete = IIf(IsEmpty(colsToDelete), Array(), colsToDelete)
End Sub
