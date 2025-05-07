Option Compare Database

Public Function GetFilePaths() As Object
    'Define Database and Recordset
    Dim dbs As DAO.Database
    Dim rsConfig As DAO.Recordset
    Dim rsFilePaths As DAO.Recordset

    'Save in Dictionary
    Dim filePathDict As Object

    'Filepaths content
    Dim reportDataDate As Date
    Dim copyFilePath As String

    Dim pathID As String
    Dim pathName As String
    Dim extensionFormat As String
    Dim fullFilePath As String
    
    'Initialize Dictionary
    Set filePathDict = CreateObject("Scripting.Dictionary")
    Set dbs = CurrentDb()

    'Fetch FolderDir and ReportDataDate
    Set rsConfig = dbs.OpenRecordset("SELECT TOP 1 CopyFilePath, ReportDataDate FROM Configuration ORDER BY DataID DESC", dbOpenSnapshot)
    
    If Not rsConfig.EOF Then
        copyFilePath = rsConfig!CopyFilePath
        reportDataDate = Format(rsConfig!ReportDataDate, "yyyymmdd")
    Else
        Set GetFilePaths = filePathDict
        MsgBox "Error: 無法取得Configuration資料表資料"
        WriteLog "Error: 無法取得Configuration資料表資料"
        rsConfig.Close
        Set rsConfig = Nothing
        Exit Function
    End If

    rsConfig.Close
    Set rsConfig = Nothing
    
    'Fetch FilePaths
    Set rsFilePaths = dbs.OpenRecordset("SELECT PathID, PathName, ExtensionFormat FROM FilePaths", dbOpenSnapshot)

    Do While Not rsFilePaths.EOF
        pathID = NZ(rsFilePaths!PathID, "")
        pathName = NZ(rsFilePaths!PathName, "")
        extensionFormat = NZ(rsFilePaths!ExtensionFormat, "txt")

        fullFilePath = copyFilePath & "\" & pathName & reportDataDate & "." & extensionFormat
    
        'Add to dictionary:
        'PathID as Key，fullFilePath as Value
        filePathDict.Add pathID, fullFilePath
        rsFilePaths.MoveNext
    Loop
    
    rsFilePaths.Close
    Set rsFilePaths = Nothing
    Set db = Nothing
    
    Set GetFilePaths = filePathDict
End Function

'---------------------------
'Save Object with Collection
'---------------------------
' Public Function GetConfigsReturnCollection(ByVal Optional reportType As Variant, _
'                            ByVal Optional reportTypeDate AS Date) As Object
'     'Define Database and Recordset
'     Dim dbs As DAO.Database
'     Dim rsConfig As DAO.Recordset
'     Dim rsFilePaths As DAO.Recordset

'     'Save in Collection and Dictionary
'     Dim configCollection As Collection
'     Dim filePathDict As Object

'     'Table Configuration
'     Dim reportDataDate As Date
'     Dim reportMonth As Date
'     Dim reportMonthString As String
'     Dim copyFilePath As String

'     'Table FilePaths
'     Dim pathID As String
'     Dim pathName As String
'     Dim extensionFormat As String
'     Dim fullFilePath As String
    
'     'Check Optional Parameters
'     Dim checkOpt As Boolean
'     Dim oneType As String

'     'Initialize Collection and Dictionary
'     'i=1, reportDataDate
'     'i=2, reportMonth
'     'i=3, reportMonthString
'     'i=4, FilePaths
'     Set configCollection = New Collection
    
'     'Initialize Dictionary
'     Set filePathDict = CreateObject("Scripting.Dictionary")
'     Set dbs = CurrentDb()

'     'Fetch FolderDir and ReportDataDate
'     Set rsConfig = dbs.OpenRecordset("SELECT TOP 1 ReportDataDate, ReportMonth, ReportMonthString, CopyFilePath FROM Configuration ORDER BY DataID DESC", dbOpenSnapshot)
    
'     If Not rsConfig.EOF Then
'         reportDataDate = rsConfig!ReportDataDate
'         reportMonth = rsConfig!ReportMonth
'         reportMonthString = rsConfig!ReportMonthString
'         copyFilePath = rsConfig!CopyFilePath
'     Else
'         Set GetConfigsReturnCollection = configCollection
'         MsgBox "Error: 無法取得Configuration資料表資料", vbCritical

'         rsConfig.Close
'         Set rsConfig = Nothing
'         Exit Function
'     End If

'     rsConfig.Close
'     Set rsConfig = Nothing
    
'     configCollection.add reportDataDate
'     configCollection.add reportMonth
'     configCollection.add reportMonthString

'     'Fetch FilePaths
'     Set rsFilePaths = dbs.OpenRecordset("SELECT PathID, PathName, ExtensionFormat FROM FilePaths", dbOpenSnapshot)
    
'     'Check if optional parameters exist
'     checkOpt = Not (IsMissing(reportType) Or IsMissing(reportTypeDate))


'     Do While Not rsFilePaths.EOF
'         pathID = Nz(rsFilePaths!PathID, "")
'         pathName = Nz(rsFilePaths!PathName, "")
'         extensionFormat = Nz(rsFilePaths!ExtensionFormat, "txt")

'         If Not checkOpt Then
'             fullFilePath = copyFilePath & "\" & pathName & Format(reportDataDate, "yyyymmdd") & "." & extensionFormat
'         Else
'             For Each oneType In reportType
'                 If oneType = pathID Then
'                     fullFilePath = copyFilePath & "\" & pathName & Format(reportTypeDate, "yyyymmdd") & "." & extensionFormat
'                 End If
'             Next oneType
'         End If
    
'         'Add to dictionary:
'         'PathID as Key，fullFilePath as Value
'         filePathDict.Add pathID, fullFilePath
'         rsFilePaths.MoveNext
'     Loop
    
'     rsFilePaths.Close
'     Set rsFilePaths = Nothing
'     Set dbs = Nothing

'     configCollection.add filePathDict 
'     Set GetConfigsReturnCollection = configCollection
' End Function


'---------------------------
'Save Object with Dictionary
'---------------------------
Public Function GetConfigsReturnDict(Optional ByVal reportType As Variant, _
                           Optional ByVal reportTypeDate As Date) As Object
    'Define Database and Recordset
    Dim dbs As DAO.Database
    Dim rsConfig As DAO.Recordset
    Dim rsFilePaths As DAO.Recordset

    'Dictionary for Configuration
    Dim configDict As Object
    Dim filePathDict As Object

    'Table Configuration Variables
    Dim reportDataDate As Date
    Dim reportMonth As Date
    Dim reportMonthString As String
    Dim copyFilePath As String

    'Table FilePaths Variables
    Dim pathID As String
    Dim pathName As String
    Dim extensionFormat As String
    Dim fullFilePath As String

    'Check Optional Parameters
    Dim checkOpt As Boolean
    Dim oneType As Variant

    'Initialize Dictionaries
    Set configDict = CreateObject("Scripting.Dictionary")
    Set filePathDict = CreateObject("Scripting.Dictionary")
    Set dbs = CurrentDb()

    'Fetch FolderDir and ReportDataDate
    Set rsConfig = dbs.OpenRecordset("SELECT TOP 1 ReportDataDate, ReportMonth, ReportMonthString, CopyFilePath FROM Configuration ORDER BY DataID DESC", dbOpenSnapshot)

    If Not rsConfig.EOF Then
        reportDataDate = rsConfig!ReportDataDate
        reportMonth = rsConfig!ReportMonth
        reportMonthString = rsConfig!ReportMonthString
        copyFilePath = rsConfig!CopyFilePath
    Else
        Set GetConfigsReturnDict = configDict
        MsgBox "Error: 無法取得 Configuration 資料表資料", vbCritical
        WriteLog "Error: 無法取得 Configuration 資料表資料"
        rsConfig.Close
        Set rsConfig = Nothing
        Exit Function
    End If

    rsConfig.Close
    Set rsConfig = Nothing

    ' Add values to dictionary
    configDict.Add "ReportDataDate", reportDataDate
    configDict.Add "ReportMonth", reportMonth
    configDict.Add "ReportMonthString", reportMonthString
    configDict.Add "CopyFilePath", copyFilePath

    'Fetch FilePaths
    Set rsFilePaths = dbs.OpenRecordset("SELECT PathID, PathName, ExtensionFormat FROM FilePaths", dbOpenSnapshot)

    'Check if optional parameters exist
    checkOpt = Not (IsMissing(reportType) Or IsMissing(reportTypeDate))

    Do While Not rsFilePaths.EOF
        pathID = Nz(rsFilePaths!PathID, "")
        pathName = Nz(rsFilePaths!PathName, "")
        extensionFormat = Nz(rsFilePaths!ExtensionFormat, "txt")

        If Not checkOpt Then
            fullFilePath = copyFilePath & "\" & pathName & Format(reportDataDate, "yyyymmdd") & "." & extensionFormat

            '針對外幣債評估表要做例外處理
            If pathID = "FXDebtEvaluation" Then fullFilePath = copyFilePath & "\" & Format(reportDataDate, "yyyymmdd") & pathName & "." & extensionFormat

            ' If pathID = "OBU_DL6320" Or pathID = "DBU_CM2810" Or pathID = "DBU_DL9360" Or pathID = "DBU_DL6850" Or pathID = "OBU_CF6320" Then fullFilePath = copyFilePath & "\" & pathName & "." & extensionFormat
        Else
            For Each oneType In reportType
                If oneType = pathID Then fullFilePath = copyFilePath & "\" & pathName & Format(reportTypeDate, "yyyymmdd") & "." & extensionFormat
                
                '針對外幣債評估表要做例外處理
                If (oneType =  pathID) And (pathID = "FXDebtEvaluation") Then fullFilePath = copyFilePath & "\" & Format(reportTypeDate, "yyyymmdd") & pathName & "." & extensionFormat

                ' If (oneType =  pathID) And ((pathID = "OBU_DL6320")  Or (pathID = "DBU_CM2810") Or (pathID = "DBU_DL9360") Or (pathID = "DBU_DL6850") Or (pathID = "OBU_CF6320")) Then fullFilePath = copyFilePath & "\" & pathName & "." & extensionFormat
                
            Next oneType
        End If

        'Add to dictionary: PathID as Key, fullFilePath as Value
        filePathDict.Add pathID, fullFilePath
        rsFilePaths.MoveNext
    Loop

    rsFilePaths.Close
    Set rsFilePaths = Nothing
    Set dbs = Nothing

    ' Add FilePaths dictionary to main dictionary
    configDict.Add "FilePaths", filePathDict

    ' Return dictionary
    Set GetConfigsReturnDict = configDict
End Function

Public Function GetTableColumns(ByVal cleaningType As String) As Variant
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim tableColumns() As String
    Dim i As Integer

    '取得 DAO 資料庫物件
    Set dbs = CurrentDb

    '查詢資料庫中的 TableName 和 DbCol
    ' Set rs = dbs.OpenRecordset("SELECT TableName, DbCol FROM DBsColTable WHERE TableName = '" & cleaningType & "'")
    Set rs = dbs.OpenRecordset("SELECT TableName, DbCol FROM DBsColTable WHERE TableName = '" & cleaningType & "'" & "ORDER BY DataID ASC;")

    '如果有資料，將 DbCol 放入陣列中
    If Not rs.EOF Then
        i = 0
        rs.MoveFirst
        Do While Not rs.EOF
            ReDim Preserve tableColumns(i)
            tableColumns(i) = rs!DbCol
            i = i + 1
            rs.MoveNext
        Loop
    End If

    '關閉 Recordset
    rs.Close
    Set rs = Nothing
    Set dbs = Nothing

    '傳回陣列
    GetTableColumns = tableColumns
End Function
        
' Compute Last Workday of each month
Public Function GetLastWorkday(ByVal yr As Integer, ByVal mth As Integer) As Date
    Dim lastDay As Date
    Dim isHoliday As Boolean
    Dim rs As DAO.Recordset
    Dim sqlString As String

    'Set first day of month
    lastDay = DateSerial(yr, mth + 1, 0)
    
    Do    
        isHoliday = False
        sqlString = "SELECT COUNT(Date) AS cnt FROM Holidays WHERE Date = #" & lastDay & "#"
        Set rs = CurrentDb.OpenRecordset(sqlString)
        If Not rs.EOF Then isHoliday = (rs!cnt > 0)
        rs.Close
        Set rs = Nothing
        'If is holiday then day + 1
        If isHoliday Then lastDay = lastDay - 1
    Loop While isHoliday  ' Is holiday then continue loop
    
    GetLastWorkday = lastDay
End Function


Public Sub CopyRawFileData(ByVal rawFilePath as string, ByVal copyFilePath as string)
    Dim fso As Object
    Dim rawDataFolder As Object
    Dim file As Object

    'Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Check whether raw data exist or not
    If Not fso.FolderExists(rawFilePath) Then
        MsgBox "原始資料夾不存在", vbExclamation
        WriteLog "原始資料夾不存在"
        Exit Sub
    End If

    Set rawDataFolder = fso.GetFolder(rawFilePath)
    
    'Copy files to destination folder
    For Each file In rawDataFolder.Files
        file.Copy copyFilePath & "\" & file.Name
    Next file
    MsgBox "原始檔案路徑: " & rawDataFolder & "| 成功複製到路徑: " & copyFilePath 
    WriteLog "原始檔案路徑: " & rawDataFolder & "| 成功複製到路徑: " & copyFilePath 
End Sub


Public Sub ConfigureExcelApp(ByRef xlApp As Excel.Application)
    With xlApp
        .Visible = False             ' 不顯示 Excel 畫面
        .ScreenUpdating = False      ' 關閉畫面更新，提高效能
        .DisplayAlerts = False       ' 關閉警告訊息，自動接受預設
        .AskToUpdateLinks = False    ' 關閉連結更新詢問
    End With
End Sub

Public Sub RestoreExcelAppSettings(ByRef xlApp As Excel.Application)
    With xlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
    End With
End Sub

Public Sub ClearCutCopyMode(ByRef xlApp As Excel.Application)
    xlApp.CutCopyMode = False
End Sub
