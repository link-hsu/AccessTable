Option Explicit

Function GetLogFileName() As String
    Dim folderPath As String
    Dim uuid As String
    Dim fileName As String
    
    folderPath = ThisWorkbook.Path & "\LogFile_Frontend\"  ' 你也可以指定其他資料夾
    uuid = CreateUUID()
    fileName = "LogFile_" & Format(Now, "yyyymmdd_hhnnss") & "_" & uuid & ".txt"
    
    GetLogFileName = folderPath & fileName
End Function

' 模擬UUID
Public Function CreateUUID() As String
    Randomize
    CreateUUID = Format(Now, "hhmmss") & _
                    Hex(Int(Rnd() * 65536)) & _
                    Hex(Int(Rnd() * 65536))
End Function

Sub WriteLog(logMessage As String, _
             Optional logFilePath As String = "")             
    Static logFile As String
    
    If logFilePath <> "" Then
        logFile = logFilePath
    ElseIf logFile = "" Then
        logFile = GetLogFileName()
    End If

    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & logMessage
    Close #fileNum
End Sub
