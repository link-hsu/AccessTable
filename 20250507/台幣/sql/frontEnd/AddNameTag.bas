Option Explicit

Sub NameMultipleSheetCellsByValue_NoRestrictions()
    Dim config As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim area As Range, cell As Range
    Dim rngAll As Range
    Dim nm As String
    
    '=== Step 1: 定義要處理的工作表與範圍 ===
    config = Array( _
        Array("TABLE10", "S2:S21,W2:W21,AA2:AA21,AE2:AE21,AI2:AI21,AM2:AM21,AQ2:AQ21"), _
        Array("TABLE15A", "S2:S10,W2:W10,AA2:AA10,AE2:AE10,AI2:AI10"), _
        Array("TABLE20", "S2:S5,W2:W5,AA2:AA5"), _
        Array("TABLE22", "S2:S5,W2:W5"), _
        Array("TABLE23", "S2:S6"), _
        Array("TABLE24", "S2:S14,W2:W14,AA2:AA14,AE2:AE14,AI2:AI14,AM2:AM14,AQ2:AQ14,AU2:AU14"), _
        Array("TABLE27", "S2:S7,W2:W7,AA2:AA7,AE2:AE7,AI2:AI7"), _
        Array("TABLE36", "S2:S4,W2:W4,AA2:AA4"), _
        Array("AI233", "S2:S5,S10:S15,W2:W5,W10:W15,AA2:AA5,AA10:AA15,AE2:AE9,AE12:AE17,AI2:AI5,AM2:AM5,AQ2:AQ5,AU2:AU9"), _
        Array("AI405", "S2:S5,W2:W5"), _
        Array("AI410", "S2:S8,W2:W8"), _
        Array("AI430", "S2:S8"), _
        Array("AI601", "S2:S65,W2:W48,AA2:AA48,AE2:AE48,AI2:AI48,AM2:AM48,AQ2:AQ48,AU2:AU48,AY2:AY48,BC2:BC48,BG2:BG48,BK2:BK48") _
        Array("AI605", "S2:S3,S5:S6,W2:W3,W5:W6,AA2:AA3,AE2:AE3,AI2:AI3,AM2:AM3,AQ2:AQ3,AU2:AU3") _
    )


    configAddr = Array( _
    Array("TABLE10", "U2:U21,Y2:Y21,AC2:AC21,AG2:AG21,AK2:AK21,AO2:AO21,AS2:AS21"), _
    Array("TABLE15A", "U2:U10,Y2:Y10,AC2:AC10,AG2:AG10,AK2:AK10"), _
    Array("TABLE20", "U2:U5,Y2:Y5,AC2:AC5"), _
    Array("TABLE22", "U2:U5,Y2:Y5"), _
    Array("TABLE23", "U2:U6"), _
    Array("TABLE24", "U2:U14,Y2:Y14,AC2:AC14,AG2:AG14,AK2:AK14,AO2:AO14,AS2:AS14,AW2:AW14"), _
    Array("TABLE27", "U2:U7,Y2:Y7,AC2:AC7,AG2:AG7,AK2:AK7"), _
    Array("TABLE36", "U2:U4,Y2:Y4,AC2:AC4"), _

    Array("AI233", "U2:U5,U10:U15,Y2:Y5,Y10:Y15,AC2:AC5,AC10:AC15,AG2:AG9,AG12:AG17,AK2:AK5,AO2:AO5,AS2:AS5,AW2:AW9"), _
    Array("AI405", "U2:U5,Y2:Y5"), _
    Array("AI410", "U2:U8,Y2:Y8"), _
    Array("AI430", "U2:U8"), _
    Array("AI601", "U2:U65,Y2:Y48,AC2:AC48,AG2:AG48,AK2:AK48,AO2:AO48,AS2:AS48,AW2:AW48,BA2:BA48,BE2:BE48,BI2:BI48,BM2:BM48") _
    Array("AI605", "U2:U3,U5:U6,Y2:Y3,Y5:Y6,AC2:AC3,AG2:AG3,AK2:AK3,AO2:AO3,AS2:AS3,AW2:AW3") _
)
    
    '=== Step 2: 逐一處理每個工作表與其對應範圍 ===
    For i = LBound(config) To UBound(config)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(config(i)(0))
        If ws Is Nothing Then
            MsgBox "找不到工作表：" & config(i)(0), vbExclamation
            Err.Clear
            GoTo NextConfig
        End If
        
        Set rngAll = ws.Range(config(i)(1))
        If rngAll Is Nothing Then
            MsgBox "無效範圍：" & config(i)(1) & "（工作表：" & ws.Name & "）", vbExclamation
            Err.Clear
            GoTo NextConfig
        End If
        On Error GoTo 0
        
        '=== Step 3: 命名每個儲存格（無名稱驗證）===
        For Each area In rngAll.Areas
            For Each cell In area.Cells
                nm = Trim(cell.Value)
                
                If Len(nm) = 0 Then GoTo SkipCell
                
                On Error Resume Next
                ThisWorkbook.Names.Add Name:=nm, RefersTo:="=" & "'" & ws.Name & "'!" & cell.Address
                If Err.Number <> 0 Then
                    MsgBox "命名失敗：""" & nm & """，位置：" & ws.Name & "!" & cell.Address & _
                           vbCrLf & "錯誤：" & Err.Description, vbCritical
                    Err.Clear
                End If
                On Error GoTo 0
                
SkipCell:
            Next cell
        Next area
        
NextConfig:
        Set ws = Nothing
        Set rngAll = Nothing
    Next i
    
    MsgBox "全部處理完成！", vbInformation
End Sub
