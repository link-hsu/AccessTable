
Option Explicit

' ''' === LOG ADD START（沿用） ===
Private Const LOG_FOLDER As String = "logs"
Private Const LOG_FILE_BASENAME As String = "parse_txt.log"
Private Const LOG_ENABLED As Boolean = True ' 要關閉 log 改成 False

Private Sub LogInfo(ByVal msg As String): WriteLog "INFO", msg: End Sub
Private Sub LogWarn(ByVal msg As String): WriteLog "WARN", msg: End Sub
Private Sub LogError(ByVal msg As String): WriteLog "ERROR", msg: End Sub

Sub 整理外幣債_依原始標題自動拆欄()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim 起始列 As Long, 結束列 As Long
    Dim i As Long, col As Long, destRow As Long, destCol As Long
    Dim 部位 As String
    Dim 合併欄位 As Variant
    Dim 原始標題 As String, 標題Parts As Variant
    Dim isHeaderRow As Boolean
    ' === NEW: 案號分頁與當前債券編號追蹤 ===
    Dim wsTradeCases As Worksheet
    Dim tradeCaseRow As Long
    Dim current債券編號 As String
    Dim currentCurrency As String

    Set wsSrc = ThisWorkbook.Sheets("評估表")
    起始列 = 6
    結束列 = 220

    ' 建立或取得整理後資料工作表
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("整理後資料")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = "整理後資料"
    End If
    On Error GoTo 0
    wsDest.Range("A:AT").ClearContents

    ' === NEW: 建立或取得「案號」工作表並清空 ===
    On Error Resume Next
    Set wsTradeCases = ThisWorkbook.Sheets("整理後資料_交易單號")
    If wsTradeCases Is Nothing Then
        Set wsTradeCases = ThisWorkbook.Sheets.Add(After:=wsDest)
        wsTradeCases.Name = "整理後資料_交易單號"
    End If
    On Error GoTo 0
    ' wsTradeCases.Cells.Clear
    Union(wsTradeCases.Range("A:D"), wsTradeCases.Range("H:O")).ClearContents
    ' 標題列：A=案號(19/20欄)，B=債券編號(第3欄)
    wsTradeCases.Cells(1, 1).Value = "交易單號"
    wsTradeCases.Cells(1, 2).Value = "Security_Id"
    wsTradeCases.Cells(1, 3).Value = "Ccy"
    wsTradeCases.Cells(1, 4).Value = "部位"
    wsTradeCases.Cells(1, 5).Value = "成本_原幣"
    wsTradeCases.Cells(1, 6).Value = "評價調整_原幣"
    wsTradeCases.Cells(1, 7).Value = "利息_原幣"
    wsTradeCases.Cells(1, 8).Value = "成本_美元"
    wsTradeCases.Cells(1, 9).Value = "評價調整_美元"
    wsTradeCases.Cells(1, 10).Value = "利息_美元"
    wsTradeCases.Cells(1, 11).Value = "成本科目代號"
    wsTradeCases.Cells(1, 12).Value = "評價調整科目代號"
    wsTradeCases.Cells(1, 13).Value = "利息科目代號"
    wsTradeCases.Cells(1, 14).Value = "成本科目名稱"
    wsTradeCases.Cells(1, 15).Value = "評價調整科目名稱"    
    wsTradeCases.Cells(1, 16).Value = "利息科目名稱"
    tradeCaseRow = 2
    current債券編號 = ""

    ' 設定要合併拆為兩欄的欄位
    合併欄位 = Array(1, 2, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 26, 27)

    ' 輸出標題列
    wsDest.Cells(1, 1).Value = "部位"
    destCol = 2
    For col = 1 To 30
        原始標題 = wsSrc.Cells(5, col).Value
        標題Parts = Split(原始標題, Chr(10)) ' 用換行分割
        If IsInArray(col, 合併欄位) Then
            If UBound(標題Parts) >= 1 Then
                wsDest.Cells(1, destCol).Value = Trim(標題Parts(0))
                wsDest.Cells(1, destCol + 1).Value = Trim(標題Parts(1))
            Else
                wsDest.Cells(1, destCol).Value = 原始標題
                wsDest.Cells(1, destCol + 1).Value = 原始標題
            End If
            destCol = destCol + 2
        Else
            wsDest.Cells(1, destCol).Value = 原始標題
            destCol = destCol + 1
        End If
    Next col

    ' 輸出資料列
    destRow = 2

    For i = 起始列 To 結束列 - 1
        ' 檢查是否為標題列 (比對第3、4欄)
        isHeaderRow = (Trim(wsSrc.Cells(i, 3).Text) = Trim(wsSrc.Cells(5, 3).Text) And _
                       Trim(wsSrc.Cells(i, 4).Text) = Trim(wsSrc.Cells(5, 4).Text))
        If isHeaderRow Then GoTo SkipCaseRow

        ' 先檢查這一列是否是部位列 (有"-"而且第二欄空白)
        If InStr(wsSrc.Cells(i, 1).Value, "-") > 0 And wsSrc.Cells(i, 2).Value = "" Then
            部位 = wsSrc.Cells(i, 1).Value
            GoTo SkipCaseRow ' 部位列自己不需要輸出資料
        End If

        ' === NEW: 追蹤第3欄「債券編號」，並把第23/24欄(一對多)映射到「案號」分頁 ===
        ' 若本列第3欄有值，更新目前的債券編號
        If wsSrc.Cells(i, 23).Value = "" And wsSrc.Cells(i, 24).Value = "" Then
            GoTo SkipCaseRow ' 部位列自己不需要輸出資料
        End If            

        If Trim$(wsSrc.Cells(i, 3).Value & "") <> "" Then
            current債券編號 = Trim$(wsSrc.Cells(i, 1).Value & "")
            currentCurrency = Trim$(wsSrc.Cells(i, 2).Value & "")
        End If

        ' 僅在已取得債券編號時，輸出第23/24欄的案號至「案號」分頁
        If current債券編號 <> "" Then
            Dim C23_value As String, C24_value As String
            C23_value = Trim$(wsSrc.Cells(i, 23).Value & "")
            C24_value = Trim$(wsSrc.Cells(i, 24).Value & "")

            ' 若第19欄有值，新增一列：A=案號(19欄)，B=債券編號
            If C23_value <> "" Then
                wsTradeCases.Cells(tradeCaseRow, 1).Value = C23_value
                wsTradeCases.Cells(tradeCaseRow, 2).Value = current債券編號
                wsTradeCases.Cells(tradeCaseRow, 3).Value = currentCurrency
                wsTradeCases.Cells(tradeCaseRow, 4).Value = 部位
                tradeCaseRow = tradeCaseRow + 1
            End If

            ' 若第20欄有值，新增一列：A=案號(20欄)，B=債券編號
            If C24_value <> "" Then
                wsTradeCases.Cells(tradeCaseRow, 1).Value = C24_value
                wsTradeCases.Cells(tradeCaseRow, 2).Value = current債券編號
                wsTradeCases.Cells(tradeCaseRow, 3).Value = currentCurrency
                wsTradeCases.Cells(tradeCaseRow, 4).Value = 部位
                tradeCaseRow = tradeCaseRow + 1
            End If
        End If

        SkipCaseRow:
    Next i        

    isHeaderRow = False

    For i = 起始列 To 結束列 - 1
        ' 檢查是否為標題列 (比對第3、4欄)
        isHeaderRow = (Trim(wsSrc.Cells(i, 3).Text) = Trim(wsSrc.Cells(5, 3).Text) And _
                       Trim(wsSrc.Cells(i, 4).Text) = Trim(wsSrc.Cells(5, 4).Text))
        If isHeaderRow Then GoTo SkipRow

        ' 先檢查這一列是否是部位列 (有"-"而且第二欄空白)
        If InStr(wsSrc.Cells(i, 1).Value, "-") > 0 And wsSrc.Cells(i, 2).Value = "" Then
            部位 = wsSrc.Cells(i, 1).Value
            GoTo SkipRow ' 部位列自己不需要輸出資料
        End If

        ' 如果上下兩列都有資料，處理資料搬移
        ' === CHANGED: 僅在第1欄上下兩列皆有值，且第23與第24欄皆有值時，才做主要資料搬移 ===
        If wsSrc.Cells(i, 1).Value <> "" And _
           wsSrc.Cells(i + 1, 1).Value <> "" Then
            wsDest.Cells(destRow, 1).Value = 部位 ' 這列的部位
            destCol = 2
            For col = 1 To 30
                If IsInArray(col, 合併欄位) Then
                    wsDest.Cells(destRow, destCol).Value = wsSrc.Cells(i, col).Value
                    wsDest.Cells(destRow, destCol + 1).Value = wsSrc.Cells(i + 1, col).Value
                    destCol = destCol + 2
                Else
                    wsDest.Cells(destRow, destCol).Value = wsSrc.Cells(i, col).Value
                    destCol = destCol + 1
                End If
            Next col
            destRow = destRow + 1
            i = i + 1 ' 因為已經處理了兩列，要多跳一列
        End If

SkipRow:
    Next i
    
    ' 取消整理後資料工作表的所有自動換列
    Dim rng As Range, cell As Range
    Set rng = wsDest.UsedRange.SpecialCells(xlCellTypeConstants)

    If Not rng Is Nothing Then
        For Each cell In rng
            If Not IsEmpty(cell) Then
                cell.Value = Replace(cell.Value, vbLf, " ") ' 清除換行字元
                cell.WrapText = False ' 關閉儲存格自動換列
            End If
        Next cell
    End If
    ThisWorkbook.Sheets("整理後資料").Activate
    MsgBox "已完成整理，總筆數：" & destRow - 2, vbInformation
End Sub

' 判斷某欄是否在合併欄位清單中
Function IsInArray(val As Long, arr As Variant) As Boolean
    Dim v
    For Each v In arr
        If v = val Then
            IsInArray = True
            Exit Function
        End If
    Next v
    IsInArray = False
End Function

' =========================
' 解析：項目 為 String() 陣列版（含檔案 Log）
' =========================

Private Function GetLogFullPath() As String
    Dim base As String
    base = ThisWorkbook.Path & Application.PathSeparator & LOG_FOLDER
    MkDirRecursive base
    GetLogFullPath = base & Application.PathSeparator & LOG_FILE_BASENAME
End Function

Private Sub MkDirRecursive(ByVal folderPath As String)
    Dim fso As Object
    Dim parent As String
    If Len(folderPath) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 統一路徑分隔符並移除尾端分隔
    folderPath = Replace(folderPath, "/", Application.PathSeparator)
    If Right$(folderPath, 1) = Application.PathSeparator Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    If Len(folderPath) = 0 Then Exit Sub

    ' 若已存在就結束
    If fso.FolderExists(folderPath) Then Exit Sub

    ' 先確保上層存在（到根為止）
    parent = fso.GetParentFolderName(folderPath)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then
        MkDirRecursive parent
    End If

    ' 建立目前層
    On Error Resume Next
    fso.CreateFolder folderPath
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal level As String, ByVal msg As String)
    If Not LOG_ENABLED Then Exit Sub
    Dim fnum As Integer, path As String
    path = GetLogFullPath()
    fnum = FreeFile
    On Error Resume Next
    Open path For Append As #fnum
    Print #fnum, Format$(Now, "yyyy-mm-dd HH:NN:SS"); " ["; level; "] "; msg
    Close #fnum
    On Error GoTo 0
End Sub

' 讀取 txt，依你的簡化規則回傳：
' Collection 其內每個元素為 Dictionary，鍵：
'   "會科名稱" 或 "會計名稱" As String   '（注意：你後文用「會計名稱」，下方已做相容）
'   "會計科目" As String                 ' 僅數字（已去破折號）
'   "項目"     As Variant                 ' String() 動態陣列
Public Function ParseTxt_ArrayVersion(ByVal filePath As String) As Collection
    On Error GoTo ErrHandler

    Dim fso As Object, ts As Object
    Dim col As New Collection            ' 回傳集合（保留首次出現順序）
    Dim index As Object                  ' key: acctDigits|subjectName -> Dictionary
    Dim curKey As String                 ' 目前小節 key
    Dim lineRaw As String
    Dim line As String                   ' ''' === FIX FF/BOM CLEAN: 新增「淨化後」變數 ===
    Dim d As Object                      ' Scripting.Dictionary
    Dim key As String
    Dim acctDigits As String, subjectName As String
    Dim firstTok As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -2) ' 1=ForReading, -2=系統預設編碼
    Set index = CreateObject("Scripting.Dictionary")

    LogInfo "Start parsing: " & filePath
    curKey = ""

    Do While Not ts.AtEndOfStream
        lineRaw = ts.ReadLine

        ' ''' === FIX FF/BOM CLEAN: 先清除控制字元/零寬/BOM/全形空白為半形 ===
        line = SanitizeLine(lineRaw)

        ' (1) 空白行不要（已清完怪字元）
        If Trim$(line) = "" Then GoTo ContinueLoop

        ' (2) 碰到「科目:」→ 新小節（直到下一個科目）
        If IsSubjectLine_Simple(line) Then
            ParseSubject_Simple line, acctDigits, subjectName
            key = acctDigits & "|" & subjectName

            If index.Exists(key) Then
                curKey = key
                LogInfo "Subject continues (merge): [" & acctDigits & "] " & subjectName
            Else
                Set d = CreateObject("Scripting.Dictionary")
                ' ''' === NEW（為相容你敘述的「會計名稱」）：同值存兩鍵 ===
                d("會科名稱") = subjectName
                d("會計名稱") = subjectName   ' === NEW: 相容鍵 ===
                d("會計科目") = acctDigits
                d("項目") = EmptyStringArray()
                index.Add key, d
                col.Add d
                curKey = key
                LogInfo "New subject: [" & acctDigits & "] " & subjectName
            End If
            GoTo ContinueLoop
        End If

        ' 尚未進入任何小節 → 跳過
        If curKey = "" Then GoTo ContinueLoop

        ' (3) 小節內擷取規則：
        '     行首非空白，且不是「33 OBU」與「帳號」開頭 → 把第一個字詞放入項目
        If StartsWithNonSpace(line) Then
            Dim lineAfterTrim As String
            lineAfterTrim = TrimLeftAllSpaces(line) ' 已無怪字元

            If Not StartsWith33OBU(lineAfterTrim) _
               And Not StartsWithHeader(lineAfterTrim) Then

                firstTok = FirstToken(lineAfterTrim)
                If firstTok <> "" Then
                    Set d = index(curKey)
                    d("項目") = PushString(d("項目"), firstTok)
                    LogInfo "Add item to [" & d("會計科目") & "] " & d("會科名稱") & " : " & firstTok
                End If
            End If
        End If

ContinueLoop:
    Loop
    ts.Close

    ' 收尾統計
    Dim i As Long, a As Variant, cnt As Long
    LogInfo "Sections: " & col.Count
    For i = 1 To col.Count
        Set d = col(i)
        a = d("項目")
        If IsEmptyArray(a) Then
            cnt = 0
        Else
            cnt = UBound(a) - LBound(a) + 1
        End If
        LogInfo "Summary [" & d("會計科目") & "] " & d("會科名稱") & " -> items=" & cnt
    Next

    Set ParseTxt_ArrayVersion = col
    LogInfo "Parsing done."
    Exit Function

ErrHandler:
    LogError "ParseTxt_ArrayVersion error: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ====== 規則與工具（保持精簡） ======

' 是否為「科目:」行
Private Function IsSubjectLine_Simple(ByVal s As String) As Boolean
    IsSubjectLine_Simple = (InStr(1, s, "科目:", vbTextCompare) > 0)
End Function

' 解析科目行 → 會計科目(只留數字) + 會科名稱
Private Sub ParseSubject_Simple(ByVal s As String, ByRef acctDigits As String, ByRef subjectName As String)
    Dim p As Long, rest As String, sp As Long, codeRaw As String
    p = InStr(1, s, "科目:", vbTextCompare)
    If p = 0 Then acctDigits = "": subjectName = "": Exit Sub

    rest = Trim$(Mid$(s, p + Len("科目:")))
    sp = InStr(1, rest, " ")
    If sp > 0 Then
        codeRaw = Left$(rest, sp - 1)
        subjectName = Trim$(Mid$(rest, sp + 1))
    Else
        codeRaw = rest
        subjectName = ""
    End If
    acctDigits = DigitsOnly_Simple(codeRaw)
End Sub

' 只留數字
Private Function DigitsOnly_Simple(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next
    DigitsOnly_Simple = out
End Function

' 行首非空白（空白=半形空白/Tab/全形空白/非斷行空白等在 Sanitize 後都變成半形空白）
Private Function StartsWithNonSpace(ByVal s As String) As Boolean
    If s = "" Then StartsWithNonSpace = False: Exit Function
    StartsWithNonSpace = (Left$(s, 1) <> " ")
End Function

' 左邊去空白（Sanitize 後只需 LTrim）
Private Function TrimLeftAllSpaces(ByVal s As String) As String
    TrimLeftAllSpaces = LTrim$(s)
End Function

' 33 OBU 判斷（Sanitize 後直接判斷開頭）
Private Function StartsWith33OBU(ByVal s As String) As Boolean
    StartsWith33OBU = (InStr(1, LCase$(s), "33 obu", vbTextCompare) = 1)
End Function

' 是否為表頭（以「帳號」開頭）
Private Function StartsWithHeader(ByVal s As String) As Boolean
    StartsWithHeader = (Left$(TrimLeftAllSpaces(s), 2) = "帳號")
End Function

' 取第一個字詞（Sanitize 後只有半形空白）
Private Function FirstToken(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s & " ", " ")
    FirstToken = Left$(s, p - 1)
End Function

' 去掉/轉換所有會干擾判斷的隱形字元（FF、BOM、零寬空白、NBSP、全形空白、控制字元等）
Private Function SanitizeLine(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long
    Dim out As String

    If LenB(s) = 0 Then SanitizeLine = "": Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)

        ' 控制字元 → 空白（含 FF 0x0C）
        If code >= 0 And code < 32 Then
            out = out & " "
        ' 零寬/BOM/NBSP/窄不換行空白
        ElseIf code = &HFEFF Or code = &H200B Or code = &H200C Or code = &H200D _
            Or code = &H2060 Or code = &H00A0 Or code = &H202F Then
            out = out & " "
        ' 全形空白
        ElseIf code = &H3000 Then
            out = out & " "
        Else
            out = out & ch
        End If
    Next

    ' 壓連續空白
    Do While InStr(out, "  ") > 0
        out = Replace$(out, "  ", " ")
    Loop

    SanitizeLine = out
End Function

' === 動態字串陣列輔助 ===
' 建立「空」的狀態（用 Empty 代表未初始化陣列）
Private Function EmptyStringArray() As Variant
    EmptyStringArray = Empty
End Function

' 判斷陣列是否為「空字串陣列」或未初始化
Private Function IsEmptyArray(ByVal arr As Variant) As Boolean
    If Not IsArray(arr) Then
        IsEmptyArray = True
        Exit Function
    End If
    On Error GoTo EH
    IsEmptyArray = (UBound(arr) < LBound(arr))
    Exit Function
EH:
    IsEmptyArray = True
End Function

' 追加一個字串到 String() 陣列，並回傳新陣列
Private Function PushString(ByVal arr As Variant, ByVal s As String) As Variant
    Dim n As Long
    If IsEmptyArray(arr) Then
        Dim a() As String
        ReDim a(0 To 0)
        a(0) = s
        PushString = a
    Else
        n = UBound(arr) + 1
        ReDim Preserve arr(0 To n)
        arr(n) = s
        PushString = arr
    End If
End Function

' ===========================================================
' === NEW: 將 Collection 寫入到 Excel（依你的 1/2/3 規則）===
' ===========================================================

' ''' === NEW: 清 Row1 並填入標題、刪除 A 欄「合計/主管/空白」列（自下而上）、再依序填入資料 ===
Public Sub ApplyCollectionToSheet(ByVal col As Collection, ByVal ws As Worksheet)
    Dim lastRow As Long, r As Long
    Dim i As Long, d As Object, items As Variant, j As Long
    Dim acct As String, subj As String

    If ws Is Nothing Then Set ws = ActiveSheet

    Application.ScreenUpdating = False
    ws.Range("F:F,H:M").Delete
    ' (1) 先清除 Row1 所有資料並填入欄位名稱
    '     A1=帳號, B1=客戶代號, C1=AO代號, D1=子目或戶名, E1=COUNTY CODE 摘要, F1=金額
    ' ''' === NEW: 清除並設置標題 ===
    ws.Rows(1).ClearContents
    ws.Range("A1").Value = "帳號"
    ws.Range("B1").Value = "客戶代號"
    ws.Range("C1").Value = "AO代號"
    ws.Range("D1").Value = "子目或戶名"
    ws.Range("E1").Value = "COUNTRY CODE 摘要"
    ws.Range("F1").Value = "金額"

    ' 先清理舊資料（若工作表先前有殘留）
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    ' (2) 刪除 A 欄位包含「合計」、「主管」以及空白的列（從最後一行至第 1 行）
    '     依你的敘述：從最後一行為起始點至第一行
    ' ''' === NEW: 自下而上刪列 ===
    For r = lastRow To 1 Step -1
        Dim aVal As String
        aVal = Trim$(CStr(ws.Cells(r, "A").Value))
        If aVal = "" Or InStr(1, aVal, "合計", vbTextCompare) > 0 Or InStr(1, aVal, "主管", vbTextCompare) > 0 Then
            ws.Rows(r).Delete
        End If
    Next r    

    ws.Columns("A:C").Insert Shift:=xlToRight
    ' 重新定位寫入起始列（標題在第 1 列，資料從第 2 列）
    r = 2

    ' (3) 依 Collection 逐一寫入：
    '     對每個 Dictionary：取 "項目" 個數 N，逐筆輸出 N 列
    '     每列：A=會計科目、B=會計名稱（或會科名稱）、C=項目值
    '     （你舉 3 筆/5 筆為例，這裡做成一般化：N 筆就寫 N 列）
    ' ''' === NEW: 逐段寫入 ===
    For i = 1 To col.Count
        Set d = col(i)
        items = d("項目")
        acct = NzDict(d, "會計科目", "")
        ' 你上文使用「會計名稱」一詞，為相容此處優先取「會計名稱」，否則用「會科名稱」
        subj = NzDict(d, "會計名稱", NzDict(d, "會科名稱", ""))

        If Not IsEmptyArray(items) Then
            For j = LBound(items) To UBound(items)
                ' 填入：A=會計科目、B=會計名稱、C=項目值
                ws.Cells(r, "A").Value = acct
                ws.Cells(r, "B").Value = subj
                ws.Cells(r, "C").Value = CStr(items(j))
                r = r + 1
            Next j
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

' ''' === NEW: 字典取字串的小工具（若 key 不存在回傳預設值） ===
Private Function NzDict(ByVal d As Object, ByVal key As String, ByVal defaultValue As String) As String
    On Error GoTo EH
    If Not d Is Nothing Then
        If d.Exists(key) Then
            NzDict = CStr(d(key))
            Exit Function
        End If
    End If
EH:
    NzDict = defaultValue
End Function

' ==============================
' === INTEGRATION WORKFLOWS ===
' ==============================
' 本區塊只新增功能，不異動你既有 Sub/Function。
' 目的：
' 1) 讓使用者各別挑選「Excel檔（含 AC5100B 分頁）」與「.txt 原始檔」。
' 2) 用既有 ParseTxt_ArrayVersion 解析 .txt → Collection。
' 3) 將結果以 ApplyCollectionToSheet 寫入所選 Excel 的 "AC5100B" 分頁（覆寫該分頁資料區域）。
' 4) 回到本工作簿（ThisWorkbook），把「整理後資料_交易編號」（若不存在則找「整理後資料_交易單號」）A欄的交易單號，
'    與 AC5100B 的第 C 欄比對，依 B 欄的文字（是否含「評價調整」/「工具息」）把 A/B/J 欄寫回 E~M 指定欄位。

' === 對外主流程：一鍵執行 ===
Public Sub Run_Update_交易編號_From_AC5100B()
    Dim excelPath As String, txtPath As String
    Dim wbAC As Workbook, wsAC As Worksheet
    Dim col As Collection

    ' 1) 選檔：Excel（目標含 AC5100B 分頁）
    excelPath = PickFile("請選擇要更新的 Excel（需包含 AC5100B 分頁）", _
                         "Excel 檔案|*.xlsx;*.xlsm;*.xlsb;*.xls")
    If excelPath = "" Then Exit Sub

    ' 2) 選檔：TXT（要解析的原始檔）
    txtPath = PickFile("請選擇要解析的 TXT 檔", _
                       "文字檔|*.txt|所有檔案|*.*")
    If txtPath = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call 整理外幣債_依原始標題自動拆欄   ' 先執行 1 的拆欄流程（不改動原本 Sub 內容）

    ' 3) 開啟目標 Excel 並定位 AC5100B 分頁
    Set wbAC = Workbooks.Open(Filename:=excelPath, ReadOnly:=False)
    On Error Resume Next
    Set wsAC = wbAC.Sheets("Sheet1")
    On Error GoTo 0
    If wsAC Is Nothing Then
        MsgBox "在選擇的 Excel 中找不到分頁：AC5100B", vbExclamation
        GoTo CleanExit
    End If

    ' 4) 解析 TXT → Collection（沿用你既有的邏輯）
    Set col = ParseTxt_ArrayVersion(txtPath)

    ' 5) 將結果寫入目標 Excel 的 AC5100B 分頁（直接覆蓋，邏輯交給 ApplyCollectionToSheet）
    ApplyCollectionToSheet col, wsAC

    ' 6) 建立 AC5100B 的索引（以 C 欄為 key → Row）
    Dim acIndex As Object
    Set acIndex = BuildIndexByColumn(wsAC, 3) ' 3 = 欄 C

    ' 7) 反寫到本工作簿：直接指定工作表「整理後資料_交易編號」
    Dim wsCases As Worksheet
    Set wsCases = ThisWorkbook.Sheets("整理後資料_交易單號")

    ' === USERFORM ROW ALLOC START ===
    ' 讓使用者以 UserForm（或後援 InputBox）設定各類別的最少行數，
    ' 並依據 D 欄類別（FVPL / FVOCI / AC）在 wsCases 中重排/預留空白列。
    Dim nFVPL As Long, nFVOCI As Long, nAC As Long
    Dim didPick As Boolean

    ' 讀取/顯示輸入框（UserForm 版本見本模組末尾說明）
    didPick = PromptRowAlloc(nFVPL, nFVOCI, nAC)
    If didPick Then
        Call EnsureRowAllocation(wsCases, nFVPL, nFVOCI, nAC)
    End If
    ' === USERFORM ROW ALLOC END ===

    Dim lastRow As Long, r As Long
    Dim key As String, acRow As Long

    ' === REFRESH LASTROW AFTER ROW-ALLOC ===
    lastRow = wsCases.Cells(wsCases.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        key = Trim$(CStr(wsCases.Cells(r, "A").Value)) ' A = 交易單號
        If key <> "" Then
            If acIndex.Exists(key) Then
                Dim rowsStr As String, parts As Variant, pi As Long
                rowsStr = CStr(acIndex(key))
                parts = Split(rowsStr, ",")
                For pi = LBound(parts) To UBound(parts)
                    acRow = Val(Trim$(parts(pi)))
                    If acRow > 0 Then
                        Call WriteBack_ByRule(wsCases, r, wsAC, acRow)
                    End If
                Next pi
            End If
        End If
    Next r

    MsgBox "完成：已更新 AC5100B 並回寫至「整理後資料_交易編號/單號」！", vbInformation

CleanExit:
    On Error Resume Next
    If Not wbAC Is Nothing Then wbAC.Close SaveChanges:=True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' === 檔案挑選小工具（不依賴參照，僅使用內建 FileDialog） ===
Private Function PickFile(ByVal titleText As String, ByVal filterText As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = titleText
        .Filters.Clear
        Call AddFilters(.Filters, filterText)
        If .Show = -1 Then
            PickFile = .SelectedItems(1)
        Else
            PickFile = ""
        End If
    End With
End Function

' 將 "描述|*.ext;*.ext2|描述2|*.x" 的格式拆解加入濾器
Private Sub AddFilters(ByVal filters As FileDialogFilters, ByVal spec As String)
    Dim parts() As String, i As Long
    parts = Split(spec, "|")
    For i = LBound(parts) To UBound(parts) Step 2
        If i + 1 <= UBound(parts) Then
            filters.Add parts(i), parts(i + 1)
        End If
    Next i
End Sub

' 以指定欄建立字典：key → 以逗號分隔的多筆 row（保留全部可能匹配）
Private Function BuildIndexByColumn(ByVal ws As Worksheet, ByVal colIndex As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long, key As String
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    For r = 2 To lastRow
        key = Trim$(CStr(ws.Cells(r, colIndex).Value))
        If key <> "" Then
            If d.Exists(key) Then
                d(key) = d(key) & "," & CStr(r)
            Else
                d.Add key, CStr(r)
            End If
        End If
    Next r
    Set BuildIndexByColumn = d
End Function

' === NEW: 依標題找欄位（逐次執行都會掃描 Row 1；找不到就拋錯）===
Private Function ColByHeader(ByVal ws As Worksheet, ByVal headerName As String, Optional ByVal headerRow As Long = 1) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).Value)), headerName, vbTextCompare) = 0 Then
            ColByHeader = c
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 513, , "找不到欄位：「" & headerName & "」。請確認工作表 """ & ws.Name & """ 第 " & headerRow & " 列的標題是否存在且拼字一致。"
End Function

' 取第一個索引行號（若有多筆相同 key，取第一筆；若有需求可改成全部寫回）

' 規則寫回：
' 若 AC!B 不含 「評價調整」「工具息」→ 成本：E=J, F=A, G=B
' 若 AC!B 含 「評價調整」→ I=J, J=A, H=B
' 若 AC!B 含 「工具息」→ K=J, L=A, M=B
Private Sub WriteBack_ByRule(ByVal wsCases As Worksheet, ByVal rowCases As Long, _
                             ByVal wsAC As Worksheet, ByVal rowAC As Long)
    Dim acctA As Variant, nameB As Variant, amtJ As Variant
    Dim catB As String

    acctA = wsAC.Cells(rowAC, 1).Value   ' A
    nameB = wsAC.Cells(rowAC, 2).Value   ' B
    catB = Trim$(CStr(nameB))            ' 用於判斷關鍵字（實際仍寫入 B）
    amtJ = wsAC.Cells(rowAC, 9).Value   ' J = 10

    ' --- 每次呼叫都「即時」定位九個欄位（用標題找） ---
    Dim 成本_原幣_Col As Long, 成本科目代號_Col As Long, 成本科目名稱_Col As Long
    Dim 評價調整_美元_Col As Long, 評價調整科目代號_Col As Long, 評價調整科目名稱_Col As Long
    Dim 利息_美元_Col As Long, 利息科目代號_Col As Long, 利息科目名稱_Col As Long

    成本_原幣_Col = ColByHeader(wsCases, "成本_原幣")
    成本科目代號_Col = ColByHeader(wsCases, "成本科目代號")
    成本科目名稱_Col = ColByHeader(wsCases, "成本科目名稱")

    評價調整_美元_Col = ColByHeader(wsCases, "評價調整_美元")
    評價調整科目代號_Col = ColByHeader(wsCases, "評價調整科目代號")
    評價調整科目名稱_Col = ColByHeader(wsCases, "評價調整科目名稱")

    利息_美元_Col = ColByHeader(wsCases, "利息_美元")
    利息科目代號_Col = ColByHeader(wsCases, "利息科目代號")
    利息科目名稱_Col = ColByHeader(wsCases, "利息科目名稱")

    If InStr(1, catB, "債務工具評價調整", vbTextCompare) > 0 Or InStr(1, catB, "金融資產評價調整", vbTextCompare) > 0 Then
        wsCases.Cells(rowCases, 評價調整_美元_Col).Value = amtJ
        wsCases.Cells(rowCases, 評價調整科目代號_Col).Value = acctA
        wsCases.Cells(rowCases, 評價調整科目名稱_Col).Value = nameB
    ElseIf InStr(1, catB, "債務工具息", vbTextCompare) > 0 Or InStr(1, catB, "金融資產息", vbTextCompare) > 0 Then
        wsCases.Cells(rowCases, 利息_美元_Col).Value = amtJ
        wsCases.Cells(rowCases, 利息科目代號_Col).Value = acctA
        wsCases.Cells(rowCases, 利息科目名稱_Col).Value = nameB
    ElseIf (InStr(1, catB, "評價損益", vbTextCompare) = 0) And (InStr(1, catB, "備抵損失", vbTextCompare) = 0) Then
    ' 當字串中「不包含」這兩個關鍵字時才執行
        wsCases.Cells(rowCases, 成本_原幣_Col).Value = amtJ
        wsCases.Cells(rowCases, 成本科目代號_Col).Value = acctA
        wsCases.Cells(rowCases, 成本科目名稱_Col).Value = nameB
    End If
End Sub






' ==============================================================
' === USERFORM 支援：讀寫 TagName 與重排 wsCases（最小更動） ===
' ==============================================================

Private Const DEF_FVPL As Long = 10
Private Const DEF_FVOCI As Long = 30
Private Const DEF_AC As Long = 20

' 讀取命名範圍（若不存在則在 wsCases 建立並寫入預設值）
Private Function GetTagCount(ByVal wsCases As Worksheet, ByVal tagName As String, ByVal defaultVal As Long) As Long
    Dim nm As Name
    On Error Resume Next
    Set nm = ThisWorkbook.Names(tagName)
    On Error GoTo 0
    If nm Is Nothing Then
        ' 在 wsCases 的 P1~P3 依序建立，不碰到 A:D 與 H:O 之間含公式欄位
        Dim tgt As Range
        Dim addr As String
        Select Case tagName
            Case "FVPL_計數": Set tgt = wsCases.Range("P1")
            Case "FVOCI_計數": Set tgt = wsCases.Range("P2")
            Case "AC_計數": Set tgt = wsCases.Range("P3")
            Case Else: Set tgt = wsCases.Range("P10") ' 後援
        End Select
        tgt.Value = defaultVal
        addr = "=" & wsCases.Name & "!" & tgt.Address(False, False)
        ThisWorkbook.Names.Add Name:=tagName, RefersTo:=addr
        GetTagCount = defaultVal
    Else
        On Error Resume Next
        GetTagCount = CLng(Val(nm.RefersToRange.Value))
        If Err.Number <> 0 Then GetTagCount = defaultVal
        On Error GoTo 0
    End If
End Function

Private Sub SetTagCount(ByVal tagName As String, ByVal newVal As Long)
    Dim nm As Name
    On Error Resume Next
    Set nm = ThisWorkbook.Names(tagName)
    If Not nm Is Nothing Then
        nm.RefersToRange.Value = newVal
    End If
    On Error GoTo 0
End Sub

' 顯示 UserForm（若表單不存在，改用 InputBox 後援）。
' 回傳 True 表示使用者按下確定；False 表示取消。
Private Function PromptRowAlloc(ByRef nFVPL As Long, ByRef nFVOCI As Long, ByRef nAC As Long) As Boolean
    Dim wsCases As Worksheet
    Set wsCases = ThisWorkbook.Sheets("整理後資料_交易單號")
    
    ' 預設值先取 TagName / 無則設為 DEF_*
    nFVPL = GetTagCount(wsCases, "FVPL_計數", DEF_FVPL)
    nFVOCI = GetTagCount(wsCases, "FVOCI_計數", DEF_FVOCI)
    nAC = GetTagCount(wsCases, "AC_計數", DEF_AC)
    
    On Error GoTo UseInputBox
    ' 若已匯入 UserForm: frmRowAlloc
    With frmRowAlloc
        .txtFVPL.Value = CStr(nFVPL)
        .txtFVOCI.Value = CStr(nFVOCI)
        .txtAC.Value = CStr(nAC)
        .Show vbModal
        If .Tag = "OK" Then
            nFVPL = CLng(Val(.txtFVPL.Value))
            nFVOCI = CLng(Val(.txtFVOCI.Value))
            nAC = CLng(Val(.txtAC.Value))
            ' 寫回 TagName
            SetTagCount "FVPL_計數", nFVPL
            SetTagCount "FVOCI_計數", nFVOCI
            SetTagCount "AC_計數", nAC
            PromptRowAlloc = True
        Else
            PromptRowAlloc = False
        End If
    End With
    Exit Function

UseInputBox:
    ' 後援：三個 InputBox（避免尚未匯入表單時無法執行）
    nFVPL = CLng(Val(InputBox("FVPL 至少行數", "行數設定", CStr(nFVPL))))
    nFVOCI = CLng(Val(InputBox("FVOCI 至少行數", "行數設定", CStr(nFVOCI))))
    nAC = CLng(Val(InputBox("AC 至少行數", "行數設定", CStr(nAC))))
    SetTagCount "FVPL_計數", nFVPL
    SetTagCount "FVOCI_計數", nFVOCI
    SetTagCount "AC_計數", nAC
    PromptRowAlloc = True
End Function

' 依 D 欄分類（包含 "FVPL"/"FVOCI"/"AC" 字樣），將資料重排為：
'  FVPL 區塊（至少 nFVPL 列，不足補空白）→ FVOCI 區塊（至少 nFVOCI）→ AC 區塊（至少 nAC）→ 其他
' 僅寫入允許的欄位 A:D 與 H:O，避免中間穿插含公式的欄位被覆寫。
Private Sub EnsureRowAllocation(ByVal ws As Worksheet, ByVal nFVPL As Long, ByVal nFVOCI As Long, ByVal nAC As Long)
    Dim lastRow As Long, r As Long
    Dim keyD As String
    Dim fvplRows As Collection, fvociRows As Collection, acRows As Collection, otherRows As Collection
    Dim outRow As Long
    Dim overMsg As String
    
    Set fvplRows = New Collection
    Set fvociRows = New Collection
    Set acRows = New Collection
    Set otherRows = New Collection
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' 分桶
    For r = 2 To lastRow
        keyD = UCase$(Trim$(CStr(ws.Cells(r, "D").Value)))
        If keyD Like "*FVPL*" Then
            fvplRows.Add r
        ElseIf keyD Like "*FVOCI*" Then
            fvociRows.Add r
        ElseIf keyD Like "*AC*" Then
            acRows.Add r
        Else
            otherRows.Add r
        End If
    Next r
    
    ' 目標輸出從第 2 列開始
    outRow = 2
    
    ' 按順序輸出：FVPL → 補空白 → FVOCI → 補空白 → AC → 補空白 → 其他
    outRow = WriteBucket(ws, fvplRows, outRow)
    outRow = PadBlank(ws, outRow, nFVPL, fvplRows.Count)
    If fvplRows.Count > nFVPL Then overMsg = overMsg & "FVPL 超過設定行數（" & fvplRows.Count & "/" & nFVPL & ")\n"
    
    outRow = WriteBucket(ws, fvociRows, outRow)
    outRow = PadBlank(ws, outRow, nFVOCI, fvociRows.Count)
    If fvociRows.Count > nFVOCI Then overMsg = overMsg & "FVOCI 超過設定行數（" & fvociRows.Count & "/" & nFVOCI & ")\n"
    
    outRow = WriteBucket(ws, acRows, outRow)
    outRow = PadBlank(ws, outRow, nAC, acRows.Count)
    If acRows.Count > nAC Then overMsg = overMsg & "AC 超過設定行數（" & acRows.Count & "/" & nAC & ")\n"
    
    outRow = WriteBucket(ws, otherRows, outRow)
    
    ' 清掉尾巴殘留（僅允許欄位）
    If outRow <= lastRow Then
        ClearRangeAllowed ws, outRow, lastRow
    End If
    
    If Len(overMsg) > 0 Then
        MsgBox "注意：以下類別的實際筆數超過預留行數，可能導致後續資料缺失：\n\n" & overMsg, vbExclamation, "行數不足提醒"
    End If
End Sub

' 將某桶的列依序寫到 ws 的 outRow 起，僅複製 A:D 與 H:O 欄位的值。
Private Function WriteBucket(ByVal ws As Worksheet, ByVal rowsCol As Collection, ByVal outRow As Long) As Long
    Dim i As Long, srcRow As Long
    For i = 1 To rowsCol.Count
        srcRow = CLng(rowsCol(i))
        CopyAllowed ws, srcRow, outRow
        outRow = outRow + 1
    Next i
    WriteBucket = outRow
End Function

' 若實際筆數 < needCount，補空白列（僅清 A:D、H:O）。
Private Function PadBlank(ByVal ws As Worksheet, ByVal outRow As Long, ByVal needCount As Long, ByVal actualCount As Long) As Long
    Dim n As Long
    If actualCount < needCount Then
        For n = 1 To (needCount - actualCount)
            ClearRowAllowed ws, outRow
            outRow = outRow + 1
        Next n
    End If
    PadBlank = outRow
End Function

' 允許欄位複製：A:D 與 H:O
Private Sub CopyAllowed(ByVal ws As Worksheet, ByVal srcRow As Long, ByVal dstRow As Long)
    ws.Range(ws.Cells(dstRow, 1), ws.Cells(dstRow, 4)).Value = ws.Range(ws.Cells(srcRow, 1), ws.Cells(srcRow, 4)).Value
    ws.Range(ws.Cells(dstRow, 8), ws.Cells(dstRow, 15)).Value = ws.Range(ws.Cells(srcRow, 8), ws.Cells(srcRow, 15)).Value
End Sub

' 允許欄位清空：A:D 與 H:O
Private Sub ClearRowAllowed(ByVal ws As Worksheet, ByVal rowIndex As Long)
    ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, 4)).ClearContents
    ws.Range(ws.Cells(rowIndex, 8), ws.Cells(rowIndex, 15)).ClearContents
End Sub

' 清除尾段殘留區（允許欄位）
Private Sub ClearRangeAllowed(ByVal ws As Worksheet, ByVal rowFrom As Long, ByVal rowTo As Long)
    If rowTo < rowFrom Then Exit Sub
    ws.Range(ws.Cells(rowFrom, 1), ws.Cells(rowTo, 4)).ClearContents
    ws.Range(ws.Cells(rowFrom, 8), ws.Cells(rowTo, 15)).ClearContents
End Sub

' === UserForm 原始碼（請在 VBE 新增 UserForm 名稱：frmRowAlloc，放置 3 個 TextBox 與 2 個 CommandButton）===
' 1) 控制項名稱：
'    TextBox：txtFVPL, txtFVOCI, txtAC
'    CommandButton：btnOK（Caption: 確定）、btnCancel（Caption: 取消）
' 2) 於 frmRowAlloc 程式碼窗貼上下列程式：
'
' Option Explicit
' Private Sub btnOK_Click()
'     ' 簡單驗證數字
'     If Not IsNumeric(Me.txtFVPL.Value) Or Not IsNumeric(Me.txtFVOCI.Value) Or Not IsNumeric(Me.txtAC.Value) Then
'         MsgBox "請輸入數字。", vbExclamation
'         Exit Sub
'     End If
'     Me.Tag = "OK"
'     Me.Hide
' End Sub
'
' Private Sub btnCancel_Click()
'     Me.Tag = "CANCEL"
'     Me.Hide
' End Sub
'
' ' （可選）在 Initialize 中預設值會由呼叫端填入，不需額外邏輯。
' ' Private Sub UserForm_Initialize()
' ' End Sub

' ==============================================================