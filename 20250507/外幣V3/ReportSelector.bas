Option Explicit

Private Sub UserForm_Initialize()
    Dim i As Long
    For i = LBound(allReportNames) To UBound(allReportNames)
        Me.lstReports.AddItem allReportNames(i)
    Next i
End Sub

Private Sub btnOK_Click()
    Dim rptSelect As Collection
    Dim i As Long
    Set rptSelect = New Collection
    For i = 0 To Me.lstReports.ListCount - 1
        If Me.lstReports.Selected(i) Then
            rptSelect.Add Me.lstReports.List(i)
        End If
    Next i
    If rptSelect.Count = 0 Then
        MsgBox "請至少選擇一個報表", vbExclamation
        Exit Sub
    End If
    ReDim gReportNames(0 To rptSelect.Count - 1)
    For i = 1 To rptSelect.Count
        gReportNames(i - 1) = rptSelect(i)
    Next i
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

    
' ==================================
' UserForm代碼

' 2. **放入控制項**

'    * **ListBox**

'      * 名稱：`lstReports`
'      * MultiSelect 屬性：`fmMultiSelectMulti` （可多選）
'      * 大小自行調整，以顯示所有報表名稱
'    * **CommandButton** ×2

'      1. 名稱：`cmdOK`；Caption：`確定`
'      2. 名稱：`cmdCancel`；Caption：`取消`

' 3. **UserForm Code（按兩下 UserForm 背景貼上）**

'    ```vb
'    Option Explicit

'    ' 傳回使用者勾選結果用的全域變數 (預先在 Module 宣告)
'    '   Public gReportNames As Variant

'    Private Sub UserForm_Initialize()
'        Dim i As Long
'        ' allReportNames 需為 Module 層級可見
'        For i = LBound(allReportNames) To UBound(allReportNames)
'            Me.lstReports.AddItem allReportNames(i)
'        Next i
'    End Sub

'    Private Sub cmdOK_Click()
'        Dim sel As Collection, i As Long
'        Set sel = New Collection
'        ' 收集被選取的項目
'        For i = 0 To Me.lstReports.ListCount - 1
'            If Me.lstReports.Selected(i) Then
'                sel.Add Me.lstReports.List(i)
'            End If
'        Next i
'        If sel.Count = 0 Then
'            MsgBox "請至少選擇一個報表", vbExclamation
'            Exit Sub
'        End If
'        ' 將 Collection 轉回陣列給 Main 用
'        ReDim gReportNames(0 To sel.Count - 1)
'        For i = 1 To sel.Count
'            gReportNames(i - 1) = sel(i)
'        Next i
'        Me.Hide
'    End Sub

'    Private Sub cmdCancel_Click()
'        ' 不做任何事就關閉，Main 看到 gReportNames 為空會中止
'        Me.Hide
'    End Sub

' 前情提要
' =====================================
