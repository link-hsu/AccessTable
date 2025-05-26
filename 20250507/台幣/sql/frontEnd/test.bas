

' === 模块：modDataAccess ===
' 通用：从 PositionMap 取出指定报表的所有映射
Public Function GetPositionMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim results() As Variant
    Dim i As Long
    
    ' 1. 建立连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. SQL: 先 INNER JOIN Report 找到 ReportID，再取 PositionMap
    sql = "SELECT pm.TargetSheetName, pm.SourceNameTag, pm.TargetCellAddress " & _
          "FROM PositionMap AS pm " & _
          "INNER JOIN Report AS r ON pm.ReportID = r.ReportID " & _
          "WHERE r.ReportName = '" & reportName & "' " & _
          "ORDER BY pm.DataId;"
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetPositionMapData = Array()  ' 没有记录
    Else
        ' 3. 把结果装进二维数组：每行一笔 mapping
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields("TargetSheetName").Value
            results(i, 1) = rs.Fields("SourceNameTag").Value
            results(i, 2) = rs.Fields("TargetCellAddress").Value
            i = i + 1
            rs.MoveNext
        Loop
        GetPositionMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function


'=== 从 QueryMap 取出指定报表的所有查询配置 ===
Public Function GetQueryMapData( _
        ByVal DBPath As String, _
        ByVal reportName As String _
    ) As Variant
    Dim conn As Object, rs As Object
    Dim sql As String, results() As Variant
    Dim i As Long
    
    ' 1. 建立 ADO 连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath
    
    ' 2. 用 INNER JOIN 取 ReportID，再拿 QueryMap
    sql = "SELECT qm.QueryTableName, " & _
          "qm.ImportColName, qm.ImportColNumber " & _
          "FROM QueryMap AS qm " & _
          "INNER JOIN Report AS r ON qm.ReportID = r.ReportID " & _
          "WHERE r.ReportName='" & reportName & "' " & _
          "ORDER BY qm.DataId;"
    
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        GetQueryMapData = Array()    ' 没有任何配置
    Else
        ' 把结果装二维 Array：(0)=QueryTableName, (1)=ImportColName, (2)=ImportColNumber
        rs.MoveLast: rs.MoveFirst
        ReDim results(0 To rs.RecordCount - 1, 0 To 2)
        i = 0
        Do Until rs.EOF
            results(i, 0) = rs.Fields("QueryTableName").Value
            results(i, 1) = rs.Fields("ImportColName").Value
            results(i, 2) = rs.Fields("ImportColNumber").Value
            i = i + 1
            rs.MoveNext
        Loop
        GetQueryMapData = results
    End If
    
    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function


我看你的這兩個函數還蠻像的，有沒有辦法把這函數併為一個，
簡化為通用函數，
讓
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
和
Public Sub Process_FM11()
共同call一隻function就好，請把完整代碼提供給我，請記得Init要以你給我的最新一個版本為基礎去修改，修改的地方請詳細標示出來




我希望下拉驗證可以單純只能查看特定Report包含哪些NameTag，以及對應的儲存格和分頁．
只有進入UserForm才能進行資料更新的動作，
請將下拉驗證和UserForm各自進行的程序詳細定義和說明清楚，
任何細節都要說明，
並且要把兩者間的互動過程仔細考慮到，
另外再更新的時候不只要更新資料庫的資料，
還要針對所更新的ReportName(即分頁名稱）、SheetName、儲存格名稱 和 對應的NameTag在Excel中進行更新，請詳細說明整個過程和各個細節