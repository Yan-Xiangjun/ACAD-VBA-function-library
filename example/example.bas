Attribute VB_Name = "复写lisp例子"
'本文件中的例子改编自科学出版社《AutoCAD完全应用指南》（2011年4月第一版）中的示例
'书中的示例为AutoLISP程序，本文件改为用VBA实现

'第21页，pbox.lsp
'画一个“光伏板”形状的图形
Sub c_gfb()
    vba_start
    p = uty.GetPoint(, "左下角点：")
    w = uty.GetReal("宽度:")
    h = uty.GetReal("高度:")
    
    p右上 = rc(p, w, h)
    AddRec2d p, p右上
    p下中 = rc(p, w / 2, 0)
    p左中 = rc(p, 0, h / 2)
    msp.AddLine p下中, rc(p下中, 0, h)
    msp.AddLine p左中, rc(p左中, w, 0)
    
End Sub
'第55页，7test0.lsp
'绘制图框
Sub c_hztk()
    vba_start
    uty.InitializeUserInput 0, "A0 A1 A2 A3 A4" '第一位为1,则不允许输入空值
    size_ = uty.GetKeyword("图纸大小A0,A1,A2,A3,A4：<A3>")
    If size_ = "" Then size_ = "A3"
    size_ = UCase(size_) '虽然initial设置为A0,但是不能阻止用户输入a0
    p1 = uty.GetPoint(, "左下角点：")
    Select Case size_
    Case "A0"
        p2 = rc(p1, 1189, 841)
    Case "A1"
        p2 = rc(p1, 841, 594)
    Case "A2"
        p2 = rc(p1, 594, 420)
    Case "A3"
        p2 = rc(p1, 420, 297)
    Case "A4"
        p2 = rc(p1, 297, 210)
    End Select
    AddRec2d p1, p2
End Sub
'第45页，MCIR.LSP
'中点画圆
Sub c_zdhy()
    vba_start
    Set obj = GetEntity("选一条线")(0)
    pm = GetMid(obj.StartPoint, obj.EndPoint)
    r = uty.GetDistance(pm, "半径：")
    msp.AddCircle pm, r
End Sub
'第81页，9test1.lsp
'更新半径
Sub c_rad_renew()
    vba_start
    uty.prompt vbCr & ">>>选择："
    Dim ss As AcadSelectionSet
    Set ss = dwg.SelectionSets.Add(0)
    ss.SelectOnScreen
    rad_new = uty.GetReal("新半径：")
    For Each o In ss
        If o.ObjectName = "AcDbCircle" Then o.Radius = rad_new
    Next
End Sub
'第90页，10test1.lsp
'读文件
Sub c_read_file()
    vba_start
    f_path = GetFile("选择要读入的文件", "", "txt", ReadFile)
    
    Set p = oCollection
    Set f = oFileOpen(f_path, ForReading)
    While Not f.AtEndOfStream
        data = f.ReadLine
        temp = Split(data, ",")
        p.Add CDbl(temp(0))
        p.Add CDbl(temp(1))
    Wend
    
    f.Close
    msp.AddLightWeightPolyline ToTypedArr(p)
End Sub
'第93页，10test3.lsp
'写文件
Sub c_write_file()
    vba_start
    Dim ss As AcadSelectionSet
    Set ss = dwg.SelectionSets.Add(0)
    ss.SelectOnScreen
    n_circle = 0: n_line = 0: n_text = 0
    For Each o In ss
        Select Case o.ObjectName
        Case "AcDbCircle"
            n_circle = n_circle + 1
        Case "AcDbLine"
            n_line = n_line + 1
        Case "AcDbText"
            n_text = n_text + 1
        End Select
    Next
    
    f_path = GetFile("选择保存位置", "", "txt", WriteOrAppendFile)
    
    Set f = oFileOpen(f_path, ForWriting)
    f.WriteLine "对象类别  数量"
    f.WriteLine "=============="
    f.WriteLine "圆的数量：" & n_circle
    f.WriteLine "直线的数量：" & n_line
    f.WriteLine "文字的数量：" & n_text
    
    f.Close
    
End Sub




