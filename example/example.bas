Attribute VB_Name = "��дlisp����"
'���ļ��е����Ӹı��Կ�ѧ�����硶AutoCAD��ȫӦ��ָ�ϡ���2011��4�µ�һ�棩�е�ʾ��
'���е�ʾ��ΪAutoLISP���򣬱��ļ���Ϊ��VBAʵ��

'��21ҳ��pbox.lsp
'��һ��������塱��״��ͼ��
Sub c_gfb()
    vba_start
    p = uty.GetPoint(, "���½ǵ㣺")
    w = uty.GetReal("���:")
    h = uty.GetReal("�߶�:")
    
    p���� = rc(p, w, h)
    AddRec2d p, p����
    p���� = rc(p, w / 2, 0)
    p���� = rc(p, 0, h / 2)
    msp.AddLine p����, rc(p����, 0, h)
    msp.AddLine p����, rc(p����, w, 0)
    
End Sub
'��55ҳ��7test0.lsp
'����ͼ��
Sub c_hztk()
    vba_start
    uty.InitializeUserInput 0, "A0 A1 A2 A3 A4" '��һλΪ1,�����������ֵ
    size_ = uty.GetKeyword("ͼֽ��СA0,A1,A2,A3,A4��<A3>")
    If size_ = "" Then size_ = "A3"
    size_ = UCase(size_) '��Ȼinitial����ΪA0,���ǲ�����ֹ�û�����a0
    p1 = uty.GetPoint(, "���½ǵ㣺")
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
'��45ҳ��MCIR.LSP
'�е㻭Բ
Sub c_zdhy()
    vba_start
    Set obj = GetEntity("ѡһ����")(0)
    pm = GetMid(obj.StartPoint, obj.EndPoint)
    r = uty.GetDistance(pm, "�뾶��")
    msp.AddCircle pm, r
End Sub
'��81ҳ��9test1.lsp
'���°뾶
Sub c_rad_renew()
    vba_start
    uty.prompt vbCr & ">>>ѡ��"
    Dim ss As AcadSelectionSet
    Set ss = dwg.SelectionSets.Add(0)
    ss.SelectOnScreen
    rad_new = uty.GetReal("�°뾶��")
    For Each o In ss
        If o.ObjectName = "AcDbCircle" Then o.Radius = rad_new
    Next
End Sub
'��90ҳ��10test1.lsp
'���ļ�
Sub c_read_file()
    vba_start
    f_path = GetFile("ѡ��Ҫ������ļ�", "", "txt", ReadFile)
    
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
'��93ҳ��10test3.lsp
'д�ļ�
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
    
    f_path = GetFile("ѡ�񱣴�λ��", "", "txt", WriteOrAppendFile)
    
    Set f = oFileOpen(f_path, ForWriting)
    f.WriteLine "�������  ����"
    f.WriteLine "=============="
    f.WriteLine "Բ��������" & n_circle
    f.WriteLine "ֱ�ߵ�������" & n_line
    f.WriteLine "���ֵ�������" & n_text
    
    f.Close
    
End Sub




