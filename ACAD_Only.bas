Attribute VB_Name = "ACAD����"
Public dwg As AcadDocument
Public msp As AcadModelSpace
Public uty As AcadUtility
Public Enum GetFileMode
    ReadFile = 2
    WriteOrAppendFile = 1
End Enum
Public Sub vba_start()
    Set dwg = ThisDrawing
    Set msp = dwg.ModelSpace
    Set uty = dwg.Utility
    ss_count = dwg.SelectionSets.Count
    For i = ss_count - 1 To 0 Step -1
        dwg.SelectionSets(i).Delete
    Next
End Sub
'VBA����->CAD����
Sub defun()
    ThisDrawing.SendCommand "REDRAW" & vbCr
    For Each pr In Application.VBE.VBProjects
        For Each Comp In pr.VBComponents
            Set codemod = Comp.CodeModule
            ct = codemod.CountOfLines
            i = 1
            While i <= ct
                n = codemod.ProcOfLine(i, 0)
                If n <> "" Then
                    i = i + codemod.ProcCountLines(n, 0)
                    If Left(n, 2) = "c_" Then
                        ThisDrawing.SendCommand "(defun c:" & Mid(n, 3) & " () (command  ""-VBARUN"" """ & n & """))" & vbCr
                    End If
                Else
                    i = i + 1
                End If
            Wend
        Next
    Next
End Sub
'���������ߵ��е�
Public Function GetMid(p1, p2) As Double()
    Dim temp(2) As Double
    temp(0) = (p1(0) + p2(0)) / 2
    temp(1) = (p1(1) + p2(1)) / 2
    temp(2) = (p1(2) + p2(2)) / 2
    GetMid = temp
End Function
'��ֱ�����귽ʽȡ����һ��
Public Function rc(p, dx, dy, Optional dz = 0) As Double()
    Dim temp(2) As Double
    temp(0) = p(0) + dx
    temp(1) = p(1) + dy
    temp(2) = p(2) + dz
    rc = temp
End Function
'�����ת��Ϊ�ַ�����,��commandʹ��
Public Function Pstr(p) As String
    Pstr = CStr(p(0)) & "," & CStr(p(1)) & "," & CStr(p(2))
End Function
'���ɾ���
Public Function AddRec2d(p1, p2) As AcadLWPolyline
    pRD = rc(p1, p2(0) - p1(0), 0)
    pLU = rc(p1, 0, p2(1) - p1(1))
    Dim p(7) As Double
    p(0) = p1(0): p(1) = p1(1)
    p(2) = pRD(0): p(3) = pRD(1)
    p(4) = p2(0): p(5) = p2(1)
    p(6) = pLU(0): p(7) = pLU(1)
    Set AddRec2d = ThisDrawing.ModelSpace.AddLightWeightPolyline(p)
    AddRec2d.Closed = True
End Function
'���Ա�ע
Public Function AddDimLinears(dist, ParamArray points()) As Variant()
    Dim out()
    ReDim out(UBound(points) - 1)
    
    If points(0)(1) = points(1)(1) Then 'ˮƽ��ע
        For i = 0 To UBound(points) - 1
            pm = rc(GetMid(points(i), points(i + 1)), 0, dist)
            Set out(i) = ThisDrawing.ModelSpace.AddDimRotated(points(i), points(i + 1), pm, 0)
        Next
    Else '��ֱ��ע
        For i = 0 To UBound(points) - 1
            pm = rc(GetMid(points(i), points(i + 1)), dist, 0)
            Set out(i) = ThisDrawing.ModelSpace.AddDimRotated(points(i), points(i + 1), pm, pi / 2)
        Next
    End If
    
    AddDimLinears = out
End Function
'������SendCommand����(�ÿո����vbCr,�ַ�������Զ����һ��vbCr)
Public Sub SendCmd(ParamArray s())
    ThisDrawing.SendCommand Replace(Join(s, vbCr), " ", vbCr) & vbCr
End Sub
'��ȡ��̬����Զ�������
Public Function GetBlockAttribute(obj, name)
    attrs = obj.GetAttributes()
    For Each a In attrs
        If a.TagString = name Then GetBlockAttribute = a.TextString: Exit For
    Next
End Function
'���ö�̬����Զ�������
Public Sub SetBlockAttribute(obj, name, val)
    attrs = obj.GetAttributes()
    For Each a In attrs
        If a.TagString = name Then a.TextString = val: Exit For
    Next
End Sub
'��ȡ��̬����Զ������
Public Function GetBlockProperty(obj, name)
    props = obj.GetDynamicBlockProperties()
    For Each p In props
        If p.PropertyName = name Then GetBlockProperty = p.Value: Exit For
    Next
End Function
'���ö�̬����Զ������
Public Sub SetBlockProperty(obj, name, val)
    props = obj.GetDynamicBlockProperties()
    For Each p In props
        If p.PropertyName = name Then p.Value = val: Exit For
    Next
End Sub
'����ȫ��GetEntity����
Public Function GetEntity(Optional prompt = "") As Variant()
    Dim obj As AcadObject
    On Error Resume Next
    ThisDrawing.Utility.GetEntity obj, click_point, prompt
    Do While obj Is Nothing
        ThisDrawing.Utility.InitializeUserInput 1, "Y N"
        op = ThisDrawing.Utility.GetKeyword("δѡ���κζ���,�Ƿ����ѡ��[Y/N]<Y>")
        If op = "Y" Then ThisDrawing.Utility.GetEntity obj, click_point, prompt Else End
        '�˳�����������End������End Function
    Loop
    GetEntity = Array(obj, click_point)
End Function
'����ѡ���ļ��Ի���
Public Function GetFile(msg, default_path, extension, mode As GetFileMode) As String
    msg = """" & msg & """ "
    default_path = """" & default_path & """ "
    extension = """" & extension & """ "
    
    ThisDrawing.SendCommand _
    "(princ (getfiled " & msg & default_path & extension & mode & "))(prin1)" & vbCr
    GetFile = ThisDrawing.GetVariable("LastPrompt")
    If GetFile = "nil" Then End
End Function
'��������ȷ�������͵�����
Public Function TypedArr(first_value, ParamArray values() As Variant)
    Dim UtyObj As Object
    Set UtyObj = ThisDrawing.Utility
    Dim temp
    UtyObj.CreateTypedArray temp, VarType(first_value), first_value
    up = UBound(values)
    ReDim Preserve temp(up + 1)
    For i = 0 To up
        temp(i + 1) = values(i)
    Next
    TypedArray = temp
End Function
'��Collection��ArrayListת��Ϊ����ȷ�������͵�����
Function ToTypedArr(a)
    Dim UtyObj As Object
    Set UtyObj = ThisDrawing.Utility
    If TypeName(a) = "Collection" Then first_elem = a(1) Else first_elem = a(0)
    Dim temp
    UtyObj.CreateTypedArray temp, VarType(first_elem), first_elem
    ReDim temp(a.Count - 1)
    For Each elem In a
        temp(i) = elem
        i = i + 1
    Next
    ToTypedArr = temp
End Function



