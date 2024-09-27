Attribute VB_Name = "Universal"
Public Const pi As Double = 3.14159265358979
Public Enum OpenFileMode
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum
'Collection对象
Public Function oCollection(ParamArray arr()) As Collection
    Set oCollection = New Collection
    For Each i In arr
        oCollection.Add i
    Next
End Function
'Collection转Array
Function CollToArr(coll As Collection)
    ub = coll.Count - 1
    Dim temp()
    If ub = -1 Then CollToArr = temp: Exit Function
    ReDim temp(ub)
    i = 0
    For Each elem In coll
        temp(i) = elem
        i = i + 1
    Next
    CollToArr = temp
End Function
'ArrayList对象（需要在“启用或关闭Windows功能”中勾选“.NET Framework 3.5”）
Public Function oList(ParamArray arr()) As Object
    Set oList = CreateObject("System.Collections.ArrayList")
    For Each i In arr
        oList.Add i
    Next
End Function
'字典对象
Public Function oDict(ParamArray items()) As Object
    Set oDict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(items) - LBound(items) Step 2
        oDict.Add items(i), items(i + 1)
    Next
End Function
'通用的Length函数
Public Function Len1(a) As Long
    If TypeName(a) = "String" Then
        Len1 = Len(a)
    ElseIf Right(TypeName(a), 2) = "()" Then
        Len1 = UBound(a) - LBound(a) + 1
    Else
        Len1 = a.Count
    End If
End Function
'切片(Slice)函数
Public Function Slc(a, st, Optional ed = Null)
    ct = Len1(a)
    If st < 0 Then st = st + ct
    If IsNull(ed) Then
        ed = ct
    ElseIf ed < 0 Then
        ed = ed + ct
    End If
    
    If TypeName(a) = "String" Then
        Slc = Mid(a, st + 1, ed - st)
    ElseIf TypeName(a) = "ArrayList" Then
        Set out = oList
        For i = st To ed - 1
            out.Add a(i)
        Next
        Set Slc = out
    ElseIf TypeName(a) = "Collection" Then
        Set out = oCollection
        For i = st To ed - 1
            out.Add a(i)
        Next
        Set Slc = out
    Else
        lb = LBound(a)
        Dim out1()
        ub1 = -1
        For i = st + lb To ed + lb - 1
            ub1 = ub1 + 1
            ReDim Preserve out1(ub1)
            out1(ub1) = a(i)
        Next
        Slc = out1
    End If
End Function
'字符串格式化
Public Function fmt(ParamArray arr()) As String
    s = arr(0)
    s = Replace(s, "\n", vbLf)
    s = Replace(s, "\r", vbCr)
    s = Replace(s, "\r\n", vbCrLf)
    s = Replace(s, "\t", vbTab)
    
    s = Replace(s, "\" & vbLf, "\n") '\\n
    s = Replace(s, "\" & vbCr, "\r") '\\r
    s = Replace(s, "\" & vbTab, "\t") '\\t
    s = Replace(s, "\\", "\")  '\\
    s = Replace(s, "'", """")
    
    If UBound(arr) > 0 Then
        For i = 1 To UBound(arr)
            s = Replace(s, "{}", CStr(arr(i)), 1, 1)
        Next
    End If
    
    fmt = s
End Function
'用一行代码设定正则表达式
Public Function oRegExp(Pattern1, Optional Global1 = Null, Optional IgnoreCase1 = Null, Optional Multiline1 = Null) As Object
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = Pattern1
    If Not IsNull(Global1) Then oRegExp.Global = Global1
    If Not IsNull(IgnoreCase1) Then oRegExp.IgnoreCase = IgnoreCase1
    If Not IsNull(Multiline1) Then oRegExp.Multiline = Multiline1
End Function
'Excel函数库
Public Function oExcelFunc() As Object
    Set xls = CreateObject("Excel.Application")
    Set oExcelFunc = xls.WorksheetFunction
End Function
'通过shell调用其他程序,并获取返回值
Public Function ShellOut(command, argv) As String
    command = command & " " & Join(argv)
    s_out = ""
    
    Set ws_shell = CreateObject("WScript.Shell")
    Set ws_exec = ws_shell.Exec(command)
    Set ws_out = ws_exec.StdOut
    Do While Not ws_out.AtEndOfStream
        s_out = s_out & ws_out.ReadLine & vbLf
    Loop
    s_out = Left(s_out, Len(s_out) - 1)
    ShellOut = s_out
End Function
'弹出选择文件夹对话框
Public Function GetFolder(msg) As String
    Set shell_app = CreateObject("shell.application")
    Set folder = shell_app.BrowseForFolder(0, msg, 1)
    If folder Is Nothing Then
        End
    Else
        GetFolder = folder.Self.Path
        GetFolder = GetFolder & IIf(Right(GetFolder, 1) = "\", "", "\")
    End If
End Function
'读写文件
Public Function oFileOpen(f_path, mode As OpenFileMode) As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If mode = ForReading Then create = False Else create = True
    Set oFileOpen = fs.OpenTextFile(f_path, mode, create, -2)
End Function
