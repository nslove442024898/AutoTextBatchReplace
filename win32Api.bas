Attribute VB_Name = "win32Api"
'expression.GetOpenFilename (FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
'expression：表示 Application 对象的变量。
'参数
'参数
'名称    必需/可选   数据类型    说明
'FileFilter  可选    Variant 指定文件筛选条件的字符串。
'FilterIndex 可选    Variant 指定默认文件筛选条件的索引号，编号从 1 直到 FileFilter 中指定的筛选器编号。 如果此参数被省略或大于存在的筛选器数，使用的是第一个文件筛选器。
'Title   可选    Variant 指定对话框的标题。 如果此参数被省略，标题为"打开"。
'ButtonText  可选    Variant 仅限 Macintosh。
'MultiSelect 可选    Variant 若为 True，允许选择多个文件名。 若为 False，仅允许选择一个文件名。 默认值为 False。
Public Function GetOpenFileName(文件类型 As Variant, 后缀名称 As Variant, 标题 As Variant, 多选 As Boolean) As Collection
    Dim isNewStartExcel As Boolean, ExcelApp As Object, fns As New Collection, openfiles As Variant
    On Error Resume Next
    '    Set ExcelApp = GetObject(, "Excel.Application")
    '    If Err <> 0 Then
    '        Err.Clear
    Set ExcelApp = CreateObject("Excel.Application")
    If Err <> 0 Then MsgBox "Could not start Excel!", vbExclamation
    isNewStartExcel = True
    '    Else
    '        isNewStartExcel = False
    '    End If
    '    ExcelApp.Visible = True
    '    Set wbkobj = ExcelApp.Workbooks.Add
    '    Set shtObj = wbkobj.Worksheets(1)
    openfiles = ExcelApp.GetOpenFileName(FileFilter:=(文件类型 & " (*" & 后缀名称 & "),*" & 后缀名称 & ", All Files (*.*),*.*"), Title:=标题, MultiSelect:=多选)
    If isNewStartExcel Then
        ExcelApp.Quit
    End If
    If openfiles <> False Then
        Dim i As Long
        For i = LBound(openfiles) To UBound(openfiles)
            fns.Add openfiles(i)
        Next
        Set GetOpenFileName = fns
    Else
        Set GetOpenFileName = Nothing
    End If
End Function


Public Function FileInUse(sFileName) As Boolean
    On Error Resume Next
    Open sFileName For Binary Access Read Lock Read As #1
    Close #1
    FileInUse = IIf(Err.Number > 0, True, False)
    On Error GoTo 0
End Function
