Attribute VB_Name = "win32Api"
'expression.GetOpenFilename (FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
'expression����ʾ Application ����ı�����
'����
'����
'����    ����/��ѡ   ��������    ˵��
'FileFilter  ��ѡ    Variant ָ���ļ�ɸѡ�������ַ�����
'FilterIndex ��ѡ    Variant ָ��Ĭ���ļ�ɸѡ�����������ţ���Ŵ� 1 ֱ�� FileFilter ��ָ����ɸѡ����š� ����˲�����ʡ�Ի���ڴ��ڵ�ɸѡ������ʹ�õ��ǵ�һ���ļ�ɸѡ����
'Title   ��ѡ    Variant ָ���Ի���ı��⡣ ����˲�����ʡ�ԣ�����Ϊ"��"��
'ButtonText  ��ѡ    Variant ���� Macintosh��
'MultiSelect ��ѡ    Variant ��Ϊ True������ѡ�����ļ����� ��Ϊ False��������ѡ��һ���ļ����� Ĭ��ֵΪ False��
Public Function GetOpenFileName(�ļ����� As Variant, ��׺���� As Variant, ���� As Variant, ��ѡ As Boolean) As Collection
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
    openfiles = ExcelApp.GetOpenFileName(FileFilter:=(�ļ����� & " (*" & ��׺���� & "),*" & ��׺���� & ", All Files (*.*),*.*"), Title:=����, MultiSelect:=��ѡ)
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
