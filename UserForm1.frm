VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�����ı��滻"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dictfls As Object
Public directory As String
Const tempSelectFileListName As String = "TemplateSelectFileList.csv"

'Public curReplaceFileIndex As Long
Private Sub btnAddFile_Click()
    Dim fl As Variant, i As Long, fls As Collection
    '���ļ�
    '    https://forums.autodesk.com/t5/inventor-customization/folder-browser-needed-for-vba-7-64-bit/m-p/4365989#M45667
    '    https://forums.augi.com/showthread.php?153961-VBA-Autocad-Office-64-bit-Open-File-dialog-box-using-comdlg32-dll
    '    Dim ofd As New clsOpenSaveFileName
    '    ' Open the Save as form through the windows api and allow the designer a chance to
    '    ' select the save location and filename
    '    Dim FileName As String
    '    With ofd
    '        .Filter = "AutoCAD Drawing FILES (*.dwg)" & Chr$(0) & "*.dwg"
    '        .StartInDir = CurDir()
    '        .Title = "Please select file to open"
    '        .MultiSelect = True
    '        .SelectedFilter = 1
    '        .OwnerHwnd = Application.HWND
    '    End With
    '
    '    FileName = ofd.ShowOpen()
    '    ' Check to make sure user hasn't selected cancel or something else which will cause the
    '    ' filename to report blank
    '    If Len(FileName) < 1 Then
    '        MsgBox ("Something has gone wrong, please try again.")
    '        Exit Sub
    '    End If
    
    Set fls = GetOpenFileName("AutoCAD", ".dwg", "", True)
    VBA.AppActivate Application.Caption
    If TypeName(fls) <> "Nothing" Then
        For i = 1 To fls.count
            If Not dictfls.Exists(fls.item(i)) Then
                If Not win32Api.FileInUse(fls.item(i)) Then
                    dictfls.Add fls.item(i), 0
                    Me.ListBox1.AddItem fls.item(i)
                Else
                    Me.Label3 = "�ļ�  " & fls.item(i) & " ����ֻ��״̬,�޷���ӵ������ļ��б�"
                End If
            Else
                Me.Label3 = "�ļ� " & fls.item(i) & " �Ѿ����,�����ظ����"
            End If
        Next
    Else
        MsgBox "ûѡ���κ��ļ�", vbCritical + vbOKOnly
    End If
    
End Sub

Private Sub btnBatchReplace_Click()
    
    Dim i As Long, var As Variant, fl As String, icount As Long, layerctl As String, txtheight As Double 'aligment As AutoCAD.AcAlignment
    If Me.CheckBox2Layer Then layerctl = Me.txtBoxLayer Else layerctl = vbNullString
    If Me.CheckBox4TxtHeight Then txtheight = CDbl(Me.txtBoxHeight) Else txtheight = 0
    'If Me.CheckBox3Aligment Then aligment = CInt(Me.TextBox2) Else aligment = 0
    If Me.ListBox1.ListCount > 0 Then
        
        If Me.txtOldStr.Text = vbNullString Or Me.TxtNewStr.Text = vbNullString Then
            Me.Label3 = "��������Ҫ�滻�������ٵ�������滻��ť"
            Exit Sub
        End If
        
        For i = 0 To Me.ListBox1.ListCount - 1
            'Me.curReplaceFileIndex = i
            DoEvents
            fl = Me.ListBox1.List(i)
            If Me.staticReplace = True Then
                icount = ReplaceTextModule.�滻�ļ��ڲ�����������(fl, Me.txtOldStr.Text, Me.TxtNewStr.Text, False, layerctl, txtheight, Me.CheckBox1��ȫƥ��.Value, _
                    Me.CheckBox5�滻��������.Value) ', Me.CheckBox6�滻���Կ�������.Value)
            Else
                icount = ReplaceTextModule.�滻�ļ��ڲ�����������(fl, Me.txtOldStr.Text, Me.TxtNewStr.Text, True, layerctl, txtheight, Me.CheckBox1��ȫƥ��.Value, _
                    Me.CheckBox5�滻��������.Value) ', Me.CheckBox6�滻���Կ�������.Value)
            End If
            If icount > 0 Then
                Me.Label3 = "�����滻 " & fl & ",�� " & Me.ListBox1.ListCount & " ��ͼֽ��Ҫ�滻,�����滻�� " & i + 1 & " ��ͼֽ,���滻 " & icount & " ������"
            Else
                Me.Label3 = "��ǰͼֽ " & fl & "�ڲ��Ҳ�����Ҫ�滻������"
            End If
        Next
        Me.Label3 = "ȫ���滻���!"
    Else
        MsgBox "δѡ���κ�ͼֽ,�޷������滻,��������ļ�!", vbExclamation
    End If
    
End Sub

'ȡ�����е���
Private Sub btnCancelAll_Click()
    Dim index As Long
    For index = Me.ListBox1.ListCount - 1 To 0 Step -1
        Me.dictfls.Remove Me.ListBox1.List(index)
        Me.ListBox1.RemoveItem index
    Next
End Sub

'ɾ��ѡ�����
Private Sub btnCancelSelect_Click()
    If Me.ListBox1.ListIndex <> -1 Then
        Dim index As Long
        For index = Me.ListBox1.ListCount - 1 To 0 Step -1
            If Me.ListBox1.Selected(index) = True Then
                Me.dictfls.Remove Me.ListBox1.List(index)
                Me.ListBox1.RemoveItem index
            End If
        Next
    End If
End Sub

'��ȡ�ļ��б�
Private Sub btnReadFileList_Click()
    Dim Arr() As String, i As Long, fso As Object, dataFn As String, sr As Object
    dataFn = Me.directory + "\" + tempSelectFileListName
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(dataFn) Then
        Set sr = fso.OpenTextFile(dataFn, 1, False)
        Do While sr.AtEndOfLine = False
            ReDim Preserve Arr(0 To i)
            Arr(i) = sr.ReadLine()
            i = i + 1
        Loop
        sr.Close
        Set sr = Nothing
        '
        Me.ListBox1.Clear
        Me.dictfls.RemoveAll
        For i = LBound(Arr) To UBound(Arr)
            If fso.FileExists(Arr(i)) Then
                If win32Api.FileInUse(Arr(i)) Then
                    Me.Label3 = "�ļ��б��� " & Chr(34) & Arr(i) & Chr(34) & " �ļ���ȡʧ��,��Ϊ�ļ���ռ��,�޷������ļ�"
                Else
                    Me.dictfls.Add Arr(i), 0
                    Me.ListBox1.AddItem Arr(i)
                End If
            Else
                Me.Label3 = "�ļ��б��� " & Chr(34) & Arr(i) & Chr(34) & " �ļ���ȡʧ��,��Ϊ�ļ��Ҳ���!"
            End If
        Next
        Me.Label3 = "�ɹ���ȡ�ļ��б�"
    Else
        Me.Label3 = "δ�ҵ�������б�"
    End If
    Set fso = Nothing
    
End Sub

Private Sub btnsaveList_Click()
    Dim fso As Object, var As Variant, ArrHelper As New BetterArray, col As New Collection, resArr As BetterArray
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(Me.directory + "\" + tempSelectFileListName) Then
        For Each var In Me.dictfls.Keys()
            col.Add var
        Next
        Set resArr = ArrHelper.CopyFromCollection(col)
        resArr.ToCSVFile Me.directory + "\" + tempSelectFileListName
        MsgBox "�ɹ������ļ��б�"
        Set resArr = Nothing
        Set col = Nothing
    Else
        MsgBox "δ�ҵ�������б�"
    End If
    Set fso = Nothing
End Sub

Private Sub OptionButton1_Click()
    If Me.OpenDrawing.Value Then
        Me.OpenDrawing.Value = False
    Else
        Me.staticReplace.Value = True
    End If
End Sub

Private Sub CheckBox2Layer_Click()
    If Me.CheckBox2Layer.Value Then
        Me.txtBoxLayer.Enabled = True
        Me.txtBoxLayer.BackStyle = fmBackStyleOpaque
    Else
        Me.txtBoxLayer.Enabled = False
        Me.txtBoxLayer.BackStyle = fmBackStyleTransparent
        Me.txtBoxLayer.Text = vbNullString
    End If
End Sub

'Private Sub CheckBox3Aligment_Click()
'    If Me.CheckBox3Aligment.Value Then
'        Me.TextBox2.Enabled = True
'        Me.TextBox2.BackStyle = fmBackStyleOpaque
'    Else
'        Me.TextBox2.Enabled = False
'        Me.TextBox2.BackStyle = fmBackStyleTransparent
'        Me.TextBox2.Text = vbNullString
'    End If
'End Sub

Private Sub CheckBox4TxtHeight_Click()
    If Me.CheckBox4TxtHeight.Value Then
        Me.txtBoxHeight.Enabled = True
        Me.txtBoxHeight.BackStyle = fmBackStyleOpaque
    Else
        Me.txtBoxHeight.Enabled = False
        Me.txtBoxHeight.BackStyle = fmBackStyleTransparent
        Me.txtBoxHeight.Text = vbNullString
    End If
End Sub


Private Sub OpenDrawing_Click()
    If Me.OpenDrawing.Value Then
        'Me.cmdGoHead.Enabled = True
        'Me.cmdPause.Enabled = True
        Me.staticReplace.Value = False
    Else
        ' Me.cmdGoHead.Enabled = False
        'Me.cmdPause.Enabled = False
        Me.staticReplace.Value = True
    End If
End Sub

Private Sub staticReplace_Click()
    If Me.staticReplace.Value Then
        'Me.cmdGoHead.Enabled = False
        'Me.cmdPause.Enabled = False
        Me.OpenDrawing.Value = False
    Else
        'Me.cmdGoHead.Enabled = True
        'Me.cmdPause.Enabled = True
        Me.OpenDrawing.Value = True
    End If
End Sub

Private Sub txtBoxHeight_Change()
    If Not VBA.IsNumeric(Me.txtBoxHeight.Text) Then Me.txtBoxHeight.Text = "0"
End Sub

Private Sub UserForm_Initialize()
    Set Me.dictfls = CreateObject("Scripting.Dictionary")
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    directory = fso.GetParentFolderName(Application.VBE.ActiveVBProject.FileName) '��ȡ��ǰDVB�ļ����ڵ��ļ��е�ַ
    Set fso = Nothing
    Me.txtBoxHeight.Enabled = False: Me.txtBoxHeight.BackStyle = fmBackStyleTransparent
    Me.txtBoxLayer.Enabled = False: Me.txtBoxLayer.BackStyle = fmBackStyleTransparent
    'Me.TextBox2.Enabled = False: Me.TextBox2.BackStyle = fmBackStyleTransparent
    Me.OpenDrawing.Value = True
End Sub


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
    ExcelApp.Visible = False
    If Err <> 0 Then MsgBox "Could not start Excel!", vbExclamation
    isNewStartExcel = True
    '    Else
    '        isNewStartExcel = False
    '    End If
    '
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
