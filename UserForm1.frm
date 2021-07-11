VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "批量文本替换"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   3  '窗口缺省
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
    Set fls = GetOpenFileName("AutoCAD", ".dwg", "", True)
    VBA.AppActivate Application.Caption
    If TypeName(fls) <> "Nothing" Then
        For i = 1 To fls.count
            If Not dictfls.Exists(fls.item(i)) Then
                If Not FileInUse(fls.item(i)) Then
                    dictfls.Add fls.item(i), 0
                    Me.ListBox1.AddItem fls.item(i)
                Else
                    Me.Label3 = "文件  " & fls.item(i) & " 处于只读状态,无法添加到工作文件列表"
                End If
            Else
                Me.Label3 = "文件 " & fls.item(i) & " 已经添加,无需重复添加"
            End If
        Next
    Else
        MsgBox "没选中任何文件", vbCritical + vbOKOnly
    End If
    
End Sub

Private Sub btnBatchReplace_Click()
    
    Dim i As Long, var As Variant, fl As String, icount As Long, layerctl As String, txtheight As Double, 完全匹配 As Boolean, 替换块内文字 As Boolean
    'aligment As AutoCAD.AcAlignment
    If Me.CheckBox2Layer Then layerctl = Me.txtBoxLayer Else layerctl = vbNullString
    If Me.CheckBox4TxtHeight Then txtheight = CDbl(Me.txtBoxHeight) Else txtheight = 0
    If Me.CheckBox1完全匹配.Value Then 完全匹配 = True Else 完全匹配 = False
    If Me.CheckBox5替换块内文字.Value Then 替换块内文字 = True Else 替换块内文字 = False
    
    'If Me.CheckBox3Aligment Then aligment = CInt(Me.TextBox2) Else aligment = 0
    If Me.ListBox1.ListCount > 0 Then
        If Me.txtOldStr.Text = vbNullString Or Me.TxtNewStr.Text = vbNullString Then
            Me.Label3 = "请输入需要替换的文字再点击批量替换按钮"
            Exit Sub
        End If
        
        For i = 0 To Me.ListBox1.ListCount - 1
            'Me.curReplaceFileIndex = i
            DoEvents
            fl = Me.ListBox1.List(i)
            If Me.staticReplace = True Then
                icount = ReplaceTextModule.替换文件内部文字主函数(fl, Me.txtOldStr.Text, Me.TxtNewStr.Text, False, layerctl, txtheight, 0, 完全匹配, 替换块内文字)
                '(fl, Me.txtOldStr.Text, Me.TxtNewStr.Text, False, layerctl, txtheight, 完全匹配, _
                    'Me.CheckBox5替换块内文字.Value) ', Me.CheckBox6替换属性块内属性.Value)
            Else
                icount = ReplaceTextModule.替换文件内部文字主函数(fl, Me.txtOldStr.Text, Me.TxtNewStr.Text, True, layerctl, txtheight, 0, 完全匹配, 替换块内文字)
            End If
            If icount > 0 Then
                Me.Label3 = "正在替换 " & fl & ",共 " & Me.ListBox1.ListCount & " 张图纸需要替换,正在替换第 " & i + 1 & " 张图纸,共替换 " & icount & " 个文字"
            Else
                MsgBox "当前图纸 " & fl & "内部找不到需要替换的文字", vbInformation + vbOKOnly
            End If
        Next
        Me.Label3 = "全部替换完成!"
    Else
        MsgBox "未选择任何图纸,无法进行替换,请先添加文件!", vbExclamation
    End If
    
End Sub

'取消所有的行
Private Sub btnCancelAll_Click()
    Dim index As Long
    For index = Me.ListBox1.ListCount - 1 To 0 Step -1
        Me.dictfls.Remove Me.ListBox1.List(index)
        Me.ListBox1.RemoveItem index
    Next
End Sub

'删除选择的行
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

'读取文件列表
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
        Dim 失败 As Boolean
        Me.ListBox1.Clear
        Me.dictfls.RemoveAll
        For i = LBound(Arr) To UBound(Arr)
            If fso.FileExists(Arr(i)) Then
                If FileInUse(Arr(i)) Then
                    Me.Label3 = "文件列表中 " & Chr(34) & Arr(i) & Chr(34) & " 文件读取失败,因为文件被占用,无法操作文件"
                Else
                    Me.dictfls.Add Arr(i), 0
                    Me.ListBox1.AddItem Arr(i)
                End If
            Else
                Me.Label3 = "文件列表中 " & Chr(34) & Arr(i) & Chr(34) & " 文件读取失败,因为文件找不到!"
                失败 = True
            End If
        Next
        If 失败 = False Then Me.Label3 = "成功读取文件列表"
    Else
        Me.Label3 = "未找到保存的列表"
    End If
    Set fso = Nothing
    
End Sub
Private Sub btnsaveList_Click()
    If Me.ListBox1.ListCount > 0 Then
        Dim fso As Object, dataFn As String, sr As Object, var As Variant
        dataFn = Me.directory + "\" + tempSelectFileListName
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(dataFn) Then Set sr = fso.OpenTextFile(dataFn, ForWriting, False) Else Set sr = fso.CreateTextFile(dataFn, True, False)
        For Each var In Me.dictfls.Keys
            sr.WriteLine var
        Next
        
        sr.Close
        Set sr = Nothing
        Me.Label3 = "存取文件列表成功"
        Set fso = Nothing
    Else
        Me.Label3 = "存取文件列表失败,由于文件列表位空!"
    End If
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


Public Function FileInUse(sFileName) As Boolean
    On Error Resume Next
    Open sFileName For Binary Access Read Lock Read As #1
    Close #1
    FileInUse = IIf(Err.Number > 0, True, False)
    On Error GoTo 0
End Function

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
    Dim fso As FileSystemObject, WshShell As Object
    'MsgBox CurDir
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'https://docs.microsoft.com/zh-cn/previous-versions//0ea7b5xe(v=vs.85)?redirectedfrom=MSDN
    Set WshShell = CreateObject("WScript.Shell")
    directory = WshShell.SpecialFolders("MyDocuments")
    Set fso = Nothing: Set WshShell = Nothing
    
    Me.txtBoxHeight.Enabled = False: Me.txtBoxHeight.BackStyle = fmBackStyleTransparent
    Me.txtBoxLayer.Enabled = False: Me.txtBoxLayer.BackStyle = fmBackStyleTransparent
    'Me.TextBox2.Enabled = False: Me.TextBox2.BackStyle = fmBackStyleTransparent
    Me.OpenDrawing.Value = True
End Sub


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
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = False
    If Err <> 0 Then MsgBox "Could not start Excel!", vbExclamation
    isNewStartExcel = True
'    '
'    Dim defualtPath As String ', WshShell As Object
'    'Set WshShell = CreateObject("WScript.Shell")
'    defualtPath = CurDir 'WshShell.SpecialFolders("Desktop")
'    'Set WshShell = Nothing
'    Dim drivename As String
'    drivename = VBA.Left(defualtPath, 3)
'    '默认打开的文件夹
'    ChDrive drivename
'    ChDir defualtPath '
    '
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
