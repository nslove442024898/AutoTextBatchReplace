VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�����ı��滻"
   ClientHeight    =   7875
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

Private Sub btnAddFile_Click()
    Dim fl As Variant, i As Long, fls As Collection
    Set fls = win32Api.GetOpenFileName("AutoCAD", ".dwg", "��AutoCAD Drawing �ļ�", True)
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
    
    Dim i As Long, var As Variant, fl As String, icount As Long, layerctl As String, txtheight As Double, aligment As AutoCAD.AcAlignment
    If Me.CheckBox2Layer Then layerctl = Me.txtBoxLayer Else layerctl = vbNullString
    If Me.CheckBox4TxtHeight Then txtheight = CDbl(Me.txtBoxHeight) Else txtheight = 0
    If Me.CheckBox3Aligment Then aligment = CInt(Me.TextBox2) Else aligment = 0
    If Me.ListBox1.ListCount > 0 Then
        For i = 0 To Me.ListBox1.ListCount - 1
            DoEvents
            fl = Me.ListBox1.List(i)
            If Me.staticReplace = True Then
                icount = ReplaceTextModule.�滻�ļ��ڲ�����������(fl, Me.txtOldStr.Text, Me.TxtNewStr.Text, False, layerctl, txtheight, _
                    aligment, Me.CheckBox1��ȫƥ��.Value, Me.CheckBox5�滻��������.Value, Me.CheckBox6�滻���Կ�������.Value)
            Else
                icount = ReplaceTextModule.�滻�ļ��ڲ�����������(fl, Me.txtOldStr.Text, Me.TxtNewStr.Text, True, layerctl, txtheight, _
                    aligment, Me.CheckBox1��ȫƥ��.Value, Me.CheckBox5�滻��������.Value, Me.CheckBox6�滻���Կ�������.Value)
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

Private Sub btnCancelAll_Click()
    Dim index As Long
    For index = Me.ListBox1.ListCount - 1 To 0 Step -1
        Me.ListBox1.RemoveItem index
        Me.dictfls.Remove Me.ListBox1.List(index)
    Next
End Sub

Private Sub btnCancelSelect_Click()
    If Me.ListBox1.ListIndex <> -1 Then
        Dim index As Long
        For index = Me.ListBox1.ListCount - 1 To 0 Step -1
            If Me.ListBox1.Selected(index) = True Then
                Me.ListBox1.RemoveItem index
                Me.dictfls.Remove Me.ListBox1.List(index)
            End If
        Next
    End If
End Sub


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
        MsgBox "�ɹ���ȡ�ļ��б�"
    Else
        MsgBox "δ�ҵ�������б�"
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

Private Sub CheckBox3Aligment_Click()
    If Me.CheckBox3Aligment.Value Then
        Me.TextBox2.Enabled = True
        Me.TextBox2.BackStyle = fmBackStyleOpaque
    Else
        Me.TextBox2.Enabled = False
        Me.TextBox2.BackStyle = fmBackStyleTransparent
        Me.TextBox2.Text = vbNullString
    End If
End Sub

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

Private Sub UserForm_Initialize()
    Set Me.dictfls = CreateObject("Scripting.Dictionary")
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    directory = fso.GetParentFolderName(Application.VBE.ActiveVBProject.FileName)
    Set fso = Nothing
    Me.txtBoxHeight.Enabled = False: Me.txtBoxHeight.BackStyle = fmBackStyleTransparent
    Me.txtBoxLayer.Enabled = False: Me.txtBoxLayer.BackStyle = fmBackStyleTransparent
    Me.TextBox2.Enabled = False: Me.TextBox2.BackStyle = fmBackStyleTransparent
    Me.staticReplace.Value = True
End Sub
