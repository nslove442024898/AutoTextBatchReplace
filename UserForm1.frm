VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "批量文本替换"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAddFile_Click()
    Dim fls As Collection, fl As Variant, i As Long
    Set fls = win32Api.GetOpenFileName("AutoCAD", ".dwg", "打开AutoCAD Drawing 文件", True)
    If TypeName(fls) <> "Nothing" Then
        For i = 1 To fls.count
            Me.ListBox1.AddItem fls.Item(i)
        Next
    Else
        MsgBox "没选中任何文件", vbCritical + vbOKOnly
    End If
    
End Sub

