VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�����ı��滻"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAddFile_Click()
    Dim fls As Collection, fl As Variant, i As Long
    Set fls = win32Api.GetOpenFileName("AutoCAD", ".dwg", "��AutoCAD Drawing �ļ�", True)
    If TypeName(fls) <> "Nothing" Then
        For i = 1 To fls.count
            Me.ListBox1.AddItem fls.Item(i)
        Next
    Else
        MsgBox "ûѡ���κ��ļ�", vbCritical + vbOKOnly
    End If
    
End Sub
