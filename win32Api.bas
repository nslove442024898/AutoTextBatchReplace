Attribute VB_Name = "win32Api"

'================================================================
' Source for using VBA with 64 bit and 32 bit office installs
' http://blog.nkadesign.com/2013/vba-for-32-and-64-bit-systems/
'
'================================================================
Option Explicit


'Notice: Don't forget to set the OwnerHwnd property to the
'handle of the calling window in order to bind the dialog
'to the calling window.


'//The Win32 API Functions///

#If VBA7 Then
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
    Public Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
    
    Public Type OPENFILENAME
        lStructSize As Long
        hWndOwner As LongPtr
        hInstance As LongPtr
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As LongPtr
        lpTemplateName As String
    End Type
    
    
#Else
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
    
    Public Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
    
    Public Type OPENFILENAME
        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If

Public Function FileInUse(sFileName) As Boolean
    On Error Resume Next
    Open sFileName For Binary Access Read Lock Read As #1
    Close #1
    FileInUse = IIf(Err.Number > 0, True, False)
    On Error GoTo 0
End Function
