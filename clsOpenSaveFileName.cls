VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenSaveFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//Available Flags///
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

'private class variables

#If VBA7 Then
    Private lngHwnd As LongPtr
#Else
    Private lngHwnd As Long
#End If

Private strFilter As String
Private strTitle As String
Private strDir As String
Private strFile As String   'elj
Private lngSelectedFilter As Long
Private blnHideReadOnly As Boolean
Private blnMode As Boolean

Private Sub Class_Initialize()
    'Set default values when
    'class is first created
    strDir = CurDir
    strTitle = "Llamas Rule"
    strFile = ""
    strFilter = "All Files" _
        & Chr$(0) & "*.*" & Chr$(0)
    lngSelectedFilter = 0
    lngHwnd = &O0 'Desktop
End Sub

#If VBA7 Then
    Public Property Let OwnerHwnd(WindowHandle As LongPtr)
        lngHwnd = WindowHandle
    End Property
#Else
    Public Property Let OwnerHwnd(WindowHandle As Long)
        '//FOR YOU TODO//
        'Use the API to validate this handle
        lngHwnd = WindowHandle
        'R14 users who just want to use this code:
        'Simple, don't set this property! the default
        'of &0 will work fine for most of your needs
    End Property
#End If

#If VBA7 Then
    Public Property Get OwnerHwnd() As LongPtr
        OwnerHwnd = lngHwnd
    End Property
#Else
    Public Property Get OwnerHwnd() As Long
        OwnerHwnd = lngHwnd
    End Property
#End If
'elj added property
Public Property Let MultiSelect(mode As Boolean)
    blnMode = mode
End Property


'elj added property
Public Property Get MultiSelect() As Boolean
    MultiSelect = blnMode
End Property
Public Property Get SelectedFilter() As Long
    SelectedFilter = lngSelectedFilter
End Property


Public Property Let SelectedFilter(FilterNumber As Long)
    lngSelectedFilter = FilterNumber
End Property

'elj added property
Public Property Let StartFile(FileName As String)
    'don't allow null strings
    If Not FileName = vbNullString Then
        strFile = FileName
    End If
End Property

'elj added property
Public Property Get StartFile() As String
    StartFile = strFile
End Property

Public Property Let StartInDir(StartDir As String)
    'don't allow null strings
    If Not StartDir = vbNullString Then
        strDir = StartDir
    End If
End Property

Public Property Get StartInDir() As String
    StartInDir = strDir
End Property

Public Property Let Title(Caption As String)
    'don't allow null strings
    If Not Caption = vbNullString Then
        strTitle = Caption
    End If
End Property

Public Property Get Title() As String
    Title = strTitle
End Property

Public Property Let Filter(ByVal FilterString As String)
    'Filters change the type of files that are
    'displayed in the dialog. I have designed this
    'validation to use the same filter format the
    'Common dialog OCX uses:
    '"All Files (*.*)|*.*"
    Dim intPos As Integer
    Do While InStr(FilterString, "|") > 0
        intPos = InStr(FilterString, "|")
        If intPos > 0 Then
            FilterString = Left$(FilterString, intPos - 1) _
                & Chr$(0) & Right$(FilterString, _
                Len(FilterString) - intPos)
        End If
    Loop
    If Right$(FilterString, 2) <> Chr$(0) & Chr$(0) Then
        FilterString = FilterString & Chr$(0)
    End If
    strFilter = FilterString
End Property


Public Property Get Filter() As String
    'Here we reverse the process and return
    'the Filter in the same format that it was
    'entered
    Dim intPos As Integer
    Dim strTemp As String
    strTemp = strFilter
    Do While InStr(strTemp, Chr$(0)) > 0
        intPos = InStr(strTemp, Chr$(0))
        If intPos > 0 Then
            strTemp = Left$(strTemp, intPos - 1) _
                & "|" & Right$(strTemp, _
                Len(strTemp) - intPos)
        End If
    Loop
    If Right$(strTemp, 1) = "|" Then
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    Filter = strTemp
End Property

Public Property Let HideReadOnly(blnVal As Boolean)
    'Simple one
    blnHideReadOnly = blnVal
End Property

Public Property Get HideReadOnly() As Boolean
    HideReadOnly = blnHideReadOnly
End Property

'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
' Display and use the File open dialog
'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
Public Function ShowOpen() As String
    Dim strTemp As String
    Dim udtStruct As OPENFILENAME
    udtStruct.lStructSize = Len(udtStruct)
    'Use our private variable
    udtStruct.hWndOwner = lngHwnd
    'Use our private variable
    udtStruct.lpstrFilter = strFilter
    udtStruct.nFilterIndex = lngSelectedFilter
    'elj start
    '  udtStruct.lpstrFile = Space$(254)    elj comment out
    If Not strFile = vbNullString Then
        udtStruct.lpstrFile = strFile & Space(254 - Len(strFile))
    Else
        udtStruct.lpstrFile = Space$(254)
    End If
    'elj end
    udtStruct.nMaxFile = 255
    udtStruct.lpstrFileTitle = Space$(254)
    udtStruct.nMaxFileTitle = 255
    'Use our private variable
    udtStruct.lpstrInitialDir = strDir
    'Use our private variable
    udtStruct.lpstrTitle = strTitle
    'Ok, here we test our boolean to
    'set the flags
    udtStruct.flags = 0
    If blnHideReadOnly Then udtStruct.flags = OFN_HIDEREADONLY + udtStruct.flags
    If blnMode Then udtStruct.flags = OFN_ALLOWMULTISELECT + udtStruct.flags
    If GetOpenFileName(udtStruct) Then
        strTemp = (Trim(udtStruct.lpstrFile))
        ShowOpen = Mid(strTemp, 1, Len(strTemp) - 1)
    End If
End Function

'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
' Display and use the File Save dialog
'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
Public Function ShowSave() As String
    Dim strTemp As String
    Dim udtStruct As OPENFILENAME
    udtStruct.lStructSize = Len(udtStruct)
    'Use our private variable
    udtStruct.hWndOwner = lngHwnd
    'Use our private variable
    udtStruct.lpstrFilter = strFilter
    '  udtStruct.lpstrFile = Space$(254)    'elj comment out
    'elj start
    If Not strFile = vbNullString Then
        udtStruct.lpstrFile = strFile & Space(254 - Len(strFile))
    Else
        udtStruct.lpstrFile = Space$(254)
    End If
    'elj end
    udtStruct.nMaxFile = 255
    udtStruct.lpstrFileTitle = Space$(254)
    udtStruct.nMaxFileTitle = 255
    'Use our private variable
    udtStruct.lpstrInitialDir = strDir
    'Use our private variable
    udtStruct.lpstrTitle = strTitle
    'Ok, here we test our flag
    If blnHideReadOnly Then
        udtStruct.flags = OFN_HIDEREADONLY
    Else
        udtStruct.flags = 0
    End If
    If GetSaveFileName(udtStruct) Then
        strTemp = (Trim(udtStruct.lpstrFile))
        ShowSave = Mid(strTemp, 1, Len(strTemp) - 1)
    End If
End Function
