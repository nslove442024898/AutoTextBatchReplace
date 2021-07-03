Attribute VB_Name = "ReplaceTextModule"
Option Explicit

'������
Public Function �滻�ļ��ڲ�����������(fn As String, oldstr As String, newStr As String, Optional ��ͼ�滻 As Boolean = False, Optional layerName As String = vbNullString, _
        Optional txtheight As Double = 0#, Optional aligmentStyle As AutoCAD.AcAlignment = 0, Optional ��ȫƥ�� As Boolean = False, Optional �滻�������� As Boolean = False, Optional �滻���Կ������� As Boolean = False) As Long
    'AcadApplication.Preferences.OpenSave.SaveAsType = ac2013_dwg
    Dim acDb As AcadDatabase, res As Long
    
    If ��ͼ�滻 Then
        Dim curDoc As AcadDocument
        Set curDoc = Application.Documents.Open(fn)
        Set acDb = curDoc.Database
        res = RepalceTxtInBlocks(oldstr, newStr, acDb, layerName, txtheight, aligmentStyle, ��ȫƥ��, �滻��������)
        If �滻���Կ������� Then res = res + RepalceTxtInBlockAttributes(oldstr, newStr, acDb, ��ȫƥ��)
        If res > 0 Then
            curDoc.SaveAs VBA.Replace(fn, ".dwg", "-" & Year(Now) & Month(Now) & Day(Now) & ".dwg")
        End If
        curDoc.Close False
        Set acDb = Nothing: Set curDoc = Nothing
        �滻�ļ��ڲ����������� = res
    Else
        Dim dbx As Object, acadVer As String
        acadVer = VBA.Left(Application.Version, 2)
        Set dbx = AcadApplication.GetInterfaceObject("ObjectDBX.AxDbDocument." & acadVer)
        dbx.Open fn
        Set acDb = dbx.Database
        res = RepalceTxtInBlocks(oldstr, newStr, acDb, layerName, txtheight, aligmentStyle, ��ȫƥ��, �滻��������)
        If �滻���Կ������� Then res = res + RepalceTxtInBlockAttributes(oldstr, newStr, acDb, ��ȫƥ��)
        If res > 0 Then dbx.SaveAs VBA.Replace(fn, ".dwg", "-" & Year(Now) & Month(Now) & Day(Now) & ".dwg")
        Set acDb = Nothing: Set dbx = Nothing
        �滻�ļ��ڲ����������� = res
    End If
    'Shell "explorer.exe " & VBA.Chr(34) & bfd.Self.path & VBA.Chr(34), 1
End Function

'�滻��������
Private Function RepalceTxtInBlocks(oldstr As String, newStr As String, ByRef acDb As AcadDatabase, Optional layerName As String = vbNullString, _
        Optional txtheight As Double = 0#, Optional aligmentStyle As AutoCAD.AcAlignment = 0, Optional ��ȫƥ�� As Boolean = False, Optional �滻�������� As Boolean = False) As Long
    Dim entInblk As AcadEntity, blkdef As AcadBlock, iReplaceCount As Long
    For Each blkdef In acDb.Blocks
        'Ҫʶ�������ͣ�ʹ��IsLayout��IsXRef ���ԡ��������������Զ��� FALSE�����Ϊһ���򵥿顣����IsXRef ����ΪTRUE�������һ���ⲿ���á�����IsLayout����ΪTRUE�����������������صļ���ͼ�Ρ�
        '  If blkdef.IsLayout Then '���ֿ����ģ�Ϳռ���ͼֽ�ռ䲼���еļ���ͼ��
        ' ElseIf blkdef.IsXRef = False And blkdef.IsLayout = False Then
        If �滻�������� Then
            iReplaceCount = iReplaceCount + RepalceTxtinBlkDefs(blkdef, oldstr, newStr, layerName, txtheight, aligmentStyle, ��ȫƥ��, �滻��������)
        Else
            If blkdef.IsLayout Then 'ģ�Ϳռ���߲��ֿռ�
                iReplaceCount = iReplaceCount + RepalceTxtinBlkDefs(blkdef, oldstr, newStr, layerName, txtheight, aligmentStyle, ��ȫƥ��, �滻��������)
            End If
        End If
    Next
    RepalceTxtInBlocks = iReplaceCount
End Function

Private Function RepalceTxtinBlkDefs(ByRef blkdef As AcadBlock, oldstr As String, newStr As String, Optional layerName As String = vbNullString, _
        Optional txtheight As Double = 0#, Optional aligmentStyle As AutoCAD.AcAlignment = 0, Optional ��ȫƥ�� As Boolean = False, Optional �滻�������� As Boolean = False)
    Dim entInblk As AcadEntity, iReplaceCount As Long
    For Each entInblk In blkdef
        If TypeOf entInblk Is IAcadText Or TypeOf entInblk Is IAcadMText Then
            If ��ȫƥ�� Then
                If entInblk.TextString = oldstr Then
                    If layerName <> vbNullString Then
                        If entInblk.Layer = layerName Then
                            entInblk.TextString = VBA.Replace(entInblk.TextString, oldstr, newStr)
                            iReplaceCount = iReplaceCount + 1
                        End If
                    Else
                        entInblk.TextString = VBA.Replace(entInblk.TextString, oldstr, newStr)
                        iReplaceCount = iReplaceCount + 1
                    End If
                    If txtheight <> 0 Then entInblk.Height = txtheight '�ı����ָ߶�
                    If TypeOf entInblk Is IAcadText Then
                        If aligmentStyle <> 0 Then entInblk.Alignment = aligmentStyle '�ı����ֶ��뷽ʽ
                    End If
                End If
            Else
                If entInblk.TextString Like "*" & oldstr & "*" Then
                    If layerName <> vbNullString Then
                        If entInblk.Layer = layerName Then
                            entInblk.TextString = VBA.Replace(entInblk.TextString, oldstr, newStr)
                            iReplaceCount = iReplaceCount + 1
                        End If
                    Else
                        entInblk.TextString = VBA.Replace(entInblk.TextString, oldstr, newStr)
                        iReplaceCount = iReplaceCount + 1
                    End If
                    If txtheight <> 0 Then entInblk.Height = txtheight '�ı����ָ߶�
                    If TypeOf entInblk Is IAcadText Then
                        If aligmentStyle <> 0 Then entInblk.Alignment = aligmentStyle '�ı����ֶ��뷽ʽ
                    End If
                End If
            End If
        End If
    Next
    RepalceTxtinBlkDefs = iReplaceCount
End Function

'�滻���Ե����Բ���
Private Function RepalceTxtInBlockAttributes(oldstr As String, newStr As String, ByRef acDb As AcadDatabase, Optional ��ȫƥ�� As Boolean = False) As Long
    Dim entInblk As AcadEntity, blkref As AcadBlockReference, iReplaceCount As Long
    For Each entInblk In acDb.ModelSpace
        If TypeOf entInblk Is IAcadBlockReference Then
            Set blkref = entInblk
            If blkref.HasAttributes Then
                Dim atts  As Variant, i As Long
                atts = blkref.GetAttributes()
                For i = LBound(atts) To UBound(atts)
                    If ��ȫƥ�� Then
                        If atts(i).TextString = oldstr Then atts(i).TextString = VBA.Replace(entInblk.TextString, oldstr, newStr): iReplaceCount = iReplaceCount + 1
                    Else
                        If atts(i).TextString Like "*" & oldstr & "*" Then atts(i).TextString = VBA.Replace(entInblk.TextString, oldstr, newStr): iReplaceCount = iReplaceCount + 1
                    End If
                Next
            End If
            Set blkref = Nothing
        End If
    Next
    RepalceTxtInBlockAttributes = iReplaceCount
End Function

