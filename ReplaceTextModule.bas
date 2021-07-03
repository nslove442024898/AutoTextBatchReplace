Attribute VB_Name = "ReplaceTextModule"
Option Explicit

'主函数
Public Function 替换文件内部文字主函数(fn As String, oldstr As String, newStr As String, Optional 开图替换 As Boolean = False, Optional layerName As String = vbNullString, _
        Optional txtheight As Double = 0#, Optional aligmentStyle As AutoCAD.AcAlignment = 0, Optional 完全匹配 As Boolean = False, Optional 替换块内文字 As Boolean = False, Optional 替换属性块内属性 As Boolean = False) As Long
    'AcadApplication.Preferences.OpenSave.SaveAsType = ac2013_dwg
    Dim acDb As AcadDatabase, res As Long
    
    If 开图替换 Then
        Dim curDoc As AcadDocument
        Set curDoc = Application.Documents.Open(fn)
        Set acDb = curDoc.Database
        res = RepalceTxtInBlocks(oldstr, newStr, acDb, layerName, txtheight, aligmentStyle, 完全匹配, 替换块内文字)
        If 替换属性块内属性 Then res = res + RepalceTxtInBlockAttributes(oldstr, newStr, acDb, 完全匹配)
        If res > 0 Then
            curDoc.SaveAs VBA.Replace(fn, ".dwg", "-" & Year(Now) & Month(Now) & Day(Now) & ".dwg")
        End If
        curDoc.Close False
        Set acDb = Nothing: Set curDoc = Nothing
        替换文件内部文字主函数 = res
    Else
        Dim dbx As Object, acadVer As String
        acadVer = VBA.Left(Application.Version, 2)
        Set dbx = AcadApplication.GetInterfaceObject("ObjectDBX.AxDbDocument." & acadVer)
        dbx.Open fn
        Set acDb = dbx.Database
        res = RepalceTxtInBlocks(oldstr, newStr, acDb, layerName, txtheight, aligmentStyle, 完全匹配, 替换块内文字)
        If 替换属性块内属性 Then res = res + RepalceTxtInBlockAttributes(oldstr, newStr, acDb, 完全匹配)
        If res > 0 Then dbx.SaveAs VBA.Replace(fn, ".dwg", "-" & Year(Now) & Month(Now) & Day(Now) & ".dwg")
        Set acDb = Nothing: Set dbx = Nothing
        替换文件内部文字主函数 = res
    End If
    'Shell "explorer.exe " & VBA.Chr(34) & bfd.Self.path & VBA.Chr(34), 1
End Function

'替换块内文字
Private Function RepalceTxtInBlocks(oldstr As String, newStr As String, ByRef acDb As AcadDatabase, Optional layerName As String = vbNullString, _
        Optional txtheight As Double = 0#, Optional aligmentStyle As AutoCAD.AcAlignment = 0, Optional 完全匹配 As Boolean = False, Optional 替换块内文字 As Boolean = False) As Long
    Dim entInblk As AcadEntity, blkdef As AcadBlock, iReplaceCount As Long
    For Each blkdef In acDb.Blocks
        '要识别块的类型，使用IsLayout和IsXRef 属性。假如这两个属性都是 FALSE，则块为一个简单块。假如IsXRef 属性为TRUE，则块是一个外部引用。假如IsLayout属性为TRUE，则块包含所有与块相关的几何图形。
        '  If blkdef.IsLayout Then '布局块代表模型空间与图纸空间布局中的几何图形
        ' ElseIf blkdef.IsXRef = False And blkdef.IsLayout = False Then
        If 替换块内文字 Then
            iReplaceCount = iReplaceCount + RepalceTxtinBlkDefs(blkdef, oldstr, newStr, layerName, txtheight, aligmentStyle, 完全匹配, 替换块内文字)
        Else
            If blkdef.IsLayout Then '模型空间或者布局空间
                iReplaceCount = iReplaceCount + RepalceTxtinBlkDefs(blkdef, oldstr, newStr, layerName, txtheight, aligmentStyle, 完全匹配, 替换块内文字)
            End If
        End If
    Next
    RepalceTxtInBlocks = iReplaceCount
End Function

Private Function RepalceTxtinBlkDefs(ByRef blkdef As AcadBlock, oldstr As String, newStr As String, Optional layerName As String = vbNullString, _
        Optional txtheight As Double = 0#, Optional aligmentStyle As AutoCAD.AcAlignment = 0, Optional 完全匹配 As Boolean = False, Optional 替换块内文字 As Boolean = False)
    Dim entInblk As AcadEntity, iReplaceCount As Long
    For Each entInblk In blkdef
        If TypeOf entInblk Is IAcadText Or TypeOf entInblk Is IAcadMText Then
            If 完全匹配 Then
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
                    If txtheight <> 0 Then entInblk.Height = txtheight '改变文字高度
                    If TypeOf entInblk Is IAcadText Then
                        If aligmentStyle <> 0 Then entInblk.Alignment = aligmentStyle '改变文字对齐方式
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
                    If txtheight <> 0 Then entInblk.Height = txtheight '改变文字高度
                    If TypeOf entInblk Is IAcadText Then
                        If aligmentStyle <> 0 Then entInblk.Alignment = aligmentStyle '改变文字对齐方式
                    End If
                End If
            End If
        End If
    Next
    RepalceTxtinBlkDefs = iReplaceCount
End Function

'替换属性的属性参照
Private Function RepalceTxtInBlockAttributes(oldstr As String, newStr As String, ByRef acDb As AcadDatabase, Optional 完全匹配 As Boolean = False) As Long
    Dim entInblk As AcadEntity, blkref As AcadBlockReference, iReplaceCount As Long
    For Each entInblk In acDb.ModelSpace
        If TypeOf entInblk Is IAcadBlockReference Then
            Set blkref = entInblk
            If blkref.HasAttributes Then
                Dim atts  As Variant, i As Long
                atts = blkref.GetAttributes()
                For i = LBound(atts) To UBound(atts)
                    If 完全匹配 Then
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

