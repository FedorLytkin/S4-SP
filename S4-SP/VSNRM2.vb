﻿Public Class VSNRM2
    Public s4 As S4.TS4App = Form1.ITS4App
    Public tp As TPServer.ITApplication
    'Public ib As ImBase.IImbaseApplication
    Public artID_Main, artId_Sub As Integer
    Public Spletter As String = "|"
    'параметр не выводящий типы объектов и компоненты с ручной связью
    'Public NO_Art_Documentacia As Boolean = 0
    'Public NO_Ru4na9_Sv9z As Boolean = 1
    'Public NO_Sborka As Boolean = 1
    'Public NO_PartSecion As Boolean = 1
    'Public NO_Standart As Boolean = 1
    'Public NO_Pro4ee As Boolean = 1
    'Public NO_Material As Boolean = 1
    'Private part_SB As Boolean = 0
    'параметр заменяющий "~" на " "
    Public ReplaceTildaOnSpace As Boolean = 0
    'параметр заменяющий "?" на "/"
    Public ReplaceQuestionOnDrob As Boolean = 0
    'параметр модульный
    Dim Verartfiltr_str As String = "Verartfiltr" 'parameter name
    Dim Verartfiltr As Boolean = 0
    Dim TPsfilter_str As String = "TPsfilter"
    Dim TPsfilter As Boolean = 0
    Dim CT_ID_in_Query As String
    ''параметр экспортирующий данные о первом уровне вложенности
    'Public OnlyFirstLewvel As Boolean = 0
    ''параметр экспортирующий данные о СОСТАВ ИЗДЕЛИЯ
    'Public NO_Sostav As Boolean = 1
    ''параметр экспортирующий данные о применяемости деталей
    'Public NO_Parts As Boolean = 1
    ''параметр экспортирующий данные о материалах и покупных
    'Public NO_Purchated As Boolean = 1
    ''параметр экспортирующий ПОДРОБНЫЕ данные о материалах и покупных
    'Public NO_Purchated_PDRBN As Boolean = 1

    ''параметр экспортирующий данные о материалах и покупных
    'Public NO_ColorNote As Boolean = 1

    ''параметр экспортирующий данные о материалах и покупных
    'Public NO_ExlProc_Visible As Boolean = 0

    ''параметр экспортирующий данные серчовских параметрах
    'Public NO_S4_Columns As Boolean = 0

    Public myImageList As New ImageList()
    'Public NO_TSE As Boolean = 0

    'массив символов заменителей
    Dim ReplaceChar() As String
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Application.DoEvents()
        generalTree_Bez_Positio(0)
    End Sub

    'отправка sqz запросов напрямую в БД и если раздел 3 тогда 
    Sub generalTree()
        TreeView1.Nodes.Clear()
        artID_Main = SelectBOM()

        If artID_Main <= 0 Then Exit Sub

        Dim Main_ArticleParam_List As Array = Get_Article_Param(artID_Main)
        Dim root = New TreeNode(Main_ArticleParam_List(0) & Spletter & Main_ArticleParam_List(1))
        TreeView1.Nodes.Add(root)

        With s4
            .OpenQuery("select * from PC where proj_aid = " & artID_Main & CT_ID_in_Query)
            Dim max_poz_mun As String = GetMaxPositionInBOM(artID_Main)

            .OpenQuery("select * from PC where proj_aid = " & artID_Main & CT_ID_in_Query)
            For i As Integer = 1 To CInt(max_poz_mun)
                .QueryGoFirst()
                For j As Integer = 1 To .QueryRecordCount
                    Dim TempPos As String = .QueryFieldByName("positio")
                    If i = CInt(TempPos) Then
                        Dim Part_AID As Integer = .QueryFieldByName("Part_AID")
                        Dim COUNT_PC As Integer = .QueryFieldByName("COUNT_PC")
                        Dim RAZDEL As Integer = .QueryFieldByName("RAZDEL")
                        Dim NOTE As String = .QueryFieldByName("NOTE")
                        Dim LINK_TYPE As String = .QueryFieldByName("LINK_TYPE")
                        Dim FORMAT As String = .QueryFieldByName("FORMAT")
                        Dim CTX_ID As Integer = .QueryFieldByName("CTX_ID")

                        Dim ge_Article_Param_List As Array = Get_Article_Param(Part_AID)
                        Dim subNode As TreeNode = TreeNodeAdd(root, Part_AID, COUNT_PC, TempPos,
                                    RAZDEL, NOTE, LINK_TYPE,
                                    FORMAT, CTX_ID, ge_Article_Param_List)

                        If RAZDEL = 3 Then
                            NextLevelInTreeView(subNode, Part_AID, TempPos)
                            'встать на тоже самое место
                            .OpenQuery("select * from PC where proj_aid = " & artID_Main & CT_ID_in_Query)
                            '.QueryGoFirst()
                            'For z As Integer = 0 To .QueryRecordCount
                            '    Dim pos As String = .QueryFieldByName("positio")
                            '    If pos = TempPos Then
                            '        GoTo ifpozRAVNOTempPoz
                            '    End If
                            '    .QueryGoNext()
                            'Next
                        End If
ifpozRAVNOTempPoz:
                        Exit For
                    End If
                    .QueryGoNext()
                Next
            Next

            .CloseQuery()
        End With
    End Sub
    Function GetMaxPositionInBOM(artID As Integer) As String
        Dim temp_Positio, Max_Positio As Integer
        Max_Positio = 0
        Try
            With s4
                .OpenQuery("select MAX(positio) as positio FROM PC WHERE proj_aid = '" & artID & "'" & CT_ID_in_Query)
                GetMaxPositionInBOM = .QueryFieldByName("positio")
                .CloseQuery()
                .OpenQuery("select * FROM PC WHERE proj_aid = '" & artID & "'" & CT_ID_in_Query)
                .QueryGoFirst()
                While .QueryEOF = 0
                    temp_Positio = .QueryFieldByName("positio")
                    If temp_Positio > Max_Positio Then
                        Max_Positio = temp_Positio
                    End If
                    .QueryGoNext()
                End While
                .CloseQuery()
            End With
            Return Max_Positio
        Catch ex As Exception
            MsgBox("Ошибка при определении максимальной позиции" & ex.Message)
            Return Max_Positio
        End Try
    End Function
    Private Function checkChaildNode(ArtID As Integer) As Boolean
        Try
            With s4
                .OpenQuery("select * from PC where proj_aid = " & ArtID & CT_ID_in_Query)
                If .QueryRecordCount > 0 Then
                    checkChaildNode = True
                    .CloseQuery()
                Else
                    checkChaildNode = False
                End If
            End With
        Catch ex As Exception
            MsgBox("Ошибка при проверке на наличие подузлов" & vbNewLine & ex.Message)
            checkChaildNode = False
        End Try
    End Function
    Sub NextLevelInTreeView(node As TreeNode, Proj_Aid As Integer, SubPositio As String)
        With s4
            .OpenQuery("select * from PC where proj_aid = " & Proj_Aid & CT_ID_in_Query)
            Dim max_poz_mun As String = GetMaxPositionInBOM(Proj_Aid)

            .OpenQuery("select * from PC where proj_aid = " & Proj_Aid & CT_ID_in_Query)

            For i As Integer = 1 To CInt(max_poz_mun)
                .QueryGoFirst()
                For j As Integer = 0 To .QueryRecordCount
                    Dim TempPos As String = .QueryFieldByName("positio")
                    Dim Temp_SubPositio As String = SubPositio & "." & TempPos
                    If i = CInt(TempPos) Then
                        Dim PRJLINK_ID As Integer = .QueryFieldByName("PRJLINK_ID")
                        Dim Part_AID As Integer = .QueryFieldByName("Part_AID")
                        Dim COUNT_PC As Integer = .QueryFieldByName("COUNT_PC")
                        Dim RAZDEL As Integer = .QueryFieldByName("RAZDEL")
                        Dim NOTE As String = .QueryFieldByName("NOTE")
                        Dim LINK_TYPE As String = .QueryFieldByName("LINK_TYPE")
                        Dim FORMAT As String = .QueryFieldByName("FORMAT")
                        Dim CTX_ID As Integer = .QueryFieldByName("CTX_ID")

                        Dim ge_Article_Param_List As Array = Get_Article_Param(Part_AID)
                        Dim subNode As TreeNode = TreeNodeAdd(node, Part_AID, COUNT_PC, Temp_SubPositio,
                                    RAZDEL, NOTE, LINK_TYPE,
                                    FORMAT, CTX_ID, ge_Article_Param_List)

                        If RAZDEL = 3 Then
                            NextLevelInTreeView(subNode, Part_AID, Temp_SubPositio)
                            'встать на тоже самое место
                            .OpenQuery("select * from PC where proj_aid = " & Proj_Aid & CT_ID_in_Query)
                            '.QueryGoFirst()
                            'For z As Integer = 0 To .QueryRecordCount
                            '    Dim pos As String = .QueryFieldByName("positio")
                            '    If pos = TempPos Then
                            '        GoTo ifpozRAVNOTempPoz
                            '    End If
                            '    .QueryGoNext()
                            'Next
                        End If
ifpozRAVNOTempPoz:
                        Exit For
                    End If
                    .QueryGoNext()
                Next
            Next

            .CloseQuery()
        End With
    End Sub
    Function TreeNodeAdd(Node As TreeNode, Part_AID As Integer, COUNT_PC As Integer, POSITIO As String,
                    RAZDEL As Integer, NOTE As String, LINK_TYPE As String,
                    FORMAT As String, CTX_ID As Integer, ge_Article_Param_List As Array) As TreeNode
        Dim NewNode As New TreeNode(POSITIO & Spletter & ge_Article_Param_List(0) & Spletter & ge_Article_Param_List(1))

        Node.Nodes.Add(NewNode)
        Return NewNode
    End Function



































    'через дерево состава
    Public Sub report()
        TreeView1.Nodes.Clear()
        artID_Main = 14059 'SelectBOM()

        With s4
            .OpenArticle(artID_Main)
            Dim root = New TreeNode(.GetFieldValue_Articles("Обозначение") & " - " & .GetFieldValue_Articles("Наименование"))
            .CloseArticle()
            TreeView1.Nodes.Add(root)
            .APIConditionMode = 1
            .OpenArticleStructure2(artID_Main, -1, 2)
            .asFirst()
            While .asEof = 0
                Dim SubArtID As Integer = .asGetArtID
                NodAdd(root, SubArtID)
                .asNext()
            End While
            .CloseArticleStructure()
        End With
    End Sub

    Public Sub NodAdd(node As TreeNode, ArtID As Integer)
        With s4
            Dim poz, oboz_sub, naim_sub As String
            Dim ArtID_InStructure, RazdelSP As Integer
            ArtID_InStructure = .asGetArtID
            poz = .asGetPosition
            oboz_sub = .asGetArtDesignation
            naim_sub = .asGetArtName
            RazdelSP = .asGetArtKind


            Dim NewNode As New TreeNode(poz & Spletter & oboz_sub & Spletter & naim_sub)
            node.Nodes.Add(NewNode)
            If RazdelSP = 3 Then
                reportArtStructure(ArtID_InStructure, NewNode)
            End If
            reportArtStructureFindArtID(ArtID, ArtID_InStructure, node)
        End With
    End Sub
    Sub reportArtStructure(artId As Integer, node As TreeNode)
        Try
            With s4
                .OpenArticleStructure2(artId, -1, 2)
                .asFirst()
                While .asEof = 0
                    Dim SubArtID As Integer = .asGetArtID
                    NodAdd(node, SubArtID)
                    reportArtStructureFindArtID(artId, SubArtID, node)
                    .asNext()
                End While
                .CloseArticleStructure()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub reportArtStructureFindArtID(ProjId As Integer, artId_Find As Integer, node As TreeNode)
        Dim sum As Integer
        Dim findArtCheck As Boolean = False
        Try
            With s4
                .OpenArticleStructure2(ProjId, -1, 2)
                .asFirst()
                For i As Integer = 0 To .asCount - 1
                    If .asGetArtID <> artId_Find Then sum = +1 Else findArtCheck = True

                    If findArtCheck And sum < i Then
                        Dim SubArtID As Integer = .asGetArtID
                        NodAdd(node, .asGetArtID)
                        'reportArtStructureFindArtID(ProjId, SubArtID, node)
                    End If

                    .asNext()
                Next
                .CloseArticleStructure()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub TPServerInitializ()
        Try
            'If TPsfilter = 0 Then Exit Sub

            If tp Is Nothing Then
                tp = CreateObject("TPServer.TApplication")
                Dim status As Integer = tp.Ready
                'While True
                '    If status = 1 Then Exit While
                'End While
            End If
        Catch ex As Exception
            MsgBox("Приложение не может запустить TechCard.Попробуйте перезапустить Search и TechCard..." & vbNewLine & ex.Message)
        End Try
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs)
        'ib = CreateObject("imbase.ImbaseApplication")
        'Dim IBKey As ImBase.ImbaseKey = ib.Utilites.ExploreKey("i6000001087D7D00000A")
        'Dim TableRecStr As String = IBKey.TableRecStr
        'Dim TableRecStrArray As Array = TableRecStr.Split(vbCrLf)
        'Dim PolnOboz, PolnObozStr As String
        'PolnObozStr = "Полное обозначение"
        'For Each str As String In TableRecStrArray
        '    str = str.Trim(vbCrLf)
        '    If str.IndexOf("Полное обозначение") > 0 Then
        '        PolnOboz = str.Replace(PolnObozStr & "=", "")
        '        PolnOboz = PolnOboz.Trim(vbLf)
        '    End If
        'Next
        ''Dim IBKey As ImBase.IImbaseKey = i
        ''Dim cat As ImBase.IImbaseCatalogs = ib.Catalogs()
        ''Dim VR As ImBase.IImbaseCatalog = cat.Item(13933)
        ''For i As Integer = 0 To VR.Folders.Count  'cat.Count - 1
        ''    Dim fName As String = VR.Folders.Item(i).Name
        ''Next
    End Sub
    'Function Cet_PolObozByIMBkey(IMBkey As String) As String
    '    Try
    '        Dim IBKey As ImBase.ImbaseKey = ib.Utilites.ExploreKey(IMBkey)
    '        Dim TableRecStr As String = IBKey.TableRecStr
    '        Dim TableRecStrArray As Array = TableRecStr.Split(vbCrLf)
    '        Dim PolnOboz, PolnObozStr As String
    '        PolnObozStr = "Полное обозначение"
    '        For Each str As String In TableRecStrArray
    '            str = str.Trim(vbCrLf)
    '            If str.IndexOf("Полное обозначение") > 0 Then
    '                PolnOboz = str.Replace(PolnObozStr & "=", "")
    '                PolnOboz = PolnOboz.Trim(vbLf)
    '            End If
    '        Next
    '        If PolnOboz Is Nothing Then
    '            MsgBox("Ошибка!" &
    '                   vbNewLine & "В таблице " & IBKey.TableName & " расположенной по адресу " & IBKey.Path _
    '                   & vbNewLine & "Не могу найти поле 'ПОЛНОЕ ОБОЗНАЧЕНИЕ'")
    '        Else
    '            Return PolnOboz
    '        End If
    '    Catch ex As Exception
    '        MsgBox("Ошибка при поиске сортамента по Ключу ImBase" & vbNewLine & ex.Message)
    '    End Try
    'End Function
    Public Function SelectBOM() As Integer
        Try
            With s4
                .StartSelectArticles()
                .SelectArticles()
                SelectBOM = .GetSelectedArticleID(0)
                .EndSelectArticles()
                If SelectBOM > 0 Then
                    .AddArticleUserEventToLog(SelectBOM, .ErrorCode, "Для объекта был выгружен отчет о составе в модуле ВСНРМ2.0")
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return SelectBOM
    End Function
    Function check_OnOffOptionsforBOMComponents(RAZDEL As Integer, LINK_TYPE As String, CTX_ID As Integer) As Boolean

        Try
            If NO_Art_Documentacia = False Then
                If RAZDEL = 1 Then
                    Return False
                End If
            End If
            If NO_Sborka = False Then
                If RAZDEL = 3 Then
                    Return False
                End If
            End If
            If NO_PartSecion = False Then
                If RAZDEL = 4 Then
                    Return False
                End If
            End If
            If NO_Standart = False Then
                If RAZDEL = 5 Then
                    Return False
                End If
            End If
            If NO_Pro4ee = False Then
                If RAZDEL = 6 Then
                    Return False
                End If
            End If
            If NO_Material = False Then
                If RAZDEL = 7 Then
                    Return False
                End If
            End If
            If NO_Ru4na9_Sv9z = False Then
                If LINK_TYPE = "M" Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function
    Function GetMaxRowInBOM(artID As Integer, PROJ_VER_ID As String)
        Dim Max_Positio As Integer
        Max_Positio = 0
        Try
            With s4
                If Verartfiltr Then
                    .OpenQuery("select * FROM PC WHERE proj_aid = '" & artID & "' And (PROJ_VER_ID = NULL OR PROJ_VER_ID = " & PROJ_VER_ID & ")" & CT_ID_in_Query)
                Else
                    .OpenQuery("select * FROM PC WHERE proj_aid = '" & artID & "'" & CT_ID_in_Query) ' And PROJ_VER_ID = " & PROJ_VER_ID)
                End If

                Max_Positio = .QueryRecordCount()
                .CloseQuery()
            End With
            Return Max_Positio
        Catch ex As Exception
            MsgBox("Ошибка при определении максимальной позиции" & vbNewLine & ex.Message)
            Return Max_Positio
        End Try
    End Function
    Function exist_BOM_ChildNodesExist(ArtID As Integer) As Boolean
        Dim Max_Positio As Integer
        Try
            With s4
                .OpenQuery("select * FROM PC WHERE proj_aid = '" & ArtID & "'" & CT_ID_in_Query)
                Max_Positio = .QueryRecordCount()
                .CloseQuery()
                If Max_Positio > 0 Then
                    Return True
                Else
                    Return False
                End If
            End With
        Catch ex As Exception
            MsgBox("Ошибка при определении поиска вхождения" & vbNewLine & ex.Message)
            Return False
        End Try
    End Function

    Function TreeNodeAdd2(Node As TreeNode, Part_AID As Integer, PRJLINK_ID As Integer, POSITIO As String, count As Double) As TreeNode
        Application.DoEvents()
        Dim nodeParamArray As Array = Node.Tag.ToString.Split(Spletter)
        Dim Count_Parrent As Double
        Try
            Count_Parrent = CDbl(nodeParamArray(3))
            If Count_Parrent = 0 Then Count_Parrent = 1
        Catch ex As Exception
            If Count_Parrent = 0 Then Count_Parrent = 1
        End Try
        Dim oboz, naim As String
        Dim ArtKind As Integer
        With s4
            .OpenArticle(Part_AID)
            oboz = .GetFieldValue_Articles("Обозначение")
            naim = .GetFieldValue_Articles("Наименование")
            ArtKind = .GetArticleKind
            .CloseArticle()
        End With
        Dim NewNode As New TreeNode(POSITIO & Spletter & oboz & Spletter & naim)
        NewNode.Tag = POSITIO & Spletter & Part_AID & Spletter & PRJLINK_ID & Spletter & count * Count_Parrent
        Select Case ArtKind
            Case 1
                NewNode.ImageIndex = 1
            Case 2
                NewNode.ImageIndex = 2
            Case 3
                NewNode.ImageIndex = 0
            Case 4
                NewNode.ImageIndex = 4
            Case 5
                NewNode.ImageIndex = 6
            Case 6
                NewNode.ImageIndex = 5
            Case 7
                NewNode.ImageIndex = 3
        End Select

        Node.Nodes.Add(NewNode)
        Return NewNode
    End Function
    Function chechUtvArch_func(ArtID As Integer) As Boolean
        With s4
            .OpenDocument(.GetDocID_ByArtID(ArtID))
            Dim ArchId As String = .GetFieldValue("Архив")
            .CloseDocument()
            If ArchId <> 7 Then
                Return True
            Else
                If MsgBox("Документ не занесен в Утв.архив. Параметры сборки еще могут изменяться" & vbNewLine & "Продолжить?", vbYesNo + vbQuestion, "ВСНРМ 2.0") = DialogResult.Yes Then
                    Return True
                Else
                    Return False
                End If
            End If
        End With
    End Function
    Function check_SP_Funk(ArtId As Integer) As Boolean
        With s4
            .OpenArticle(ArtId)
            Dim ArtKind As Integer = .GetArticleKind()
            Dim ArtKind_Name As String = .GetArtKindName
            .CloseArticle()
            If ArtKind <> 3 Then
                If MsgBox(ArtKind_Name & " может не содержать вложений!" & vbNewLine & "Продолжить?", vbYesNo + MsgBoxStyle.Question, "ВСНРМ 2.0") = MsgBoxResult.Yes Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        End With
    End Function
    Sub ImagesAdd()
        Try
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "ASSEMBLY.ICO"))
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "DOCUMENT.ICO"))
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "KOMPLEKT.ICO"))
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "MATERIAL.ICO"))
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "PART.ICO"))
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "RADIO2.ICO"))
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "STANDARD.ICO"))
            myImageList.Images.Add(Image.FromFile(Application.StartupPath & "\imagelist\" & "SET.ICO"))
        Catch ex As Exception

        End Try
    End Sub
    Sub generalTree_Bez_Positio(TreeUpdate As Boolean)
        Application.DoEvents()
        If Not TreeUpdate Then
            artID_Main = SelectBOM()
        End If

        If artID_Main <= 0 Then Exit Sub
        Dim chechUtvArch As Boolean = chechUtvArch_func(artID_Main)
        If Not chechUtvArch Then
            Exit Sub
        End If
        Dim check_SP As Boolean = check_SP_Funk(artID_Main)
        If Not check_SP Then
            Exit Sub
        End If
        TreeView1.Nodes.Clear()
        Dim Main_ArticleParam_List As Array = Get_Short_Article_Param(artID_Main)
        Dim root = New TreeNode(Main_ArticleParam_List(0) & Spletter & Main_ArticleParam_List(1))
        root.Tag = artID_Main
        ImagesAdd()

        TreeView1.ImageList = myImageList
        TreeView1.Nodes.Add(root)

        With s4
            .OpenArticle(artID_Main)
            Dim VERS As String = .GetFieldValue_Articles("Art_Ver_ID")
            .CloseArticle()

            Dim max_poz_mun As Integer = GetMaxRowInBOM(artID_Main, VERS)
            PBarLoadSQL(max_poz_mun)

            If Verartfiltr Then
                .OpenQuery("select * from PC where proj_aid = " & artID_Main & " and (PROJ_VER_ID = NULL OR PROJ_VER_ID = " & VERS & ")" & CT_ID_in_Query)
            Else
                .OpenQuery("select * from PC where proj_aid = " & artID_Main & CT_ID_in_Query) ' & " and PROJ_VER_ID = " & VERS)
            End If
            .QueryGoFirst()
            For i As Integer = 0 To max_poz_mun - 1
                Dim PRJLINK_ID As Integer = .QueryFieldByName("PRJLINK_ID")
                Dim TempPos As String = .QueryFieldByName("positio")
                Dim Part_AID As Integer = .QueryFieldByName("Part_AID")
                Dim RAZDEL As Integer = .QueryFieldByName("RAZDEL")
                Dim LINK_TYPE As String = .QueryFieldByName("LINK_TYPE")
                Dim CTX_ID As Integer = .QueryFieldByName("CTX_ID")
                Dim COUNT_PC As Double = .QueryFieldByName("COUNT_PC")

                'Dim ge_Article_Param_List As Array = Get_Article_Param(Part_AID)
                If check_OnOffOptionsforBOMComponents(RAZDEL, LINK_TYPE, CTX_ID) Then
                    Dim subNode As TreeNode = TreeNodeAdd2(root, Part_AID, PRJLINK_ID, TempPos, COUNT_PC)

                    If exist_BOM_ChildNodesExist(Part_AID) And (Not OnlyFirstLewvel Or RAZDEL = 4) Then
                        NextLevelInTreeView_Bez_Positio(subNode, Part_AID, TempPos)
                        Perehov_Na_NUGHUY_strocu(artID_Main, i, VERS)
                    Else
                        Perehov_Na_NUGHUY_strocu(artID_Main, i, VERS)
                    End If
                Else
                    Perehov_Na_NUGHUY_strocu(artID_Main, i, VERS)
                End If
                PBarStep()
            Next

            .CloseQuery()
        End With
    End Sub
    Sub Perehov_Na_NUGHUY_strocu(ArtID As Integer, rownum As Integer, PROJ_VER_ID As String)
        Application.DoEvents()
        Try
            With s4
                If Verartfiltr Then
                    .OpenQuery("select * FROM PC WHERE proj_aid = '" & ArtID & "' and (PROJ_VER_ID = NULL OR PROJ_VER_ID = " & PROJ_VER_ID & ")" & CT_ID_in_Query)
                Else
                    .OpenQuery("select * FROM PC WHERE proj_aid = '" & ArtID & "'" & CT_ID_in_Query) ' & " and PROJ_VER_ID = " & PROJ_VER_ID)
                End If
                .QueryGoFirst()

                For i As Integer = 0 To rownum
                    .QueryGoNext()
                Next
            End With
        Catch ex As Exception
            MsgBox("Ошибка при переходе на следующую позицию" & vbNewLine & ex.Message)
        End Try
    End Sub
    Sub firstAppShow()
        Try
            If Not GetReestrKeyExist(OPT_FirstShow) Then
                info()
                set_reesrt_value(OPT_FirstShow, 1)
            Else
                set_reesrt_value(OPT_FirstShow, CInt(get_reesrt_value(OPT_FirstShow, 1)) + 1)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub VSNRM2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Application.DoEvents()
        ReadOptionsAll() 'считываем параметры с блокнота с опциями
        БезРазделаДокументацияToolStripMenuItem.Checked = NO_Art_Documentacia
        БезТехнолическихСвязейToolStripMenuItem.Checked = NO_Ru4na9_Sv9z
        СПеречнемДеталейToolStripMenuItem.Checked = NO_Parts
        СПеречнемМатериаловToolStripMenuItem.Checked = NO_Purchated
        ЗаменятьToolStripMenuItem.Checked = ReplaceTildaOnSpace
        ПоказыватьПроцессЭкспортаToolStripMenuItem.Checked = NO_ExlProc_Visible
        ПоказыватьСистемнуюИнформациюToolStripMenuItem.Checked = NO_S4_Columns
        СРазделомДеталиToolStripMenuItem.Checked = NO_PartSecion
        СРазделомСборочныеЕдиницыToolStripMenuItem.Checked = NO_Sborka
        СРазделомСтандартныеЕдиницыToolStripMenuItem.Checked = NO_Standart
        СРазделомПрочиеИзделияToolStripMenuItem.Checked = NO_Pro4ee
        СРазделомМатериалыToolStripMenuItem.Checked = NO_Material
        ПоказыватьПояснениеПоЗаливкеToolStripMenuItem.Checked = NO_ColorNote
        ДетальСборочнаяЕдиницаToolStripMenuItem.Checked = Part_SB
        ЗаменятьНаToolStripMenuItem.Checked = ReplaceQuestionOnDrob
        СПеречнемПокупныхПДРБНToolStripMenuItem.Checked = NO_Purchated_PDRBN
        ТолькоПервыйУровеньToolStripMenuItem.Checked = OnlyFirstLewvel
        ССоставомИзделияToolStripMenuItem.Checked = NO_Sostav
        TPServerInitializ()
        firstAppShow()
        OutOptionsLoad()

        If ToolStripComboBox1.SelectedItem Is Nothing Or ToolStripComboBox1.SelectedItem = "" Then '
            ToolStripComboBox1.SelectedItem = ToolStripComboBox1.Items(0)
            'CT_ID_in_Query_Change()
        End If
    End Sub
    Private Sub OutOptionsLoad()
        Try
            Dim file As System.IO.StreamReader = New System.IO.StreamReader("\\INTERMECH\im\Search\Search4.ini")

            Dim counter As Integer = 0
            Dim line As String = " "
            While line IsNot Nothing
                line = file.ReadLine()
                'проверка актуальной версии
                If UCase(line).IndexOf(UCase(Verartfiltr_str)) > 0 Then
                    Try
                        Verartfiltr = (line.Substring(line.LastIndexOf("=") + 1).Trim(" "))
                    Catch ex As Exception

                    End Try
                End If
                'соединение с TPServ
                If UCase(line).IndexOf(UCase(TPsfilter_str)) > 0 Then
                    Try
                        TPsfilter = (line.Substring(line.LastIndexOf("=") + 1).Trim(" "))
                    Catch ex As Exception

                    End Try
                End If
                counter += 1
            End While
            file.Close()
        Catch ex As Exception

        End Try
    End Sub


    'Case 1 'документация
    '            temp_color = Color.Aqua
    '        Case 2 'комплексы
    '            temp_color = Color.Aquamarine
    '        Case 3 'сборки
    '            temp_color = Color.Yellow
    '        Case 4 'детали
    '            temp_color = Color.FromArgb(216, 228, 188) 'Color.DarkKhaki 'GreenYellow (255, 153, 0)
    '        Case 5 'стандартные
    '            temp_color = Color.IndianRed
    '        Case 6 'прочие
    '            temp_color = Color.LightBlue
    '        Case 7 'материалы
    '            temp_color = Color.Khaki
    '        Case 8 'комплекты
    '            temp_color = Color.LightPink

    Public AssemblyColor_Sostav As Color = Color.Yellow
    Public DocumColor_Sostav As Color = Color.Aqua
    Public PartColor_Sostav As Color = Color.FromArgb(216, 228, 188)
    Public StandartColor_Sostav As Color = Color.IndianRed
    Public Pro4eeColor_Sostav As Color = Color.LightBlue
    Public MaterialColor_Sostav As Color = Color.Khaki
    Public KomplekTColor_Sostav As Color = Color.LightPink
    Public TechContex_Sostav As Color = Color.Gray ' Color.LightPink

    Public StandartColor_Purchated As Color = Color.LightGreen
    Public Pro4eeColor_Purchated As Color = Color.LightGreen
    Public VspomMaterialColor_Purchated As Color = Color.DarkSalmon
    Public SortamentColor_Purchated As Color = Color.Tan
    Public SortamentColorKonstructor_Purchated As Color = Color.LightSteelBlue
    Public MaterialColor_Purchated As Color = Color.LightSkyBlue


    Public OPT_FirstShow As String = "ExcelReportFirstShow"
    Public OPT_LastInPut As String = "LastInPut"

    'ExcelSheets column numbers ВТОРОЙ ЛИСТ(ПЕРЕЧЕНЬ ПОКУПНЫХ(например))
    Public ShName_Purchated As String = "ПЕРЕЧЕНЬ ПОКУПНЫХ"
    Public RowN_Purchated_First As Integer = 3
    Public CN_Purchated_Naim As Integer = 1
    Public CN_Purchated_IBKey As Integer = 2
    Public CN_Purchated_Count As Integer = 3
    Public CN_Purchated_MU As Integer = 4
    Private Sub addExTitlePurchated(tmp_node As TreeNode)
        Add_NewSheet(ShName_Purchated)
        Sheet_Delete_by_SheetName("Лист1")
        Dim ArtID, PRJLINK_ID As Integer
        Try
            Dim MainTPInfo, MainArtInfo As Array
            Dim Positio, Count_Summ As String
            If tmp_node.Tag.ToString.IndexOf(Spletter) > 0 Then
                Dim param_Array As Array = tmp_node.Tag.ToString.Split(Spletter)
                'MsgBox(param_Array(0) & Spletter & param_Array(1) & Spletter & param_Array(2))
                Positio = param_Array(0)
                ArtID = param_Array(1)
                PRJLINK_ID = param_Array(2)
                Count_Summ = param_Array(3)
            Else
                ArtID = tmp_node.Tag
            End If
            MainArtInfo = Get_Short_Article_Param(ArtID)

            set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, RowN_Purchated_First - 2, MainArtInfo(0) & " " & MainArtInfo(1))
            'выделить первую строку зеленым цветом
            SetCellsColor(ShName_Purchated, 1, 1, CN_Purchated_MU, 1, Color.LawnGreen)

            CellsTextBold(ShName_Purchated, 1, 1, CN_Purchated_MU, 1)
        Catch ex As Exception

        End Try
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, RowN_Purchated_First - 1, "Наименование")
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First - 1, "Ключ IMBASE")
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, RowN_Purchated_First - 1, "Общ. кол-во/масса")
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, RowN_Purchated_First - 1, "Ед.изм.")

        Dim CN_Purchated_last_ColNum As Integer = Get_Last_Column(ShName_Purchated, RowN_Purchated_First - 1)
        SetCellsBorderLineStyle2(ShName_Purchated, 1, 1, CN_Purchated_last_ColNum, RowN_Purchated_First - 1, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
        Try
            Dim VSpom_Mater_for_Main As Array = Get_Vspom_Mater_Array(ArtID, -1)
            If VSpom_Mater_for_Main IsNot Nothing Then
                excel_write_aboutVSPomMater_in_Purchated(VSpom_Mater_for_Main, 1)
            End If
        Catch ex As Exception
        End Try
        'зафиксировать шапку
        SetTitleFixirovano(ShName_Purchated, CN_Purchated_Naim, RowN_Purchated_First - 1, True)
    End Sub

    'ExcelSheets column numbers ТРЕТИЙ ЛИСТ(ПЕРЕЧЕНЬ ДЕТАЛЕЙ(например))
    Public ShName_pdrbn As String = "ПОДРОБНЫЙ ПЕРЕЧЕНЬ ПОКУПНЫХ"
    Public RowN_pdrbn_First As Integer = 3
    Public CN_pdrbn_Naim As Integer = 1
    Public CN_pdrbn_IBKey As Integer = 2
    Public CN_pdrbn_ObjType As Integer = 3
    Public CN_pdrbn_prjid As Integer = 4
    Public CN_pdrbn_Primen9emost As Integer = 5
    Public CN_pdrbn_Count As Integer = 6
    Public CN_pdrbn_MU As Integer = 7
    Public CN_pdrbn_TotalCount As Integer = 8
    Private Sub addExTitlePDRBN(tmp_node As TreeNode)
        Add_NewSheet(ShName_pdrbn)
        Sheet_Delete_by_SheetName("Лист1")
        Dim ArtID, PRJLINK_ID As Integer
        Try
            Dim MainTPInfo, MainArtInfo As Array
            Dim Positio, Count_Summ As String
            If tmp_node.Tag.ToString.IndexOf(Spletter) > 0 Then
                Dim param_Array As Array = tmp_node.Tag.ToString.Split(Spletter)
                'MsgBox(param_Array(0) & Spletter & param_Array(1) & Spletter & param_Array(2))
                Positio = param_Array(0)
                ArtID = param_Array(1)
                PRJLINK_ID = param_Array(2)
                Count_Summ = param_Array(3)
            Else
                ArtID = tmp_node.Tag
            End If
            MainArtInfo = Get_Short_Article_Param(ArtID)

            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Naim, RowN_pdrbn_First - 2, MainArtInfo(0) & " " & MainArtInfo(1))
            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, RowN_pdrbn_First - 2, ArtID)
            'выделить первую строку зеленым цветом
            SetCellsColor(ShName_pdrbn, 1, 1, CN_pdrbn_TotalCount, 1, Color.LawnGreen)

            CellsTextBold(ShName_pdrbn, 1, 1, CN_pdrbn_TotalCount, 1)
        Catch ex As Exception

        End Try
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Naim, RowN_pdrbn_First - 1, "Наименование")
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_IBKey, RowN_pdrbn_First - 1, "Ключ IMBASE")
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, RowN_pdrbn_First - 1, "Тип объекта")
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, RowN_pdrbn_First - 1, "PROJID")
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, RowN_pdrbn_First - 1, "Применяемость")
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, RowN_pdrbn_First - 1, "Кол-во/масса")
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, RowN_pdrbn_First - 1, "Общ. кол-во/масса")
        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, RowN_pdrbn_First - 1, "Ед.изм.")

        Dim CN_pdrbn_last_ColNum As Integer = Get_Last_Column(ShName_pdrbn, RowN_pdrbn_First - 1)
        SetCellsBorderLineStyle2(ShName_pdrbn, 1, 1, CN_pdrbn_last_ColNum, RowN_pdrbn_First - 1, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
        Try
            Dim VSpom_Mater_for_Main As Array = Get_Vspom_Mater_Array(ArtID, -1)
            If VSpom_Mater_for_Main IsNot Nothing Then
                'тут создать процедуру для записи вспомогательных материалов
                excel_write_aboutVSPomMater_in_Purchated_PDRBN(VSpom_Mater_for_Main, 1, ArtID)
            End If
        Catch ex As Exception
        End Try
        'зафиксировать шапку
        SetTitleFixirovano(ShName_pdrbn, CN_Purchated_Naim, RowN_pdrbn_First - 1, True)
    End Sub

    'ExcelSheets column numbers ТРЕТИЙ ЛИСТ(ПЕРЕЧЕНЬ ДЕТАЛЕЙ(например))
    Public ShName_PartList As String = "ПЕРЕЧЕНЬ ДЕТАЛЕЙ"
    Public RowN_PL_First As Integer = 3
    Public CN_PL_ArtID As Integer = 1
    Public CN_PL_Oboz As Integer = 2
    Public CN_PL_Naim As Integer = 3
    Public CN_PL_PROJ_ID As Integer = 4
    Public CN_PL_Primen9emost As Integer = 5
    Public CN_PL_Count As Integer = 6
    Public CN_PL_TotalCount As Integer = 7
    'Public CN_PL_MU As Integer = 8
    Public CN_PL_LinkType As Integer = 8
    Public CN_PL_ZagCount As Integer = 9
    Public CN_PL_ZagSumCount As Integer = 10
    Public CN_PL_Sort As Integer = 11
    Public CN_PL_Sort_IBKey As Integer = 12
    Public CN_PL_Referenc As Integer = 13
    Private Sub addExTitlePartList(tmp_node As TreeNode)
        Add_NewSheet(ShName_PartList)
        Sheet_Delete_by_SheetName("Лист1")
        Try
            Dim MainTPInfo, MainArtInfo As Array
            Dim ArtID, PRJLINK_ID As Integer
            Dim Positio, Count_Summ As String
            If tmp_node.Tag.ToString.IndexOf(Spletter) > 0 Then
                Dim param_Array As Array = tmp_node.Tag.ToString.Split(Spletter)
                'MsgBox(param_Array(0) & Spletter & param_Array(1) & Spletter & param_Array(2))
                Positio = param_Array(0)
                ArtID = param_Array(1)
                PRJLINK_ID = param_Array(2)
                Count_Summ = param_Array(3)
            Else
                ArtID = tmp_node.Tag
            End If
            MainArtInfo = Get_Short_Article_Param(ArtID)

            set_Value_From_Cell(ShName_PartList, CN_PL_ArtID, RowN_PL_First - 2, ArtID)
            set_Value_From_Cell(ShName_PartList, CN_PL_Oboz, RowN_PL_First - 2, MainArtInfo(0))
            set_Value_From_Cell(ShName_PartList, CN_PL_Naim, RowN_PL_First - 2, MainArtInfo(1))
            'выделить первую строку зеленым цветом
            SetCellsColor(ShName_PartList, 1, 1, CN_PL_Sort_IBKey, 1, Color.LawnGreen)

            CellsTextBold(ShName_PartList, 1, 1, CN_PL_Sort_IBKey, 1)
        Catch ex As Exception

        End Try


        set_Value_From_Cell(ShName_PartList, CN_PL_ArtID, RowN_PL_First - 1, "ArtID")
        set_Value_From_Cell(ShName_PartList, CN_PL_Oboz, RowN_PL_First - 1, "Обозначение Детали")
        set_Value_From_Cell(ShName_PartList, CN_PL_Naim, RowN_PL_First - 1, "Наименование Детали")
        set_Value_From_Cell(ShName_PartList, CN_PL_PROJ_ID, RowN_PL_First - 1, "PROJID")
        set_Value_From_Cell(ShName_PartList, CN_PL_Primen9emost, RowN_PL_First - 1, "Применяемость")
        set_Value_From_Cell(ShName_PartList, CN_PL_Count, RowN_PL_First - 1, "Кол-во на сб, шт.")
        set_Value_From_Cell(ShName_PartList, CN_PL_TotalCount, RowN_PL_First - 1, "Общ.кол-во, шт.")
        'set_Value_From_Cell(ShName_PartList, CN_PL_MU, RowN_PL_First - 1, "ед.изм")
        set_Value_From_Cell(ShName_PartList, CN_PL_LinkType, RowN_PL_First - 1, "Тип связи")
        set_Value_From_Cell(ShName_PartList, CN_PL_ZagCount, RowN_PL_First - 1, "Кол-во деталей из заготовки")
        set_Value_From_Cell(ShName_PartList, CN_PL_ZagSumCount, RowN_PL_First - 1, "Кол-во необх.заготовки")
        set_Value_From_Cell(ShName_PartList, CN_PL_Sort, RowN_PL_First - 1, "Сортамент")
        set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, RowN_PL_First - 1, "Ключ IMBASE")
        Dim CN_PL_last_ColNum As Integer = Get_Last_Column(ShName_PartList, RowN_PL_First - 1)
        SetCellsBorderLineStyle2(ShName_PartList, 1, 1, CN_PL_last_ColNum, RowN_PL_First - 1, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)

        'зафиксировать шапку
        SetTitleFixirovano(ShName_PartList, CN_PL_Oboz, RowN_PL_First - 1, True)
    End Sub
    'ExcelSheets column numbers ПЕРВЫЙ ЛИСТ(СОСТАВ ИЗДЕЛИЯ(например))
    Public Sheetname As String = "СОСТАВ ИЗДЕЛИЯ"
    Public RowN_First As Integer = 3
    Public last_ColNum As Integer
    Public last_Rowmun As Integer
    Public CN_NumPP As Integer = 1
    Public CN_ArtID As Integer = 2
    Public CN_ArchName As Integer = 3
    Public CN_Oboz As Integer = 4
    Public CN_Naim As Integer = 5
    Public CN_Kolvo As Integer = 6
    Public CN_KolvoSbor As Integer = 7
    Public CN_EdIzm As Integer = 8
    Public CN_Mater As Integer = 9
    Public CN_MaterIBKey As Integer = 10
    Public CN_Mass As Integer = 11
    Public CN_Mass_MUIDStr As Integer = 12
    Public CN_Prim As Integer = 13
    Public CN_Type_Article As Integer = 14
    Public CN_Purchated As Integer = 15
    Public CN_IBKey_purchated As Integer = 16
    Public CN_TypeLink As Integer = 17
    Public CN_ContexrtType As Integer = 18
    Public CN_TPArtKey As Integer = 19
    Public CN_Raszehovka As Integer = 20
    Public CN_MatZagotovki As Integer = 21
    Public CN_MatZagotovki_IBKey As Integer = 22
    Public CN_RazmerZag As Integer = 23
    Public CN_ZagCount As Integer = 24
    Public CN_ZagSumCount As Integer = 25
    Public CN_NormaRashoda As Integer = 26
    Public CN_NormaRashoda_Sum As Integer = 27
    Public CN_NormaRashoda_MU As Integer = 28
    'Public CN_KIM As Integer = 25
    Public CN_VspomMat As Integer = 28
    Public CN_VspomMat_IBKey As Integer = 8

    Sub WriteEXCellTitle(tmp_node As TreeNode)
        Application.DoEvents()
        Add_NewSheet(Sheetname)
        Sheet_Delete_by_SheetName("Лист1")
        Dim MainTPInfo, MainArtInfo As Array
        Try
            Dim ArtID, PRJLINK_ID As Integer
            Dim Positio, Count_Summ As String
            If tmp_node.Tag.ToString.IndexOf(Spletter) > 0 Then
                Dim param_Array As Array = tmp_node.Tag.ToString.Split(Spletter)
                'MsgBox(param_Array(0) & Spletter & param_Array(1) & Spletter & param_Array(2))
                Positio = param_Array(0)
                ArtID = param_Array(1)
                PRJLINK_ID = param_Array(2)
                Count_Summ = param_Array(3)
            Else
                ArtID = tmp_node.Tag
            End If
            MainArtInfo = Get_Short_Article_Param(ArtID)

            If TPsfilter Then
                MainTPInfo = Get_TP_ParmArray(artID_Main, -1)
            End If
            set_Value_From_Cell(Sheetname, CN_ArtID, 1, artID_Main)

            set_Value_From_Cell(Sheetname, CN_Oboz, 1, MainArtInfo(0)) ' & " - " & MainArtInfo(1))
            set_Value_From_Cell(Sheetname, CN_Naim, 1, MainArtInfo(1))
            set_Value_From_Cell(Sheetname, CN_Mass, 1, MainArtInfo(2))
            set_Value_From_Cell(Sheetname, CN_ArchName, 1, MainArtInfo(4))
            set_Value_From_Cell(Sheetname, CN_ContexrtType, 1, ToolStripComboBox1.Text)
        Catch ex As Exception

        End Try
        Try
            set_Value_From_Cell(Sheetname, CN_Raszehovka, 1, MainTPInfo(0))
            'set_Value_From_Cell(Sheetname, CN_VspomMat, 1, MainTPInfo(9))
        Catch ex As Exception

        End Try

        set_Value_From_Cell(Sheetname, CN_NumPP, RowN_First - 1, "№" & vbNewLine & "п/п")
        set_Value_From_Cell(Sheetname, CN_ArtID, RowN_First - 1, "ArtID")
        set_Value_From_Cell(Sheetname, CN_ArchName, RowN_First - 1, "Архив")
        set_Value_From_Cell(Sheetname, CN_Oboz, RowN_First - 1, "Обозначение ДСЕ")
        set_Value_From_Cell(Sheetname, CN_Naim, RowN_First - 1, "Наименование ДСЕ")
        set_Value_From_Cell(Sheetname, CN_Kolvo, RowN_First - 1, "Кол." & vbNewLine & "на" & vbNewLine & "ед.")
        set_Value_From_Cell(Sheetname, CN_KolvoSbor, RowN_First - 1, "Кол." & vbNewLine & "на" & vbNewLine & "сбор.")
        set_Value_From_Cell(Sheetname, CN_EdIzm, RowN_First - 1, "ед.изм")
        set_Value_From_Cell(Sheetname, CN_Mater, RowN_First - 1, "Материал" & vbNewLine & "(Конструкторский)")
        set_Value_From_Cell(Sheetname, CN_MaterIBKey, RowN_First - 1, "Ключ IMBase" & vbNewLine & "для материала")
        set_Value_From_Cell(Sheetname, CN_Mass, RowN_First - 1, "Масса")
        set_Value_From_Cell(Sheetname, CN_Mass_MUIDStr, RowN_First - 1, "ед.изм" & vbNewLine & "массы")
        set_Value_From_Cell(Sheetname, CN_Prim, RowN_First - 1, "Прим.")
        set_Value_From_Cell(Sheetname, CN_Type_Article, RowN_First - 1, "Тип Объекта")
        set_Value_From_Cell(Sheetname, CN_Purchated, RowN_First - 1, "Признак" & vbNewLine & "изготовления")
        set_Value_From_Cell(Sheetname, CN_IBKey_purchated, RowN_First - 1, "Ключ IMBase" & vbNewLine & "для покупного")
        set_Value_From_Cell(Sheetname, CN_TypeLink, RowN_First - 1, "Тип связи")
        set_Value_From_Cell(Sheetname, CN_ContexrtType, RowN_First - 1, "Контекст связи")
        set_Value_From_Cell(Sheetname, CN_TPArtKey, RowN_First - 1, "TCZagKey")
        set_Value_From_Cell(Sheetname, CN_Raszehovka, RowN_First - 1, "Расцеховка")
        set_Value_From_Cell(Sheetname, CN_MatZagotovki, RowN_First - 1, "Сортамент")
        set_Value_From_Cell(Sheetname, CN_MatZagotovki_IBKey, RowN_First - 1, "Ключ IMBase" & vbNewLine & "для заготовки")
        set_Value_From_Cell(Sheetname, CN_RazmerZag, RowN_First - 1, "Размер заготовки")
        'set_Value_From_Cell(Sheetname, CN_KIM, RowN_First - 1, "КИМ")
        set_Value_From_Cell(Sheetname, CN_NormaRashoda, RowN_First - 1, "Норма расхода" & vbNewLine & "на ед.")
        set_Value_From_Cell(Sheetname, CN_NormaRashoda_Sum, RowN_First - 1, "Норма расхода" & vbNewLine & "общ.")
        set_Value_From_Cell(Sheetname, CN_NormaRashoda_MU, RowN_First - 1, "ед.изм")
        set_Value_From_Cell(Sheetname, CN_ZagCount, RowN_First - 1, "Кол-во деталей" & vbNewLine & "из заготовки")
        set_Value_From_Cell(Sheetname, CN_ZagSumCount, RowN_First - 1, "Кол-во необх." & vbNewLine & "заготовки")
        'set_Value_From_Cell(Sheetname, CN_VspomMat, RowN_First - 1, "Вспомогательный материал" & vbNewLine & "^&Ключ IMBase")
        'set_Value_From_Cell(Sheetname, CN_VspomMat_IBKey, RowN_First - 1, "Ключ IMBase" & vbNewLine & "Вспомогательный материал")

        Set_RowTextHorisontAllign(Sheetname, RowN_First - 1, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter)
        Set_RowVerticalAllign(Sheetname, RowN_First - 1, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter)

        last_ColNum = Get_Last_Column(Sheetname, RowN_First - 1)

        'выделить первую строку зеленым цветом
        SetCellsColor(Sheetname, 1, 1, last_ColNum, 1, Color.LawnGreen)
        'выделить красным, если головная специя лежит в архиве РАЗАРБОТКА
        Try
            If MainArtInfo(3) = "7" Or MainArtInfo(3) Is Nothing Then
                SetCellsColor(Sheetname, CN_ArchName, 1, CN_ArchName, 1, Color.Red)
            End If
        Catch ex As Exception

        End Try
        CellsTextBold(Sheetname, 1, 1, last_ColNum, 1)

        SetCellsBorderLineStyle2(Sheetname, 1, 1, last_ColNum, RowN_First - 1, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
        SetCellsTextSize(Sheetname, 1, 1, last_ColNum, 1, 14)
        'высота строки
        Set_RowHeigt(Sheetname, RowN_First - 1, 45)
        'зафиксировать шапку
        SetTitleFixirovano(Sheetname, 1, RowN_First - 1, True)
        'автоширина
        SetAutoFIT(Sheetname)
    End Sub
    Function PBarLoadSQL(MaxRow As Integer)
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Visible = True
        ToolStripProgressBar1.Maximum = MaxRow - 1
        ToolStripProgressBar1.Step = 1

    End Function

    Function PBarLoad(tmp_node As TreeNode)
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Visible = True
        ToolStripProgressBar1.Maximum = tmp_node.GetNodeCount(True) - 1
        ToolStripProgressBar1.Step = 1

    End Function
    Function PBarStep()
        ToolStripProgressBar1.PerformStep()
        If ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum Then
            ToolStripProgressBar1.Visible = False
        End If
    End Function
    Sub ReplaceCharArray()
        'собирает данные о заменяемых символах в специальный массив
        Dim objStreamReader As System.IO.StreamReader
        Dim strLine As String
        objStreamReader = New System.IO.StreamReader(CharOption.Char_opt_path)
        strLine = objStreamReader.ReadLine
        Dim w As Integer = 0
        ReDim Preserve ReplaceChar(w)
        Do While Not strLine Is Nothing
            ReDim Preserve ReplaceChar(w)
            ReplaceChar(w) = strLine
            strLine = objStreamReader.ReadLine
            w += 1
        Loop
        objStreamReader.Close()
    End Sub
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If TreeView1.Nodes.Count = 0 Then
            MsgBox("Для выгрузки Ведомости, выберите Объект")
            Exit Sub
        End If
        ReplaceCharArray()
        add_Vedomost(TreeView1.Nodes.Item(0))
    End Sub
    Sub add_Vedomost(tmp_node As TreeNode)
        Application.DoEvents()
        PBarLoad(tmp_node)

        Create_EX_Doc(NO_ExlProc_Visible)
        If NO_Parts Then addExTitlePartList(tmp_node) 'создать ВЕДОМОСТЬ ДЕТАЛЕЙ
        If NO_Purchated_PDRBN Then addExTitlePDRBN(tmp_node) 'создать ПОДРОБНУЮ ВЕДОМОСТЬ ПОКУПНЫХ
        If NO_Purchated Then addExTitlePurchated(tmp_node) 'создать ВЕДОМОСТЬ ПОКУПНЫХ
        If NO_Sostav Then WriteEXCellTitle(tmp_node) 'создать СОСТАВ ИЗДЕЛИЯ

        read_NEXTLEVELtreeview(tmp_node)
        If NO_Sostav Then
            last_Rowmun = Get_LastRowInOneColumn(Sheetname, CN_Naim)
            'ровнение текста в шапке
            CellsTextHorisAligment(Sheetname, 1, 1, last_ColNum, RowN_First - 1, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter)
            CellsTextVerticalAligment(Sheetname, 1, 1, last_ColNum, RowN_First - 1, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter)

            CellsTextHorisAligment(Sheetname, CN_NumPP, RowN_First, last_ColNum, last_Rowmun, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft)
        End If
        'пояснение к цветовой гамме
        If NO_ColorNote Then ColorNote()

        If NO_Sostav Then SetAutoFIT(Sheetname)
        If NO_Purchated Then SetAutoFIT(ShName_Purchated)
        If NO_Parts Then SetAutoFIT(ShName_PartList)
        If NO_Purchated_PDRBN Then SetAutoFIT(ShName_pdrbn)

        If Not NO_ExlProc_Visible Then EX_Doc_VisibleChanche(Not NO_ExlProc_Visible)
        If Not NO_S4_Columns Then ChancheVisibleColumn()
        'создание свойства
        CustomDocProp()
        sheet_sort()
        'SetCellsCdbl(Sheetname, CN_NumPP, RowN_First, CN_NumPP, last_Rowmun) 'преобразовал в число столбцы с позицией
        'SetCellsCdbl(Sheetname, CN_Mass, RowN_First, CN_Mass, last_Rowmun) 'преобразовал в число столбцы с массой 
        PBarStep()
        Beep()
    End Sub
    Sub sheet_sort()
        Try
            SortSheets_by_SheetName(ShName_Purchated, Sheetname)
            SortSheets_by_SheetName(ShName_pdrbn, ShName_Purchated)
            SortSheets_by_SheetName(ShName_PartList, ShName_pdrbn)
            Sheets_Activate_by_SheetName(Sheetname)
        Catch ex As Exception

        End Try
    End Sub
    Sub ChancheVisibleColumn()
        'настройки видимости колонок для листа СОСТАВ ИЗДЕЛИЯ
        SetColumnVisible(Sheetname, CN_ArtID, True)
        SetColumnVisible(Sheetname, CN_ArchName, True)
        SetColumnVisible(Sheetname, CN_MaterIBKey, True)
        SetColumnVisible(Sheetname, CN_Prim, True)
        SetColumnVisible(Sheetname, CN_Type_Article, True)
        SetColumnVisible(Sheetname, CN_IBKey_purchated, True)
        SetColumnVisible(Sheetname, CN_TypeLink, True)
        SetColumnVisible(Sheetname, CN_TPArtKey, True)
        SetColumnVisible(Sheetname, CN_ContexrtType, True)
        SetColumnVisible(Sheetname, CN_PL_Sort_IBKey, True)

        'настройки видимости колонок для листа ПЕРЕЧЕНЬ ПОКУПНЫХ
        SetColumnVisible(ShName_Purchated, CN_Purchated_IBKey, True)

        'настройки видимости колонок для листа ПЕРЕЧЕНЬ ДЕТАЛЕЙ
        SetColumnVisible(ShName_PartList, CN_PL_ArtID, True)
        SetColumnVisible(ShName_PartList, CN_PL_PROJ_ID, True)
        SetColumnVisible(ShName_PartList, CN_PL_LinkType, True)
        SetColumnVisible(ShName_PartList, CN_PL_Sort_IBKey, True)

        'настройки видимости колонок для листа ПОДРОБНЫЙ ПЕРЕЧЕНЬ ПОКУПНЫХ
        SetColumnVisible(ShName_pdrbn, CN_pdrbn_IBKey, True)
        SetColumnVisible(ShName_pdrbn, CN_pdrbn_prjid, True)
    End Sub
    Sub CustomDocProp()
        'Сведения о продукте(программе)
        CreateExcelCustomProperty("Программа", Application.ProductName, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
        CreateExcelCustomProperty("Версия", Application.ProductVersion, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
        CreateExcelCustomProperty("Дата создания", Now, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
        CreateExcelCustomProperty("Пользователь", Form1.ToolStripTextBox1.Text, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
        CreateExcelCustomProperty("Рабочая станция", Environment.MachineName, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
        'сведения о документе(объекте)
        CreateExcelCustomProperty("ArtID", artID_Main, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
        Dim t_DocID As Integer = get_DocIDbyArtId(artID_Main)

        CreateExcelCustomProperty("DocID", t_DocID, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
        CreateExcelCustomProperty("Версия документа", get_VersIDbyDocID(t_DocID), Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, 0)
    End Sub
    Sub ColorNote()
        Dim row_num As Integer
        Dim col_num As Integer
        If NO_Sostav Then
            row_num = Get_LastRowInOneColumn(Sheetname, CN_Naim) + 1 ' RowN_First
            col_num = last_ColNum + 3
            'пояснение цветовой гаммы для листа СОСТАВ ИЗДЕЛИЯ
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Пояснение к цв.схеме(заливке)", 1, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, DocumColor_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Документация", 0, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, AssemblyColor_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Сборочные единицы", 0, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, PartColor_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Детали", 0, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, StandartColor_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Стандартные изделия", 0, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, Pro4eeColor_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Прочие изделия", 0, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, MaterialColor_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Материалы", 0, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, KomplekTColor_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Комплекты", 0, 1, 0)
            row_num += 1

            SetCellsColor(Sheetname, col_num, row_num, col_num, row_num, TechContex_Sostav)
            set_Value_From_Cell_with_Proerty(Sheetname, col_num, row_num, "Технологич.связь", 0, 1, 0)
        End If

        'пояснение цветовой гаммы для листа ПЕРЕЧЕНЬ ПОКУПНЫХ
        If NO_Purchated Then
            row_num = Get_LastRowInOneColumn(ShName_Purchated, CN_Purchated_Naim) + 1 ' RowN_Purchated_First
            col_num = CN_Purchated_MU + 3

            set_Value_From_Cell_with_Proerty(ShName_Purchated, col_num, row_num, "Пояснение к цв.схеме(заливке)", 1, 1, 0)
            row_num += 1

            SetCellsColor(ShName_Purchated, col_num, row_num, col_num, row_num, VspomMaterialColor_Purchated)
            set_Value_From_Cell_with_Proerty(ShName_Purchated, col_num, row_num, "Материал вспомогательный", 0, 1, 0)
            row_num += 1

            SetCellsColor(ShName_Purchated, col_num, row_num, col_num, row_num, SortamentColor_Purchated)
            set_Value_From_Cell_with_Proerty(ShName_Purchated, col_num, row_num, "Сортамент изделия", 0, 1, 0)
            row_num += 1

            SetCellsColor(ShName_Purchated, col_num, row_num, col_num, row_num, SortamentColorKonstructor_Purchated)
            set_Value_From_Cell_with_Proerty(ShName_Purchated, col_num, row_num, "Материал изделия(Констр)", 0, 1, 0)
            row_num += 1

            SetCellsColor(ShName_Purchated, col_num, row_num, col_num, row_num, StandartColor_Purchated)
            set_Value_From_Cell_with_Proerty(ShName_Purchated, col_num, row_num, "Стандартные изделия", 0, 1, 0)
            row_num += 1

            SetCellsColor(ShName_Purchated, col_num, row_num, col_num, row_num, Pro4eeColor_Purchated)
            set_Value_From_Cell_with_Proerty(ShName_Purchated, col_num, row_num, "Прочие изделия", 0, 1, 0)
            row_num += 1

            SetCellsColor(ShName_Purchated, col_num, row_num, col_num, row_num, MaterialColor_Purchated)
            set_Value_From_Cell_with_Proerty(ShName_Purchated, col_num, row_num, "Материал", 0, 1, 0)
        End If
    End Sub
    Private TmpNode As TreeNode
    Sub read_NEXTLEVELtreeview(myNextNode As TreeNode)
        Application.DoEvents()
        Dim myNode As TreeNode
        For Each myNode In myNextNode.Nodes
            TmpNode = myNode
            Dim param_Array As Array = myNode.Tag.ToString.Split(Spletter)
            'MsgBox(param_Array(0) & Spletter & param_Array(1) & Spletter & param_Array(2))
            Dim ArtID, PRJLINK_ID As Integer
            Dim Positio, Count_Summ As String
            Positio = param_Array(0)
            ArtID = param_Array(1)
            PRJLINK_ID = param_Array(2)
            Count_Summ = param_Array(3)

            Dim PRJLINK_Param, Art_Param, TP_Array, TP_Vspom_Mater_Array As Array
            'Dim TP_VspMat_Array As Array
            PRJLINK_Param = Get_Link_Param(PRJLINK_ID)
            Art_Param = Get_Article_Param(ArtID)
            'TP_VspMat_Array = TP_VspMat_Array_func(ArtID)
            If TPsfilter Then
                TP_Array = Get_TP_ParmArray(ArtID, PRJLINK_Param(0))
                If Not OnlyFirstLewvel Then
                    TP_Vspom_Mater_Array = Get_Vspom_Mater_Array(ArtID, PRJLINK_ID)
                    Try
                        If TP_Vspom_Mater_Array(0, 1) IsNot Nothing Then excel_write_aboutVSPomMater_in_Purchated(TP_Vspom_Mater_Array, Count_Summ) 'And (NO_Purchated And Not OnlyFirstLewvel And PRJLINK_Param(4) <> 3 And myNextNode IsNot TreeView1.Nodes.Item(0))
                    Catch ex As Exception
                    End Try
                    Try
                        If TP_Vspom_Mater_Array(0, 1) IsNot Nothing Then excel_write_aboutVSPomMater_in_Purchated_PDRBN(TP_Vspom_Mater_Array, Count_Summ, ArtID) ' And (NO_Purchated_PDRBN And Not OnlyFirstLewvel And PRJLINK_Param(4) <> 3 And myNextNode IsNot TreeView1.Nodes.Item(0)) 
                    Catch ex As Exception
                    End Try
                End If
            End If
            PBarStep()

            If myNode.Nodes.Count > 0 Then
                If NO_Sostav Then excel_write_about_TreeNode(Positio, ArtID, PRJLINK_ID, Count_Summ, PRJLINK_Param, Art_Param, TP_Array)
                read_NEXTLEVELtreeview(myNode)
            Else
                If NO_Sostav Then excel_write_about_TreeNode(Positio, ArtID, PRJLINK_ID, Count_Summ, PRJLINK_Param, Art_Param, TP_Array)
                If NO_Purchated Then excel_write_aboutPart_in_Purchated(ArtID, Art_Param, TP_Array, PRJLINK_Param, Count_Summ)
                If NO_Parts Then excel_write_aboutPart_in_PartList(ArtID, PRJLINK_ID, Art_Param, TP_Array, PRJLINK_Param, Count_Summ)
                If NO_Purchated_PDRBN Then excel_write_about_PDRBN_Purchated(ArtID, PRJLINK_ID, Art_Param, TP_Array, PRJLINK_Param, Count_Summ)
            End If
        Next
    End Sub
    Sub excel_write_about_PDRBN_Purchated(ArtID As Integer, PRJLINK_ID As Integer, Art_Info As Array, TC_Info As Array, PRJLINK_Param As Array, Total_Count As String) '(ArtID As Integer, Art_Info As Array, TC_Info As Array, PRJLINK_Param As Array, Total_Count As String)
        Application.DoEvents()
        Dim tmp_colors As Color = Color.Khaki
        Dim lastRowNumPurchated As Integer = Get_LastRowInOneColumn(ShName_pdrbn, CN_pdrbn_Count) + 1
        Dim tmp_row As Integer
        Dim sum, totalsum As Double

        Select Case Art_Info(12)
            Case 4 'детали
                Try
                    If TC_Info(4) IsNot Nothing Then
                        tmp_colors = SortamentColor_Purchated '  Color.Tan
                        tmp_row = get_value_bay_FindText_Strong(ShName_pdrbn, CN_pdrbn_IBKey, RowN_pdrbn_First, TC_Info(4)) + 1

                        If tmp_row > 0 Then
                            'создание новой строки
                            add_NewRow(ShName_pdrbn, tmp_row)
                            'объединить ячейки из созданной и предыдущей строки
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_Naim, tmp_row - 1, CN_pdrbn_Naim, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка наименование
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_IBKey, tmp_row - 1, CN_pdrbn_IBKey, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка IBKey
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, CN_pdrbn_TotalCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Общее колво 
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row - 1, CN_pdrbn_ObjType, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Тип объекта
                            'дописать в новой строчке информацию о новой детали

                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, tmp_row, ArtID) 'ArtID применяемого объекта (деталь)
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, tmp_row, Art_Info(0) & " (" & Art_Info(1) & ")") 'Колонка обозначение применяемого обекта
                            sum = Math.Round((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + CDbl(Total_Count) * CDbl(TC_Info(2))), 3) 'CDbl(Total_Count) / CDbl(TC_Info(6))
                            'sum = (CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + CDbl(Total_Count))
                            'sum = Math.Round((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count) / CDbl(TC_Info(6)) * CDbl(TC_Info(2))), 3)
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, sum.ToString.Replace(",", ".")) 'Общ кол-во/масса
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, tmp_row, (CDbl(Total_Count) * CDbl(TC_Info(2))).ToString.Replace(",", ".")) 'Кол-во/масса
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, tmp_row, TC_Info(10)) 'Ед.изм. 
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row, "Сортамент изделия") 'тип объекта
                        Else
                            set_Value_From_Cell(ShName_pdrbn, CN_Purchated_Naim, lastRowNumPurchated, TC_Info(3)) ' наименование
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_IBKey, lastRowNumPurchated, TC_Info(4))  ' ключ IMBASE
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, lastRowNumPurchated, ArtID) 'ArtID применяемого объекта (деталь)
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, lastRowNumPurchated, Art_Info(0) & " (" & Art_Info(1) & ")") 'обозначение применяемого обекта
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, lastRowNumPurchated, (CDbl(Total_Count) * CDbl(TC_Info(2))).ToString.Replace(",", ".")) 'Кол-во/масса

                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, lastRowNumPurchated, TC_Info(10)) 'ед изм
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated, ((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count) * CDbl(TC_Info(2))).ToString.Replace(",", "."))) 'Общ кол-во/масса  CDbl(Total_Count) / CDbl(TC_Info(6))
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, lastRowNumPurchated, "Сортамент изделия") 'тип объекта
                        End If
                    Else
                        tmp_colors = SortamentColorKonstructor_Purchated ' Color.LightSteelBlue
                        tmp_row = get_value_bay_FindText_Strong(ShName_pdrbn, CN_pdrbn_IBKey, RowN_pdrbn_First, Art_Info(6)) + 1
                        'If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                        'set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Naim, lastRowNumPurchated, Art_Info(5))
                        'set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_IBKey, lastRowNumPurchated, Art_Info(6))
                        'set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, lastRowNumPurchated, Art_Info(9))

                        'sum = Math.Round((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count) * CDbl(Art_Info(2))), 3)
                        'set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated, sum.ToString.Replace(",", "."))

                        If tmp_row > 0 Then
                            'создание новой строки
                            add_NewRow(ShName_pdrbn, tmp_row)
                            'объединить ячейки из созданной и предыдущей строки
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_Naim, tmp_row - 1, CN_pdrbn_Naim, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка наименование
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_IBKey, tmp_row - 1, CN_pdrbn_IBKey, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка IBKey
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, CN_pdrbn_TotalCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Общее колво 
                            CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row - 1, CN_pdrbn_ObjType, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Тип объекта
                            'дописать в новой строчке информацию о новой детали

                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, tmp_row, ArtID) 'ArtID применяемого объекта (деталь)
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, tmp_row, Art_Info(0) & " (" & Art_Info(1) & ")") 'Колонка обозначение применяемого обекта
                            sum = Math.Round((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + CDbl(Total_Count) * CDbl(Art_Info(2))), 3)
                            'sum = (CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + CDbl(Total_Count))
                            'sum = Math.Round((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count) / CDbl(TC_Info(6)) * CDbl(TC_Info(2))), 3)
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, sum.ToString.Replace(",", ".")) 'Общ кол-во/масса
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, tmp_row, (CDbl(Total_Count) * CDbl(Art_Info(2))).ToString.Replace(",", ".")) 'Кол-во/масса
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, tmp_row, Art_Info(9)) 'Ед.изм. 
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row, "Материал изделия(Констр)") 'тип объекта
                        Else
                            set_Value_From_Cell(ShName_pdrbn, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(5)) ' наименование
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_IBKey, lastRowNumPurchated, Art_Info(6))  ' ключ IMBASE
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, lastRowNumPurchated, ArtID) 'ArtID применяемого объекта (деталь)
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, lastRowNumPurchated, Art_Info(0) & " (" & Art_Info(1) & ")") 'обозначение применяемого обекта
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, lastRowNumPurchated, (CDbl(Total_Count) * CDbl(Art_Info(2))).ToString.Replace(",", ".")) 'Кол-во/масса

                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, lastRowNumPurchated, Art_Info(9)) 'ед изм
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated, (CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count) * CDbl(Art_Info(2))).ToString.Replace(",", ".")) 'Общ кол-во/масса 
                            set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, lastRowNumPurchated, "Материал изделия(Констр)") 'тип объекта
                        End If

                    End If
                Catch ex As Exception

                End Try
                'get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(0))1
            Case 5, 6 'станд изделия
                Try
                    If Art_Info(12) = 5 Then
                        tmp_colors = StandartColor_Purchated  ' Color.LightGreen
                    Else
                        tmp_colors = Pro4eeColor_Purchated
                    End If
                    tmp_row = get_value_bay_FindText_Strong(ShName_pdrbn, CN_pdrbn_IBKey, RowN_Purchated_First, Art_Info(3)) + 1
                    If tmp_row > 0 Then
                        'создание новой строки
                        add_NewRow(ShName_pdrbn, tmp_row)
                        'объединить ячейки из созданной и предыдущей строки
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_Naim, tmp_row - 1, CN_pdrbn_Naim, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка наименование
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_IBKey, tmp_row - 1, CN_pdrbn_IBKey, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка IBKey
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, CN_pdrbn_TotalCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Общее колво 
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row - 1, CN_pdrbn_ObjType, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Тип объекта
                        'дописать в новой строчке информацию о новой детали

                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, tmp_row, PRJLINK_Param(0)) 'ArtID применяемого объекта 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, tmp_row, get_OBOZ_by_Part_AID(PRJLINK_Param(0)) & " (" & get_Naim_by_Part_AID(PRJLINK_Param(0)) & ")") 'Колонка обозначение применяемого обекта
                        sum = (CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + CDbl(Total_Count))
                        'sum = Math.Round((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count) / CDbl(TC_Info(6)) * CDbl(TC_Info(2))), 3)
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, sum.ToString.Replace(",", ".")) 'Общ кол-во/масса
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, tmp_row, CDbl(Total_Count)) 'Кол-во/масса
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, tmp_row, PRJLINK_Param(11)) 'Ед.изм. 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row, Art_Info(7)) 'тип объекта
                    Else
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Naim, lastRowNumPurchated, Art_Info(1)) ' наименование
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_IBKey, lastRowNumPurchated, Art_Info(3))  ' ключ IMBASE
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, lastRowNumPurchated, PRJLINK_Param(0)) 'ArtID применяемого объекта 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, lastRowNumPurchated, get_OBOZ_by_Part_AID(PRJLINK_Param(0)) & " (" & get_Naim_by_Part_AID(PRJLINK_Param(0)) & ")") 'обозначение применяемого обекта
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, lastRowNumPurchated, CDbl(Total_Count)) 'Кол-во/масса

                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, lastRowNumPurchated, PRJLINK_Param(11)) 'ед изм
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated, CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count)) 'Общ кол-во/масса 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, lastRowNumPurchated, Art_Info(7)) 'тип объекта
                    End If
                Catch ex As Exception

                End Try
                'Case 6 'прочие            Case 7 'материалы
            Case 7 'материалы
                Try
                    tmp_colors = MaterialColor_Purchated ' Color.LightSkyBlue
                    tmp_row = get_value_bay_FindText_Strong(ShName_pdrbn, CN_pdrbn_IBKey, RowN_pdrbn_First, Art_Info(3)) + 1

                    If tmp_row > 0 Then
                        'создание новой строки
                        add_NewRow(ShName_pdrbn, tmp_row)
                        'объединить ячейки из созданной и предыдущей строки
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_Naim, tmp_row - 1, CN_pdrbn_Naim, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка наименование
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_IBKey, tmp_row - 1, CN_pdrbn_IBKey, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка IBKey
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, CN_pdrbn_TotalCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Общее колво 
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row - 1, CN_pdrbn_ObjType, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Тип объекта
                        'дописать в новой строчке информацию о новой детали

                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, tmp_row, PRJLINK_Param(0)) 'ArtID применяемого объекта 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, tmp_row, get_OBOZ_by_Part_AID(PRJLINK_Param(0)) & " (" & get_Naim_by_Part_AID(PRJLINK_Param(0)) & ")") 'Колонка обозначение применяемого обекта 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, tmp_row, CDbl(Total_Count)) 'Кол-во/масса
                        If CDbl(TC_Info(6)) <> 0 Then
                            sum = CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + (Total_Count) * CDbl(TC_Info(2))
                        Else
                            sum = (CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + CDbl(Total_Count))
                        End If
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, sum.ToString.Replace(",", ".")) 'Общ кол-во/масса
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, tmp_row, PRJLINK_Param(11)) 'Ед.изм.  
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row, Art_Info(7)) 'тип объекта
                    Else
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Naim, lastRowNumPurchated, Art_Info(1)) ' наименование
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_IBKey, lastRowNumPurchated, Art_Info(3))  ' ключ IMBASE
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, tmp_row, PRJLINK_Param(0)) 'ArtID применяемого объекта 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, lastRowNumPurchated, get_OBOZ_by_Part_AID(PRJLINK_Param(0)) & " (" & get_Naim_by_Part_AID(PRJLINK_Param(0)) & ")") 'обозначение применяемого обекта
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, lastRowNumPurchated, CDbl(Total_Count)) 'Кол-во/масса 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, lastRowNumPurchated, PRJLINK_Param(11)) 'ед изм

                        If CDbl(TC_Info(6)) <> 0 Then
                            sum = CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + (Total_Count / CDbl(TC_Info(6))) * CDbl(TC_Info(2))
                        Else
                            sum = (CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + CDbl(Total_Count))
                        End If
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated, CDbl(sum)) 'Общ кол-во/масса 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, lastRowNumPurchated, Art_Info(7)) 'тип объекта
                    End If

                Catch ex As Exception

                End Try
        End Select
        If tmp_row > 0 Then lastRowNumPurchated = tmp_row
        'красим в свой цвет
        SetCellsColor(ShName_pdrbn, CN_pdrbn_Naim, lastRowNumPurchated, CN_pdrbn_TotalCount, lastRowNumPurchated, tmp_colors)
        SetCellsBorderLineStyle2(ShName_pdrbn, CN_pdrbn_Naim, lastRowNumPurchated, CN_pdrbn_TotalCount, lastRowNumPurchated, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)

    End Sub
    Sub excel_write_aboutPart_in_PartList(ArtID As Integer, PRJLINK_ID As Integer, Art_Info As Array, TC_Info As Array, PRJLINK_Param As Array, Total_Count As String)
        Application.DoEvents()
        Dim sum, totalsum As Double
        Dim tmp_row As Integer
        Select Case Art_Info(12)
            Case 4
                Dim lastRowNumInPartlist As Integer = Get_LastRowInOneColumn(ShName_PartList, CN_PL_LinkType) + 1
                Try
                    tmp_row = get_value_bay_FindText_Strong(ShName_PartList, CN_PL_ArtID, RowN_PL_First, ArtID) + 1

                    If tmp_row > 0 Then
                        'создание новой строки
                        add_NewRow(ShName_PartList, tmp_row)
                        'объединить ячейки из созданной и предыдущей строки
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_ArtID, tmp_row - 1, CN_PL_ArtID, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка ArtID
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_Oboz, tmp_row - 1, CN_PL_Oboz, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка обозначение
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_Naim, tmp_row - 1, CN_PL_Naim, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка наименование
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_TotalCount, tmp_row - 1, CN_PL_TotalCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Общее колво
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_ZagCount, tmp_row - 1, CN_PL_ZagCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Общее колво
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_Sort, tmp_row - 1, CN_PL_Sort, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка сортамент
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_Sort_IBKey, tmp_row - 1, CN_PL_Sort_IBKey, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка ключ ImBase
                        CellsMergeWithTextAlignment(ShName_PartList, CN_PL_ZagSumCount, tmp_row - 1, CN_PL_ZagSumCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Кол-во необх.заготовки
                        'дописать в новой строчке информацию о новой детали
                        set_Value_From_Cell(ShName_PartList, CN_PL_PROJ_ID, tmp_row, PRJLINK_Param(0)) 'Колонка PROJ_ID
                        set_Value_From_Cell(ShName_PartList, CN_PL_Primen9emost, tmp_row, get_OBOZ_by_Part_AID(PRJLINK_Param(0))) 'обозначение применяемого обекта

                        set_Value_From_Cell(ShName_PartList, CN_PL_LinkType, tmp_row, PRJLINK_Param(9)) 'тип связи

                        If TC_Info(6) IsNot Nothing Then
                            sum = CDbl(Total_Count) / CDbl(TC_Info(6))
                            totalsum = CDbl(get_Value_From_Cell(ShName_PartList, CN_PL_TotalCount, tmp_row - 1)) + CDbl(Total_Count)
                        Else
                            sum = CDbl(PRJLINK_Param(2)) '* CDbl(Total_Count)
                            totalsum = CDbl(get_Value_From_Cell(ShName_PartList, CN_PL_TotalCount, tmp_row - 1)) + CDbl(Total_Count)
                        End If
                        set_Value_From_Cell(ShName_PartList, CN_PL_Count, tmp_row, CDbl(PRJLINK_Param(2))) ' sum) 'колво
                        set_Value_From_Cell(ShName_PartList, CN_PL_TotalCount, tmp_row - 1, totalsum) 'Общее колво


                        Try
                            Dim ZagSumCount As Double
                            If CDbl(TC_Info(6)) <> 0 Then
                                ZagSumCount = Math.Round(CDbl(get_Value_From_Cell(ShName_PartList, CN_PL_ZagSumCount, tmp_row - 1)) + Total_Count / CDbl(TC_Info(6)), 3)
                                set_Value_From_Cell(ShName_PartList, CN_PL_ZagSumCount, tmp_row - 1, ZagSumCount)
                                set_Value_From_Cell(ShName_PartList, CN_PL_ZagCount, tmp_row, TC_Info(6)) 'колво заготовок
                            Else
                                set_Value_From_Cell(ShName_PartList, CN_PL_ZagSumCount, tmp_row, 0)
                            End If
                        Catch ex As Exception

                        End Try
                    Else
                        set_Value_From_Cell(ShName_PartList, CN_PL_ArtID, lastRowNumInPartlist, ArtID)
                        set_Value_From_Cell(ShName_PartList, CN_PL_Oboz, lastRowNumInPartlist, Art_Info(0))
                        set_Value_From_Cell(ShName_PartList, CN_PL_Naim, lastRowNumInPartlist, Art_Info(1))
                        set_Value_From_Cell(ShName_PartList, CN_PL_PROJ_ID, lastRowNumInPartlist, PRJLINK_Param(0))
                        set_Value_From_Cell(ShName_PartList, CN_PL_Primen9emost, lastRowNumInPartlist, get_OBOZ_by_Part_AID(PRJLINK_Param(0))) 'обозначение применяемого обекта

                        'set_Value_From_Cell(ShName_PartList, CN_PL_MU, lastRowNumInPartlist, PRJLINK_Param(11)) 'ед изм
                        set_Value_From_Cell(ShName_PartList, CN_PL_LinkType, lastRowNumInPartlist, PRJLINK_Param(9)) 'тип связи

                        If TC_Info(6) IsNot Nothing Then
                            sum = CDbl(Total_Count) / CDbl(TC_Info(6))
                            totalsum = CDbl(Total_Count)
                        Else
                            sum = CDbl(PRJLINK_Param(2)) '* CDbl(Total_Count)
                            totalsum = CDbl(Total_Count)
                        End If
                        set_Value_From_Cell(ShName_PartList, CN_PL_Count, lastRowNumInPartlist, CDbl(PRJLINK_Param(2))) ' sum) 'колво
                        set_Value_From_Cell(ShName_PartList, CN_PL_TotalCount, lastRowNumInPartlist, totalsum) 'колво общее

                        If TC_Info(3) IsNot Nothing Or TC_Info(3) <> "" Then
                            set_Value_From_Cell(ShName_PartList, CN_PL_Sort, lastRowNumInPartlist, TC_Info(3)) 'Сортамент
                            set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, lastRowNumInPartlist, TC_Info(4)) 'ключ имбайз сортамента 
                        Else
                            set_Value_From_Cell(ShName_PartList, CN_PL_Sort, lastRowNumInPartlist, Art_Info(5)) 'Сортамент
                            set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, lastRowNumInPartlist, Art_Info(6)) 'ключ имбайз сортамента
                        End If
                        'колво заготовок
                        Try
                            Dim ZagSumCount As Double
                            If CDbl(TC_Info(6)) <> 0 Then
                                ZagSumCount = Math.Round(CDbl(get_Value_From_Cell(ShName_PartList, CN_PL_ZagSumCount, lastRowNumInPartlist)) + Total_Count / CDbl(TC_Info(6)), 3)
                                set_Value_From_Cell(ShName_PartList, CN_PL_ZagCount, lastRowNumInPartlist, TC_Info(6)) 'колво заготовок
                                set_Value_From_Cell(ShName_PartList, CN_PL_ZagSumCount, lastRowNumInPartlist, ZagSumCount)
                            Else
                                set_Value_From_Cell(ShName_PartList, CN_PL_ZagSumCount, lastRowNumInPartlist, 0)
                            End If
                        Catch ex As Exception

                        End Try
                    End If


                Catch ex As Exception

                End Try

                'Try

                '    set_Value_From_Cell(ShName_PartList, CN_PL_Sort, lastRowNumInPartlist, TC_Info(3))
                '    set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, lastRowNumInPartlist, TC_Info(4))
                'Catch ex As Exception
                '    set_Value_From_Cell(ShName_PartList, CN_PL_Sort, lastRowNumInPartlist, Art_Info(5))
                '    set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, lastRowNumInPartlist, Art_Info(6))
                'End Try
                Dim tmp_colors As Color
                If PRJLINK_Param(8) = 2 Then
                    tmp_colors = TechContex_Sostav
                Else
                    tmp_colors = PartColor_Sostav
                End If
                SetCellsBorderLineStyle2(ShName_PartList, 1, lastRowNumInPartlist, CN_PL_Sort_IBKey, lastRowNumInPartlist, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
                'красим в свой цвет
                SetCellsColor(ShName_PartList, 1, lastRowNumInPartlist, CN_PL_Sort_IBKey, lastRowNumInPartlist, tmp_colors)
        End Select
    End Sub
    Function get_OBOZ_by_Part_AID(Part_AID As Integer) As String
        Application.DoEvents()
        Dim oboz_str As String
        Try
            With s4
                .OpenArticle(Part_AID)
                oboz_str = .GetArticleDesignation
                .CloseArticle()
                Return oboz_str
            End With
        Catch ex As Exception

        End Try
    End Function
    Function get_Naim_by_Part_AID(Part_AID As Integer) As String
        Application.DoEvents()
        Dim Naim_str As String
        Try
            With s4
                .OpenArticle(Part_AID)
                Naim_str = .GetArticleName
                .CloseArticle()
                Return Naim_str
            End With
        Catch ex As Exception

        End Try
    End Function
    Function get_DocIDbyArtId(ArtID As Integer) As Integer
        Application.DoEvents()
        Dim DocID As Integer
        Try
            With s4
                .OpenArticle(ArtID)
                DocID = .GetArticleDocID
                .CloseArticle()
                Return DocID
            End With
        Catch ex As Exception
            Return -3
        End Try
    End Function
    Function get_VersIDbyDocID(DocID As Integer) As Integer
        Application.DoEvents()
        Dim Vers As Integer
        Try
            With s4
                .OpenDocument(DocID)
                Vers = .GetDocActualVersionID
                .CloseDocument()
                Return Vers
            End With
        Catch ex As Exception
            Return -3
        End Try
    End Function
    Sub excel_write_aboutPart_in_Purchated(ArtID As Integer, Art_Info As Array, TC_Info As Array, PRJLINK_Param As Array, Total_Count As String)
        Application.DoEvents()
        Dim tmp_colors As Color = Color.Khaki
        Dim lastRowNumPurchated As Integer = Get_LastRowInOneColumn(ShName_Purchated, CN_Purchated_Naim) + 1
        Dim tmp_row As Integer
        Dim sum As Double
        Select Case Art_Info(12)
            Case 4 'детали
                Try
                    If TC_Info(4) IsNot Nothing Then
                        tmp_colors = SortamentColor_Purchated '  Color.Tan
                        tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, TC_Info(4))
                        If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, TC_Info(3))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, TC_Info(4))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, TC_Info(10))

                        sum = Math.Round((CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(Total_Count) * CDbl(TC_Info(2))), 3) ' CDbl(Total_Count) / CDbl(TC_Info(6))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))
                    Else
                        tmp_colors = SortamentColorKonstructor_Purchated ' Color.LightSteelBlue
                        tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, Art_Info(6))
                        If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(5))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, Art_Info(6))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, Art_Info(9))

                        sum = Math.Round((CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(Total_Count) * CDbl(Art_Info(2))), 3)
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))
                    End If
                Catch ex As Exception

                End Try
                'get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(0))1
            Case 5, 6 'станд изделия
                Try
                    If Art_Info(12) = 5 Then
                        tmp_colors = StandartColor_Purchated  ' Color.LightGreen
                    Else
                        tmp_colors = Pro4eeColor_Purchated
                    End If
                    tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, Art_Info(3))
                    If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(1))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, Art_Info(3))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, PRJLINK_Param(11))

                    sum = (CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(Total_Count))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))
                    'красим вспомогательный материал в свой цвет
                    'SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)

                Catch ex As Exception

                End Try
                'Case 6 'прочие            Case 7 'материалы
            Case 7 'материалы
                Try
                    tmp_colors = MaterialColor_Purchated ' Color.LightSkyBlue
                    tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, Art_Info(3))
                    If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                    If TC_Info(3) IsNot Nothing Or TC_Info(3) <> "" Then
                        tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, TC_Info(4))
                        If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, TC_Info(3))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, TC_Info(4))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, TC_Info(10))
                    Else
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(1))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, Art_Info(3))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, PRJLINK_Param(11))
                    End If
                    If CDbl(TC_Info(6)) <> 0 Then
                        sum = CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + (Total_Count) * CDbl(TC_Info(2)) 'Total_Count / CDbl(TC_Info(6)))
                    Else
                        sum = (CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(Total_Count))
                    End If
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))
                    'красим вспомогательный материал в свой цвет
                    'SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)


                    'Dim Norma_ras_SUM As Double
                    'If CDbl(TP_Array(6)) <> 0 Then
                    '    ZagSumCount = Count_Summ / CDbl(TP_Array(6))
                    '    Norma_ras_SUM = ZagSumCount * CDbl(TP_Array(2))
                    '    set_Value_From_Cell(Sheetname, CN_ZagSumCount, lastRowNum, ZagSumCount)
                    '    set_Value_From_Cell(Sheetname, CN_NormaRashoda_Sum, lastRowNum, Norma_ras_SUM.ToString.Replace(",", "."))
                    'Else
                    '    set_Value_From_Cell(Sheetname, CN_ZagSumCount, lastRowNum, 0)
                    '    set_Value_From_Cell(Sheetname, CN_NormaRashoda_Sum, lastRowNum, 0)
                    'End If

                Catch ex As Exception

                End Try
                'Try
                '    tmp_colors = Color.LightSkyBlue
                '    tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, Art_Info(3))
                '    If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                '    If TC_Info(3) IsNot Nothing Or TC_Info(3) <> "" Then
                '        tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, TC_Info(4))
                '        If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                '        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, TC_Info(3))
                '        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, TC_Info(4))
                '        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, TC_Info(10))
                '    Else
                '        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(1))
                '        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, Art_Info(3))
                '        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, PRJLINK_Param(11))
                '    End If
                '    If CDbl(TC_Info(6)) <> 0 Then
                '        sum = CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + (Total_Count / CDbl(TC_Info(6))) * CDbl(TC_Info(2))
                '    Else
                '        sum = (CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(Total_Count))
                '    End If
                '    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))
                '    'красим вспомогательный материал в свой цвет
                '    SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)


                '    'Dim Norma_ras_SUM As Double
                '    'If CDbl(TP_Array(6)) <> 0 Then
                '    '    ZagSumCount = Count_Summ / CDbl(TP_Array(6))
                '    '    Norma_ras_SUM = ZagSumCount * CDbl(TP_Array(2))
                '    '    set_Value_From_Cell(Sheetname, CN_ZagSumCount, lastRowNum, ZagSumCount)
                '    '    set_Value_From_Cell(Sheetname, CN_NormaRashoda_Sum, lastRowNum, Norma_ras_SUM.ToString.Replace(",", "."))
                '    'Else
                '    '    set_Value_From_Cell(Sheetname, CN_ZagSumCount, lastRowNum, 0)
                '    '    set_Value_From_Cell(Sheetname, CN_NormaRashoda_Sum, lastRowNum, 0)
                '    'End If

                'Catch ex As Exception

                'End Try
        End Select
        'красим в свой цвет
        SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)
        SetCellsBorderLineStyle2(ShName_Purchated, 1, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)

    End Sub
    Sub excel_write_aboutVSPomMater_in_Purchated(TP_Vspom_Mater_Array As Array, Count_Sum As Double)
        Application.DoEvents()
        Dim lastRowNumPurchated As Integer = Get_LastRowInOneColumn(ShName_Purchated, CN_Purchated_Naim) + 1
        Dim tmp_row As Integer
        Dim sum As Double
        Dim tmp_colors As Color = VspomMaterialColor_Purchated  'Color.DarkSalmon
        Try
            If UBound(TP_Vspom_Mater_Array, 1) >= 0 Then
                For q As Integer = 0 To UBound(TP_Vspom_Mater_Array, 1)
                    lastRowNumPurchated = Get_LastRowInOneColumn(ShName_Purchated, CN_Purchated_Naim) + 1
                    tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, TP_Vspom_Mater_Array(q, 1))
                    If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 0))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 1))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 3))

                    sum = Math.Round((CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + Count_Sum * CDbl(TP_Vspom_Mater_Array(q, 2))), 3)
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))

                    'красим вспомогательный материал в свой цвет
                    SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)
                    SetCellsBorderLineStyle2(ShName_Purchated, 1, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub excel_write_aboutVSPomMater_in_Purchated_PDRBN(TP_Vspom_Mater_Array As Array, Count_Sum As Double, PRJ_ID As Integer)
        Application.DoEvents()
        Dim lastRowNumPurchated As Integer = Get_LastRowInOneColumn(ShName_pdrbn, CN_pdrbn_Primen9emost) + 1
        Dim tmp_row As Integer
        Dim sum As Double
        Dim tmp_colors As Color = VspomMaterialColor_Purchated  'Color.DarkSalmon
        Try
            If UBound(TP_Vspom_Mater_Array, 1) >= 0 Then
                For q As Integer = 0 To UBound(TP_Vspom_Mater_Array, 1)
                    lastRowNumPurchated = Get_LastRowInOneColumn(ShName_pdrbn, CN_pdrbn_Primen9emost) + 1
                    tmp_row = get_value_bay_FindText_Strong(ShName_pdrbn, CN_pdrbn_IBKey, RowN_pdrbn_First, TP_Vspom_Mater_Array(q, 1)) + 1
                    'If tmp_row > 0 Then lastRowNumPurchated = tmp_row

                    'set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 0))
                    'set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 1))
                    'set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 3))

                    'sum = Math.Round((CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + Count_Sum * CDbl(TP_Vspom_Mater_Array(q, 2))), 3)
                    'set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))

                    If tmp_row > 0 Then
                        'создание новой строки
                        add_NewRow(ShName_pdrbn, tmp_row)
                        'объединить ячейки из созданной и предыдущей строки
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_Naim, tmp_row - 1, CN_pdrbn_Naim, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка наименование
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_IBKey, tmp_row - 1, CN_pdrbn_IBKey, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка IBKey
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, CN_pdrbn_TotalCount, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Общее колво 
                        CellsMergeWithTextAlignment(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row - 1, CN_pdrbn_ObjType, tmp_row, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft) 'Колонка Тип объекта
                        'дописать в новой строчке информацию о новой детали

                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, tmp_row, PRJ_ID)
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, tmp_row, get_OBOZ_by_Part_AID(PRJ_ID) & " (" & get_Naim_by_Part_AID(PRJ_ID) & ")") 'Колонка обозначение применяемого обекта 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, tmp_row, (Count_Sum * CDbl(TP_Vspom_Mater_Array(q, 2))).ToString.Replace(",", ".")) 'Кол-во/масса 
                        sum = Math.Round((CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1)) + Count_Sum * CDbl(TP_Vspom_Mater_Array(q, 2))), 3)
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, tmp_row - 1, sum.ToString.Replace(",", ".")) 'Общ кол-во/масса
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, tmp_row, TP_Vspom_Mater_Array(q, 3)) 'Ед.изм.  
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, tmp_row, "Вспомогательный материал")
                    Else
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Naim, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 0)) ' наименование
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_IBKey, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 1))  ' ключ IMBASE
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_prjid, lastRowNumPurchated, PRJ_ID)
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Primen9emost, lastRowNumPurchated, get_OBOZ_by_Part_AID(PRJ_ID) & " (" & get_Naim_by_Part_AID(PRJ_ID) & ")") 'обозначение применяемого обекта
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_Count, lastRowNumPurchated, (Count_Sum * CDbl(TP_Vspom_Mater_Array(q, 2))).ToString.Replace(",", ".")) 'Кол-во/масса 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_MU, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 3)) 'ед изм

                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated, (CDbl(get_Value_From_Cell(ShName_pdrbn, CN_pdrbn_TotalCount, lastRowNumPurchated)) + Count_Sum * CDbl(TP_Vspom_Mater_Array(q, 2))).ToString.Replace(",", ".")) 'Общ кол-во/масса 
                        set_Value_From_Cell(ShName_pdrbn, CN_pdrbn_ObjType, lastRowNumPurchated, "Вспомогательный материал")
                    End If
                    'красим вспомогательный материал в свой цвет
                    SetCellsColor(ShName_pdrbn, CN_pdrbn_Naim, lastRowNumPurchated, CN_pdrbn_TotalCount, lastRowNumPurchated, tmp_colors)
                    SetCellsBorderLineStyle2(ShName_pdrbn, 1, lastRowNumPurchated, CN_pdrbn_TotalCount, lastRowNumPurchated, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Function ItsLastTreenode() As Boolean
        '0 если есть вложенности
        '1 есть нет вложенности
        If TmpNode.Nodes.Count > 0 Then
            Return 0
        Else
            Return 1
        End If
    End Function
    Sub excel_write_about_TreeNode(Positio As String, ArtID As Integer, PRJLINK_ID As Integer, Count_Summ As Double, PRJLINK_Param As Array, Art_Param As Array, TP_Array As Array)

        Application.DoEvents()
        'Dim PRJLINK_Param, Art_Param, TP_Array As Array
        ''Dim TP_VspMat_Array As Array
        Dim temp_color As Color
        'PRJLINK_Param = Get_Link_Param(PRJLINK_ID)
        'Art_Param = Get_Article_Param(ArtID)
        ''TP_VspMat_Array = TP_VspMat_Array_func(ArtID)
        'If TPsfilter Then
        '    TP_Array = Get_TP_ParmArray(ArtID)
        'End If
        Dim lastRowNum As Integer = Get_LastRowInOneColumn(Sheetname, CN_ContexrtType) + 1
        'If lastRowNum < Get_LastRowInOneColumn(Sheetname, last_ColNum) + 1 Then lastRowNum = Get_LastRowInOneColumn(Sheetname, last_ColNum) + 1
        set_Value_From_Cell(Sheetname, CN_NumPP, lastRowNum, "'" & Positio)
        set_Value_From_Cell(Sheetname, CN_ArchName, lastRowNum, Art_Param(11))
        set_Value_From_Cell(Sheetname, CN_Oboz, lastRowNum, Art_Param(0))
        set_Value_From_Cell(Sheetname, CN_Naim, lastRowNum, Art_Param(1))
        set_Value_From_Cell(Sheetname, CN_Mass, lastRowNum, Art_Param(2))
        set_Value_From_Cell(Sheetname, CN_Mass_MUIDStr, lastRowNum, Art_Param(9))
        set_Value_From_Cell(Sheetname, CN_Purchated, lastRowNum, Art_Param(13))
        set_Value_From_Cell(Sheetname, CN_IBKey_purchated, lastRowNum, Art_Param(3))
        Try
            If Art_Param(2) IsNot Nothing Then set_Value_From_Cell(Sheetname, CN_Mass, lastRowNum, Art_Param(2).ToString.Replace(",", "."))
            set_Value_From_Cell(Sheetname, CN_Mater, lastRowNum, Art_Param(5))
            set_Value_From_Cell(Sheetname, CN_MaterIBKey, lastRowNum, Art_Param(6))
        Catch ex As Exception
        End Try

        If Part_SB Then
            If Not ItsLastTreenode() Then
                set_Value_From_Cell(Sheetname, CN_Type_Article, lastRowNum, "Сборочная единица")
            Else
                set_Value_From_Cell(Sheetname, CN_Type_Article, lastRowNum, Art_Param(7))
            End If
        Else
            set_Value_From_Cell(Sheetname, CN_Type_Article, lastRowNum, Art_Param(7))
        End If

        set_Value_From_Cell(Sheetname, CN_Kolvo, lastRowNum, PRJLINK_Param(2))
        set_Value_From_Cell(Sheetname, CN_KolvoSbor, lastRowNum, Count_Summ)
        set_Value_From_Cell(Sheetname, CN_EdIzm, lastRowNum, PRJLINK_Param(11))
        set_Value_From_Cell(Sheetname, CN_Prim, lastRowNum, PRJLINK_Param(5))
        set_Value_From_Cell(Sheetname, CN_TypeLink, lastRowNum, PRJLINK_Param(10))
        set_Value_From_Cell(Sheetname, CN_ContexrtType, lastRowNum, PRJLINK_Param(9))
        set_Value_From_Cell(Sheetname, CN_ArtID, lastRowNum, ArtID)
        'set_Value_From_Cell(Sheetname, CN_VspomMat, lastRowNum, TP_Array(9))

        Try
            set_Value_From_Cell(Sheetname, CN_Raszehovka, lastRowNum, TP_Array(0))
            set_Value_From_Cell(Sheetname, CN_MatZagotovki, lastRowNum, TP_Array(3))
            set_Value_From_Cell(Sheetname, CN_MatZagotovki_IBKey, lastRowNum, TP_Array(4))
            set_Value_From_Cell(Sheetname, CN_TPArtKey, lastRowNum, TP_Array(5))
            set_Value_From_Cell(Sheetname, CN_ZagCount, lastRowNum, TP_Array(6))
            set_Value_From_Cell(Sheetname, CN_RazmerZag, lastRowNum, TP_Array(7))
            If TP_Array(2) IsNot Nothing Then set_Value_From_Cell(Sheetname, CN_NormaRashoda, lastRowNum, (TP_Array(2).ToString.Replace(",", ".")))
            Try
                set_Value_From_Cell(Sheetname, CN_NormaRashoda_MU, lastRowNum, TP_Array(10))
            Catch ex As Exception

            End Try
            'If TP_Array(1) IsNot Nothing Then set_Value_From_Cell(Sheetname, CN_KIM, lastRowNum, TP_Array(1).ToString.Replace(",", "."))
            Try
                Dim ZagSumCount As Double
                Dim Norma_ras_SUM As Double
                If CDbl(TP_Array(6)) <> 0 Then
                    ZagSumCount = Count_Summ / CDbl(TP_Array(6))
                    Norma_ras_SUM = Count_Summ * CDbl(TP_Array(2))
                    set_Value_From_Cell(Sheetname, CN_ZagSumCount, lastRowNum, ZagSumCount)
                    set_Value_From_Cell(Sheetname, CN_NormaRashoda_Sum, lastRowNum, Norma_ras_SUM.ToString.Replace(",", "."))
                Else
                    set_Value_From_Cell(Sheetname, CN_ZagSumCount, lastRowNum, 0)
                    set_Value_From_Cell(Sheetname, CN_NormaRashoda_Sum, lastRowNum, 0)
                End If
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try

        SetCellsBorderLineStyle2(Sheetname, 1, lastRowNum, last_ColNum, lastRowNum, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
        Select Case PRJLINK_Param(4)
            Case 1 'документация
                temp_color = DocumColor_Sostav ' Color.Aqua
            Case 2 'комплексы
                temp_color = Color.Aquamarine
            Case 3 'сборки
                temp_color = AssemblyColor_Sostav ' Color.Yellow
            Case 4 'детали
                temp_color = PartColor_Sostav ' Color.FromArgb(216, 228, 188) 'Color.DarkKhaki 'GreenYellow (255, 153, 0)
            Case 5 'стандартные
                temp_color = StandartColor_Sostav ' Color.IndianRed
            Case 6 'прочие
                temp_color = Pro4eeColor_Sostav '  Color.LightBlue
            Case 7 'материалы
                temp_color = MaterialColor_Sostav ' Color.Khaki
            Case 8 'комплекты
                temp_color = KomplekTColor_Sostav ' Color.LightPink
        End Select
        SetCellsColor(Sheetname, 1, lastRowNum, last_ColNum, lastRowNum, temp_color)
        If PRJLINK_Param(8) = 2 Then
            temp_color = TechContex_Sostav ' Color.Gray
            SetCellsColor(Sheetname, CN_ContexrtType, lastRowNum, last_ColNum, lastRowNum, temp_color)
        End If
        'Try
        '    Dim matVSCount As Integer = UBound(TP_VspMat_Array, 1)
        '    For i As Integer = 0 To matVSCount
        '        set_Value_From_Cell(Sheetname, CN_VspomMat, lastRowNum + i, TP_VspMat_Array(i, 0))
        '        set_Value_From_Cell(Sheetname, CN_VspomMat_IBKey, lastRowNum + i, TP_VspMat_Array(i, 1))
        '    Next
        'Catch ex As Exception

        'End Try
    End Sub
    Function TP_VspMat_Array_func(Art_ID As Integer) As Object
        Application.DoEvents()
        Try
            Dim MatNAme, MatIMKey, NormaG As String
            Dim TArt As TPServer.ITArticle = tp.Articles.ByArchCode(Art_ID)
            Dim mats As TPServer.ITMaterials = TArt.Materials
            Dim Mats_Count As Integer = mats.Count - 1
            Dim array(Mats_Count, 1) As Object  ' Array
            Dim index As Integer = 0
            Dim mat As TPServer.ITMaterial = mats.First
            While mats.EOF <> 1
                MatNAme = mat.Value("Овсм")
                MatIMKey = mat.Value("%MAT")
                'NormaG = mat.Value("НРг")
                array(index, 0) = MatNAme
                array(index, 1) = MatIMKey
                index = +1
                mat = mats.Next
            End While
            Return array
        Catch ex As Exception

        End Try
    End Function
    Dim arSize As Integer
    Dim arVspMat(0, 0) As String

    'Public Property Part_SB As Boolean
    '    Get
    '        Return _part_SB
    '    End Get
    '    Set(value As Boolean)
    '        _part_SB = value
    '    End Set
    'End Property

    Function WalkInGroupMats(grpmats As TPServer.ITGroupMaterials, ReadInArray As Boolean)
        Try
            Dim Group As TPServer.ITGroupMaterial = grpmats.First
            While grpmats.EOF <> 1
                If Group.Status = 0 Then
                    GetVspMatInGrMat(Group.Materials, ReadInArray)

                    If Group.GroupMaterials.Count > 0 Then WalkInGroupMats(Group.GroupMaterials, ReadInArray)
                End If
                Group = grpmats.Next()
            End While
        Catch ex As Exception

        End Try
    End Function
    Function GetVspMatInGrMat(mats As TPServer.ITMaterials, ReadInArray As Boolean)
        Try
            Dim mat As TPServer.ITMaterial = mats.First
            Dim MatNAme, MatIMKey, NormaG, MaterMU As String
            While mats.EOF <> 1
                If mat.Materials.Count > 0 Then
                    GetVspMatInGrMat(mat.Materials, ReadInArray)
                Else
                    arSize += 1
                    If ReadInArray Then
                        Try
                            MatNAme = mat.Value("Овсм")
                            MatIMKey = mat.Value("%MAT")
                            NormaG = mat.Norma
                            MaterMU = mat.Value("едНв")
                            'If MaterMU Is Nothing Then
                            '    If mats.EOF = 1 Then
                            '        mat = mats.Prior
                            '        mat = mats.Next
                            '        MaterMU = mat.Value("едНв")

                            '    Else
                            '        mat = mats.Next
                            '        mat = mats.Prior
                            '        MaterMU = mat.Value("едНв")
                            '    End If
                            'End If
                            arVspMat(arSize - 1, 0) = get_FindReplace_Par(MatNAme)
                            arVspMat(arSize - 1, 1) = MatIMKey
                            arVspMat(arSize - 1, 2) = NormaG
                            arVspMat(arSize - 1, 3) = MaterMU
                        Catch ex As Exception
                        End Try
                    Else
                        ReDim arVspMat(arSize - 1, 3)
                    End If
                End If
                mat = mats.Next
            End While
        Catch ex As Exception
        End Try
    End Function
    Function Get_Vspom_Mater_Array(Art_ID As Integer, Proj_Id As Integer) As Array
        Application.DoEvents()
        Dim MaterName, MaterIBKey, MaterNorma, MaterMU As String
        ReDim arVspMat(0, 0)
        Try
            Dim TArt As TPServer.ITArticle = tp.Articles.ByArchCode(Art_ID)
            Dim GrMat As TPServer.ITGroupMaterials = TArt.GroupMaterials
            arSize = 0
            WalkInGroupMats(GrMat, False) 'считаем размер массива arVspMat(кол-во вспомогательных материалов)
            arSize = 0
            WalkInGroupMats(GrMat, True) 'записывает данные в массив

            Return arVspMat
        Catch ex As Exception
            Return arVspMat
        End Try
    End Function
    Function Get_TP_ParmArray(Art_ID As Integer, Proj_Id As Integer) As Array
        Application.DoEvents()
        Dim array(10) As String
        Try
            Dim rascexovka, KIM, Norma, Sortament, SortamentIMKey, ZagCount, Zag_key, RazmZag, Norma_MU As String
            Dim TArt As TPServer.ITArticle = tp.Articles.ByArchCode(Art_ID)

            Dim TRout As TPServer.TRoute = TArt.Route
            Dim count_Rout As Integer = TRout.VarCount
            Dim TRout1 As TPServer.TRouteVariant = TRout.First()
            For g As Integer = 0 To count_Rout - 1
                If TRout1.isDefault = 1 Then
                    rascexovka = TRout1.Stroka
                    Exit For
                End If
                TRout1 = TRout.Next()
            Next
            Dim zagsCount As Integer
            Try
                Dim Zags As TPServer.ITZags = TArt.Zags
                zagsCount = Zags.Count
                Dim Zag As TPServer.ITZag = Zags.First
                For i As Integer = 0 To zagsCount - 1
                    If Zag.IsMainVar <> 1 Then
                        Dim RefArts As TPServer.ITArticle '= Zag.InArts.First
                        For j As Integer = 0 To Zag.InArts.Count - 1
                            RefArts = Zag.InArts.Articles(j)
                            'ищем заготовку по применяемости! Если условие не строгое, то нужно дать значение "-1"
                            Try
                                If (RefArts.ArchID = Proj_Id Or Proj_Id = -1 Or Proj_Id = 0) Then

                                    With Zag
                                        KIM = .Value("КИМ")
                                        Norma = .Value("НР")
                                        Norma_MU = .Value("едНР")
                                        Sortament = .Value("SORT")
                                        SortamentIMKey = .Value("%ZAG")
                                        ZagCount = .Value("КЗаг")
                                        Zag_key = .Key
                                        RazmZag = .Value("РАЗМ")
                                    End With
                                    Exit For
                                End If
                            Catch ex As Exception

                            End Try
                            'RefArts = Zag.InArts.Next
                        Next
                    ElseIf Zag.IsMainVar = 1 Then
                        With Zag
                            KIM = .Value("КИМ")
                            Norma = .Value("НР")
                            Norma_MU = .Value("едНР")
                            Sortament = .Value("SORT")
                            SortamentIMKey = .Value("%ZAG")
                            ZagCount = .Value("КЗаг")
                            Zag_key = .Key
                            RazmZag = .Value("РАЗМ")
                        End With
                        Exit For
                    End If
                    Zag = Zags.Next
                Next

                If ReplaceTildaOnSpace Then
                    Sortament = Replace(Sortament, "~", " ")
                End If
                If ReplaceQuestionOnDrob Then
                    Sortament = Replace(Sortament, "?", "/")
                End If
                Sortament = get_FindReplace_Par(Sortament)
            Catch ex As Exception

            End Try

            If Norma_MU = "" Or Norma_MU Is Nothing Then Norma_MU = "кг"

            array(0) = rascexovka
            array(1) = KIM
            array(2) = Norma
            array(3) = Sortament
            array(4) = SortamentIMKey
            array(5) = Zag_key
            array(6) = ZagCount
            array(7) = RazmZag
            array(8) = zagsCount
            'array(9) = MatVspom
            array(10) = Norma_MU
            Return array
        Catch ex As Exception
            Return array
        End Try
    End Function
    Function get_FindReplace_Par(Input_paramet As String)
        Try
            For u As Integer = 0 To UBound(ReplaceChar)
                Dim tmp_line() As String = ReplaceChar(u).Split("=")
                Dim Finf_txt As String = tmp_line(0).Trim("=")
                Dim Replace_txt As String = tmp_line(1).Trim("=")
                Input_paramet = Input_paramet.Replace(Finf_txt, Replace_txt)
            Next
            Return Input_paramet
        Catch ex As Exception

        End Try
    End Function
    Function Get_Link_Param(PRJLINK_ID As Integer) As Array
        Application.DoEvents()
        Try
            With s4

                .OpenQuery("select * from PC where PRJLINK_ID = " & PRJLINK_ID)
                Dim PROJ_AID As Integer = .QueryFieldByName("PROJ_AID")
                Dim Part_AID As Integer = .QueryFieldByName("Part_AID")
                Dim COUNT_PC As Double = .QueryFieldByName("COUNT_PC")
                Dim MU_ID As Integer = .QueryFieldByName("MU_ID")
                Dim RAZDEL As Integer = .QueryFieldByName("RAZDEL")
                Dim NOTE As String = .QueryFieldByName("NOTE")
                Dim LINK_TYPE As String = .QueryFieldByName("LINK_TYPE")
                Dim FORMAT As String = .QueryFieldByName("FORMAT")
                Dim CTX_ID As Integer = .QueryFieldByName("CTX_ID")
                Dim CTX_IDStr As String
                Dim LINK_TYPE_str As String
                Dim MU_ID_str As String
                Select Case CTX_ID
                    Case 2
                        CTX_IDStr = "Технологический"
                    Case Else
                        CTX_IDStr = "Конструкторский"
                End Select
                If LINK_TYPE IsNot Nothing Then LINK_TYPE_str = "Ручная"
                .CloseQuery()
                .OpenQuery("select * from MU where MU_ID = '" & MU_ID & "'")
                MU_ID_str = .QueryFieldByName("MU_SHORT_NAME")
                .CloseQuery()
                Dim array(11) As String
                array(0) = PROJ_AID
                array(1) = Part_AID
                array(2) = COUNT_PC
                array(3) = MU_ID
                array(4) = RAZDEL
                array(5) = NOTE
                array(6) = LINK_TYPE
                array(7) = FORMAT
                array(8) = CTX_ID
                array(9) = CTX_IDStr
                array(10) = LINK_TYPE_str
                array(11) = MU_ID_str
                Return array
            End With
        Catch ex As Exception

        End Try
    End Function
    Function Get_Short_Article_Param(Art_ID As Integer) As Array
        Application.DoEvents()
        Dim array(4) As String
        Try
            With s4
                .OpenArticle(Art_ID)
                '.ReturnFieldValueWithImbaseKey = 1
                Dim oboz, naim, mass, ArchID, ArchName As String
                oboz = .GetFieldValue_Articles("Обозначение")
                naim = .GetFieldValue_Articles("Наименование")
                mass = .GetArticleMassa
                Try
                    Dim docID As Integer = .GetArticleDocID
                    If docID > 0 Then
                        .OpenDocument(docID)
                        ArchID = .GetFieldValue("Архив")
                        ArchName = .nfoGetArchiveDescription(ArchID)
                        .CloseDocument()
                    Else
                        ArchName = "NODOC"
                    End If
                Catch ex As Exception
                    ArchName = "NODOC"
                End Try

                .CloseArticle()
                array(0) = oboz
                array(1) = naim
                array(2) = mass
                array(3) = ArchID
                array(4) = ArchName

                Return array
            End With
        Catch ex As Exception

            Return array
        End Try
    End Function
    Function Get_Article_Param(Art_ID As Integer) As Array
        Application.DoEvents()
        Dim array(13) As String
        Try
            With s4
                .OpenArticle(Art_ID)
                .ReturnFieldValueWithImbaseKey = 0
                Dim oboz, naim, mass, Imbase_key, purchased, purchased_str, material, materialIMKey, ArtKindName, mass_MU_ID, mass_MU_ID_Str, ArchID, ArchName, ArtKind As String
                oboz = .GetFieldValue_Articles("Обозначение")
                naim = .GetFieldValue_Articles("Наименование")
                mass = .GetArticleMassa
                Imbase_key = .GetFieldValue_Articles("Ключ Imbase")
                purchased = .GetArticlePurchased

                Select Case purchased
                    Case "+"
                        purchased_str = "Покупное"
                    Case "*"
                        purchased_str = "По кооперации"
                    Case "-"
                        purchased_str = "Собственное"
                    Case "!"
                        purchased_str = "Не изготавливать"
                    Case "&"
                        purchased_str = "Неопределен"
                    Case "^"
                        purchased_str = "Доработка"
                End Select

                material = .GetArticleMaterial
                If ReplaceTildaOnSpace Then
                    material = Replace(material, "~", " ")
                    naim = Replace(naim, "~", " ")
                End If
                If ReplaceQuestionOnDrob Then
                    material = Replace(material, "?", "/")
                    naim = Replace(naim, "?", "/")
                End If
                material = get_FindReplace_Par(material)
                naim = get_FindReplace_Par(naim)

                materialIMKey = .GetFieldImbaseKey_Articles("Материал")
                ArtKindName = .GetArtKindName
                ArtKind = .GetArticleKind
                Try
                    Dim docID As Integer = .GetArticleDocID
                    If docID > 0 Then
                        .OpenDocument(docID)
                        ArchID = .GetFieldValue("Архив")
                        ArchName = .nfoGetArchiveDescription(ArchID)
                        .CloseDocument()
                    Else
                        ArchName = "NODOC"
                    End If
                Catch ex As Exception
                    ArchName = "NODOC"
                End Try
                'If material IsNot Nothing Then
                '    Dim mat_pat_array As Array = material.Split(mat_splitter)
                '    material = mat_pat_array(0).ToString.Trim(mat_splitter)
                '    materialIMKey = mat_pat_array(1).ToString.Trim(mat_splitter)
                'End If

                .CloseArticle()
                .OpenQuery("select * from ARTICLES where ART_ID = '" & Art_ID & "'")
                mass_MU_ID = .QueryFieldByName("MU_ID")
                .CloseQuery()

                .OpenQuery("select * from MU where MU_ID = '" & mass_MU_ID & "'")
                mass_MU_ID_Str = .QueryFieldByName("MU_SHORT_NAME")
                .CloseQuery()

                array(0) = oboz
                array(1) = naim
                array(2) = mass
                array(3) = Imbase_key
                array(4) = purchased
                array(5) = material
                array(6) = materialIMKey
                array(7) = ArtKindName
                array(8) = mass_MU_ID
                array(9) = mass_MU_ID_Str
                array(10) = ArchID
                array(11) = ArchName
                array(12) = ArtKind
                array(13) = purchased_str
                Return array
            End With
        Catch ex As Exception

            Return array
        End Try
    End Function

    Private Sub БезРазделаДокументацияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БезРазделаДокументацияToolStripMenuItem.Click
        БезРазделаДокументацияToolStripMenuItem.Checked = Not (БезРазделаДокументацияToolStripMenuItem.Checked)
        NO_Art_Documentacia = БезРазделаДокументацияToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Art_Documentacia_OpName, NO_Art_Documentacia)
        CT_ID_in_Query_Change()
    End Sub

    Private Sub БезТехнолическихСвязейToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БезТехнолическихСвязейToolStripMenuItem.Click
        БезТехнолическихСвязейToolStripMenuItem.Checked = Not (БезТехнолическихСвязейToolStripMenuItem.Checked)
        NO_Ru4na9_Sv9z = БезТехнолическихСвязейToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Ru4na9_Sv9z_OpName, NO_Ru4na9_Sv9z)
        CT_ID_in_Query_Change()
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles Справка.Click
        info()
    End Sub
    Sub info()
        Process.Start("http://www.evernote.com/l/AjJuLpie6c9F_ZTbh_5rFgRgRcQN9CrJnZ8/")
    End Sub
    Private Sub ОткрытьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьToolStripMenuItem.Click
        Dim selectNode As TreeNode = TreeView1.SelectedNode
        Try
            Dim param_Array As Array = selectNode.Tag.ToString.Split(Spletter)
            With s4
                .OpenArticle(param_Array(1))
                .EditParameters_Article()
                .CloseArticle()
            End With
        Catch ex As Exception
            With s4
                .OpenArticle(selectNode.Tag.ToString)
                .EditParameters_Article()
                .CloseArticle()
            End With
        End Try
    End Sub

    Private Sub ОчиститьДеревоToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОчиститьДеревоToolStripMenuItem.Click
        TreeCleaner()
    End Sub
    Private Sub TreeCleaner()
        Try
            TreeView1.Nodes.Item(0).Remove()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub УдалитьВетвьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьВетвьToolStripMenuItem.Click
        Try
            Dim selNode As TreeNode = TreeView1.SelectedNode
            If selNode Is Nothing Then Exit Sub
            If MsgBox("Удалить выделенный узел?", vbYesNo + MsgBoxStyle.Question, "Удалить ветвь") = MsgBoxResult.Yes Then
                Try
                    TreeView1.SelectedNode.Remove()
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ЗаменятьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗаменятьToolStripMenuItem.Click
        ЗаменятьToolStripMenuItem.Checked = Not (ЗаменятьToolStripMenuItem.Checked)
        If ЗаменятьToolStripMenuItem.Checked Then ReplaceTildaOnSpace = 1 Else ReplaceTildaOnSpace = 0
    End Sub

    Sub CT_ID_in_Query_Change()
        If ToolStripComboBox1.Text = "Конструкторский" Then
            CT_ID_in_Query = " and (ctx_id = 0 or ctx_id = 1) "
        ElseIf ToolStripComboBox1.Text = "Комбинированный" Then
            CT_ID_in_Query = " "
        Else
            CT_ID_in_Query = " and (ctx_id = 0 or ctx_id = 2)"
        End If
        treeUpdate()
    End Sub
    Sub treeUpdate()
        If TreeView1.Nodes.Count = 0 Then Exit Sub
        If MsgBox("В настройке произошли изменения" & vbNewLine & "Обновить дерево состава в приложении?", vbYesNo + vbQuestion, "ВСНРМ 2.0") = DialogResult.Yes Then
            'процедура обновления дерево состава
            generalTree_Bez_Positio(1)
        End If
    End Sub
    Private Sub ToolStripComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.TextChanged
        CT_ID_in_Query_Change()
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        CT_ID_in_Query_Change()
    End Sub

    Private Sub СПеречнемДеталейToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СПеречнемДеталейToolStripMenuItem.Click
        СПеречнемДеталейToolStripMenuItem.Checked = Not (СПеречнемДеталейToolStripMenuItem.Checked)
        NO_Parts = СПеречнемДеталейToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Parts_OpName, NO_Parts)
    End Sub

    Private Sub СПеречнемМатериаловToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СПеречнемМатериаловToolStripMenuItem.Click
        СПеречнемМатериаловToolStripMenuItem.Checked = Not (СПеречнемМатериаловToolStripMenuItem.Checked)
        NO_Purchated = СПеречнемМатериаловToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Purchated_OpName, NO_Purchated)
    End Sub


    Private Sub ToolStripButton4_Click_1(sender As Object, e As EventArgs)
        Get_Article_Param(15209)
        'Dim values As New List(Of Tuple(Of String, String, String))
        'values.Add(Tuple.Create("a1", "a2", "a2"))
        'values.Add(Tuple.Create("b1", "b2", "a2"))
        'values.Add(Tuple.Create("c1", "c2", "a2"))
        'Get_Vspom_Mater_Array(15383, -1)
    End Sub

    Private Sub ПоказыватьПроцессЭкспортаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоказыватьПроцессЭкспортаToolStripMenuItem.Click
        ПоказыватьПроцессЭкспортаToolStripMenuItem.Checked = Not (ПоказыватьПроцессЭкспортаToolStripMenuItem.Checked)
        NO_ExlProc_Visible = ПоказыватьПроцессЭкспортаToolStripMenuItem.Checked
        chekced_ParamChancge(NO_ExlProc_Visible_OpName, NO_ExlProc_Visible)
    End Sub

    Private Sub ToolStripButton3_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        TreeCleaner()
    End Sub

    Private Sub ПоказыватьСистемнуюИнформациюToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоказыватьСистемнуюИнформациюToolStripMenuItem.Click
        ПоказыватьСистемнуюИнформациюToolStripMenuItem.Checked = Not (ПоказыватьСистемнуюИнформациюToolStripMenuItem.Checked)
        NO_S4_Columns = ПоказыватьСистемнуюИнформациюToolStripMenuItem.Checked
        chekced_ParamChancge(NO_S4_Columns_OpName, NO_S4_Columns)
    End Sub

    Private Sub СРазделомСборочныеЕдиницыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СРазделомСборочныеЕдиницыToolStripMenuItem.Click
        no_SborkaProc()
    End Sub
    Sub no_SborkaProc()
        СРазделомСборочныеЕдиницыToolStripMenuItem.Checked = Not (СРазделомСборочныеЕдиницыToolStripMenuItem.Checked)
        NO_Sborka = СРазделомСборочныеЕдиницыToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Sborka_OpName, NO_Sborka)
        CT_ID_in_Query_Change()
    End Sub
    Private Sub СРазделомДеталиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СРазделомДеталиToolStripMenuItem.Click
        no_PartProc()
    End Sub
    Sub no_PartProc()
        СРазделомДеталиToolStripMenuItem.Checked = Not (СРазделомДеталиToolStripMenuItem.Checked)
        NO_PartSecion = СРазделомДеталиToolStripMenuItem.Checked
        chekced_ParamChancge(NO_PartSecion_OpName, NO_PartSecion)
        CT_ID_in_Query_Change()
    End Sub
    Private Sub СРазделомСтандартныеЕдиницыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СРазделомСтандартныеЕдиницыToolStripMenuItem.Click
        СРазделомСтандартныеЕдиницыToolStripMenuItem.Checked = Not (СРазделомСтандартныеЕдиницыToolStripMenuItem.Checked)
        NO_Standart = СРазделомСтандартныеЕдиницыToolStripMenuItem.Checked
        CT_ID_in_Query_Change()
    End Sub
    Sub no_StandProc()
        СРазделомСтандартныеЕдиницыToolStripMenuItem.Checked = Not (СРазделомСтандартныеЕдиницыToolStripMenuItem.Checked)
        NO_Standart = СРазделомСтандартныеЕдиницыToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Standart_OpName, NO_Standart)
        CT_ID_in_Query_Change()
    End Sub
    Private Sub СРазделомПрочиеИзделияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СРазделомПрочиеИзделияToolStripMenuItem.Click
        no_Pro4Proc()
    End Sub
    Sub no_Pro4Proc()
        СРазделомПрочиеИзделияToolStripMenuItem.Checked = Not (СРазделомПрочиеИзделияToolStripMenuItem.Checked)
        NO_Pro4ee = СРазделомПрочиеИзделияToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Pro4ee_OpName, NO_Pro4ee)
        CT_ID_in_Query_Change()
    End Sub
    Private Sub СРазделомМатериалыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СРазделомМатериалыToolStripMenuItem.Click
        no_MatProc()
    End Sub
    Sub no_MatProc()
        СРазделомМатериалыToolStripMenuItem.Checked = Not (СРазделомМатериалыToolStripMenuItem.Checked)
        NO_Material = СРазделомМатериалыToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Material_OpName, NO_Material)
        CT_ID_in_Query_Change()
    End Sub
    Private Sub ToolStripButton4_Click_2(sender As Object, e As EventArgs)
        Get_Vspom_Mater_Array(16188, -1)
    End Sub

    Private Sub VSNRM2_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown

        If e.Control And e.KeyCode.ToString = "O" Then
            Me.KeyPreview = False
            no_MatProc()
            no_PartProc()
            no_Pro4Proc()
            no_SborkaProc()
            no_StandProc()


        End If
        Me.KeyPreview = True
    End Sub

    Private Sub ToolStripButton4_Click_3(sender As Object, e As EventArgs)
        Get_TP_ParmArray(16358, -1)
    End Sub

    Private Sub ToolStripButton4_Click_4(sender As Object, e As EventArgs)
        TP_VspMat_Array_func(16353)
    End Sub

    Private Sub ПоказыватьПояснениеПоЗаливкеToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub НастройкиЗаливкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НастройкиЗаливкиToolStripMenuItem.Click
        ColorOptions.ShowDialog()
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        Try
            Dim t_node As TreeNode = TreeView1.SelectedNode
            t_node.SelectedImageIndex = t_node.ImageIndex
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ДетальСборочнаяЕдиницаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДетальСборочнаяЕдиницаToolStripMenuItem.Click
        ДетальСборочнаяЕдиницаToolStripMenuItem.Checked = Not (ДетальСборочнаяЕдиницаToolStripMenuItem.Checked)
        Part_SB = ДетальСборочнаяЕдиницаToolStripMenuItem.Checked
        chekced_ParamChancge(Part_SB_OpName, part_SB)
        'CT_ID_in_Query_Change()
    End Sub

    Private Sub СПеречнемПокупныхПДРБНToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СПеречнемПокупныхПДРБНToolStripMenuItem.Click
        СПеречнемПокупныхПДРБНToolStripMenuItem.Checked = Not (СПеречнемПокупныхПДРБНToolStripMenuItem.Checked)
        If СПеречнемПокупныхПДРБНToolStripMenuItem.Checked Then NO_Purchated_PDRBN = 1 Else NO_Purchated_PDRBN = 0
        chekced_ParamChancge(NO_Purchated_PDRBN_OpName, NO_Purchated_PDRBN)
    End Sub

    Private Sub ЗаменятьНаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗаменятьНаToolStripMenuItem.Click
        ЗаменятьНаToolStripMenuItem.Checked = Not (ЗаменятьНаToolStripMenuItem.Checked)
        If ЗаменятьНаToolStripMenuItem.Checked Then ReplaceQuestionOnDrob = 1 Else ReplaceQuestionOnDrob = 0
    End Sub


    Private Sub ToolStripDropDownButton1_OwnerChanged(sender As Object, e As EventArgs) Handles ToolStripDropDownButton1.OwnerChanged
        If Keys.KeyCode = Keys.Shift Then
            Me.ContextMenuStrip1.Close(ToolStripDropDownCloseReason.CloseCalled)
        End If
    End Sub

    Private Sub ToolStripDropDownButton1_DropDownOpened(sender As Object, e As EventArgs) Handles ToolStripDropDownButton1.DropDownOpened

    End Sub

    Private Sub ССоставомИзделияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ССоставомИзделияToolStripMenuItem.Click
        ССоставомИзделияToolStripMenuItem.Checked = Not (ССоставомИзделияToolStripMenuItem.Checked)
        NO_Sostav = ССоставомИзделияToolStripMenuItem.Checked
        chekced_ParamChancge(NO_Sostav_OpName, NO_Sostav)
    End Sub

    Private Sub ТолькоПервыйУровеньToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТолькоПервыйУровеньToolStripMenuItem.Click
        ТолькоПервыйУровеньToolStripMenuItem.Checked = Not (ТолькоПервыйУровеньToolStripMenuItem.Checked)
        OnlyFirstLewvel = ТолькоПервыйУровеньToolStripMenuItem.Checked
        chekced_ParamChancge(OnlyFirstLewvel_OpName, OnlyFirstLewvel)
        CT_ID_in_Query_Change()
    End Sub

    Private Sub ВедомостьДляВыбранногоУзлаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВедомостьДляВыбранногоУзлаToolStripMenuItem.Click
        ReplaceCharArray()
        Dim selectNode As TreeNode = TreeView1.SelectedNode
        add_Vedomost(selectNode)
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        Try
            Dim selectNode As TreeNode = TreeView1.SelectedNode
            If selectNode.GetNodeCount(True) = 0 Then
                ВедомостьДляВыбранногоУзлаToolStripMenuItem.Enabled = False
            Else
                ВедомостьДляВыбранногоУзлаToolStripMenuItem.Enabled = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub НастройкиЗаменыТекстаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НастройкиЗаменыТекстаToolStripMenuItem.Click
        CharOption.ShowDialog()
    End Sub

    Private Sub ПоказыватьПояснениеПоЗаливкеToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ПоказыватьПояснениеПоЗаливкеToolStripMenuItem.Click
        ПоказыватьПояснениеПоЗаливкеToolStripMenuItem.Checked = Not (ПоказыватьПояснениеПоЗаливкеToolStripMenuItem.Checked)
        NO_ColorNote = ПоказыватьПояснениеПоЗаливкеToolStripMenuItem.Checked
        chekced_ParamChancge(NO_ColorNote_OpName, NO_ColorNote)
    End Sub

    Private Sub ToolStripButton4_Click_5(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        CT_ID_in_Query_Change()
    End Sub

    Sub NextLevelInTreeView_Bez_Positio(node As TreeNode, Proj_Aid As Integer, SubPositio As String)
        Application.DoEvents()
        With s4

            .OpenArticle(Proj_Aid)
            Dim VERS As String = .GetFieldValue_Articles("Art_Ver_ID")
            .CloseArticle()

            Dim max_poz_mun As Integer = GetMaxRowInBOM(Proj_Aid, VERS)
            If Verartfiltr Then
                .OpenQuery("select * from PC where proj_aid = " & Proj_Aid & " and (PROJ_VER_ID = NULL OR PROJ_VER_ID = " & VERS & ")" & CT_ID_in_Query)
            Else
                .OpenQuery("select * from PC where proj_aid = " & Proj_Aid & CT_ID_in_Query) ' & " and PROJ_VER_ID = " & VERS)
            End If
            .QueryGoFirst()
            For i As Integer = 0 To max_poz_mun - 1
                Dim TempPos As String = SubPositio & "." & .QueryFieldByName("positio")
                Dim PRJLINK_ID As Integer = .QueryFieldByName("PRJLINK_ID")
                Dim Part_AID As Integer = .QueryFieldByName("Part_AID")
                Dim RAZDEL As Integer = .QueryFieldByName("RAZDEL")
                Dim LINK_TYPE As String = .QueryFieldByName("LINK_TYPE")
                Dim CTX_ID As Integer = .QueryFieldByName("CTX_ID")
                Dim COUNT_PC As Double = .QueryFieldByName("COUNT_PC")
                'Dim ge_Article_Param_List As Array = Get_Article_Param(Part_AID)

                If check_OnOffOptionsforBOMComponents(RAZDEL, LINK_TYPE, CTX_ID) Then
                    Dim subNode As TreeNode = TreeNodeAdd2(node, Part_AID, PRJLINK_ID, TempPos, COUNT_PC)

                    If exist_BOM_ChildNodesExist(Part_AID) Then
                        NextLevelInTreeView_Bez_Positio(subNode, Part_AID, TempPos)
                        Perehov_Na_NUGHUY_strocu(Proj_Aid, i, VERS)
                    Else
                        Perehov_Na_NUGHUY_strocu(Proj_Aid, i, VERS)
                    End If
                Else
                    Perehov_Na_NUGHUY_strocu(Proj_Aid, i, VERS)
                End If
            Next

            .CloseQuery()
        End With
    End Sub

End Class