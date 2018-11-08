Public Class VSNRM2
    Public s4 As S4.TS4App = Form1.ITS4App
    Public tp As TPServer.ITApplication
    'Public ib As ImBase.IImbaseApplication
    Public artID_Main, artId_Sub As Integer
    Public Spletter As String = "%"
    'параметр не выводящий типы объектов и компоненты с ручной связью
    Public NO_Art_Documentacia As Boolean = 0
    Public NO_Ru4na9_Sv9z As Boolean = 0
    'параметр заменяющий "~" на " "
    Public ReplaceTildaOnSpace As Boolean = 1
    'параметр модульный
    Dim Verartfiltr_str As String = "Verartfiltr" 'parameter name
    Dim Verartfiltr As Boolean = 0
    Dim TPsfilter_str As String = "TPsfilter"
    Dim TPsfilter As Boolean = 0
    Dim CT_ID_in_Query As String
    'параметр экспортирующий данные о применяемости деталей
    Public NO_Parts As Boolean = 1
    'параметр экспортирующий данные о материалах и покупных
    Public NO_Purchated As Boolean = 1

    Public myImageList As New ImageList()
    'Public NO_TSE As Boolean = 0
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
            MsgBox("Ошибка при проверке на наличие подузлов" & vbNewLine & ex.Message )
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
    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)
        tp = CreateObject("TPServer.TApplication")
        Dim status As Integer = tp.Ready
        While True
            If status = 1 Then Exit While
        End While

        Dim TArt As TPServer.ITArticle = tp.Articles.ByArchCode(13775)

        Dim oboz As String = TArt.Designation


        Dim TRout As TPServer.TRoute = TArt.Route
        Dim count_Rout As Integer = TRout.VarCount
        Dim TRout1 As TPServer.TRouteVariant = TRout.First()

        For j As Integer = 0 To count_Rout - 1
            Dim all_par As String = TRout1.AllParams
            TRout1 = TRout.Next()
        Next
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

                        If exist_BOM_ChildNodesExist(Part_AID) Then
                            NextLevelInTreeView_Bez_Positio(subNode, Part_AID, TempPos)
                            Perehov_Na_NUGHUY_strocu(artID_Main, i, VERS)
                        Else
                            Perehov_Na_NUGHUY_strocu(artID_Main, i, VERS)
                        End If
                    Else
                        Perehov_Na_NUGHUY_strocu(artID_Main, i, VERS)
                    End If
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
        БезРазделаДокументацияToolStripMenuItem.Checked = NO_Art_Documentacia
        БезТехнолическихСвязейToolStripMenuItem.Checked = NO_Ru4na9_Sv9z
        СПеречнемДеталейToolStripMenuItem.Checked = NO_Parts
        СПеречнемМатериаловToolStripMenuItem.Checked = NO_Purchated
        ЗаменятьToolStripMenuItem.Checked = ReplaceTildaOnSpace
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

    Public OPT_FirstShow As String = "ExcelReportFirstShow"
    Public OPT_LastInPut As String = "LastInPut"

    'ExcelSheets column numbers ВТОРОЙ ЛИСТ(ПЕРЕЧЕНЬ ПОКУПНЫХ(например))
    Public ShName_Purchated As String = "ПЕРЕЧЕНЬ ПОКУПНЫХ"
    Public RowN_Purchated_First As Integer = 3
    Public CN_Purchated_Naim As Integer = 1
    Public CN_Purchated_IBKey As Integer = 2
    Public CN_Purchated_Count As Integer = 3
    Public CN_Purchated_MU As Integer = 4
    Private Sub addExTitlePurchated()
        Add_NewSheet(ShName_Purchated)
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, RowN_Purchated_First - 1, "Наименование")
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First - 1, "Ключ IMBASE")
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, RowN_Purchated_First - 1, "Общ. кол-во/масса")
        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, RowN_Purchated_First - 1, "Ед.изм.")

        Dim CN_Purchated_last_ColNum As Integer = Get_Last_Column(ShName_Purchated, RowN_Purchated_First - 1)
        SetCellsBorderLineStyle2(ShName_Purchated, 1, 1, CN_Purchated_last_ColNum, RowN_Purchated_First - 1, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
        Try
            Dim VSpom_Mater_for_Main As Array = Get_Vspom_Mater_Array(TreeView1.Nodes.Item(0).Tag, -1)
            If VSpom_Mater_for_Main IsNot Nothing Then
                excel_write_aboutVSPomMater_in_Purchated(VSpom_Mater_for_Main)
            End If
        Catch ex As Exception
        End Try
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
    Public CN_PL_Sort As Integer = 9
    Public CN_PL_Sort_IBKey As Integer = 10
    Public CN_PL_Referenc As Integer = 11
    Private Sub addExTitlePartList()
        Add_NewSheet(ShName_PartList)
        set_Value_From_Cell(ShName_PartList, CN_PL_ArtID, RowN_PL_First - 1, "ArtID")
        set_Value_From_Cell(ShName_PartList, CN_PL_Oboz, RowN_PL_First - 1, "Обозначение Детали")
        set_Value_From_Cell(ShName_PartList, CN_PL_Naim, RowN_PL_First - 1, "Наименование Детали")
        set_Value_From_Cell(ShName_PartList, CN_PL_PROJ_ID, RowN_PL_First - 1, "PROJID")
        set_Value_From_Cell(ShName_PartList, CN_PL_Primen9emost, RowN_PL_First - 1, "Применяемость")
        set_Value_From_Cell(ShName_PartList, CN_PL_Count, RowN_PL_First - 1, "Кол-во на сб, шт.")
        set_Value_From_Cell(ShName_PartList, CN_PL_TotalCount, RowN_PL_First - 1, "Общ.кол-во, шт.")
        'set_Value_From_Cell(ShName_PartList, CN_PL_MU, RowN_PL_First - 1, "ед.изм")
        set_Value_From_Cell(ShName_PartList, CN_PL_LinkType, RowN_PL_First - 1, "Тип связи")
        set_Value_From_Cell(ShName_PartList, CN_PL_Sort, RowN_PL_First - 1, "Сортамент")
        set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, RowN_PL_First - 1, "Ключ IMBASE")
        Dim CN_PL_last_ColNum As Integer = Get_Last_Column(ShName_PartList, RowN_PL_First - 1)
        SetCellsBorderLineStyle2(ShName_PartList, 1, 1, CN_PL_last_ColNum, RowN_PL_First - 1, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)

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

    Sub WriteEXCellTitle()
        Application.DoEvents()
        Create_EX_Doc(True)
        SheetReName("Лист1", Sheetname)
        Dim MainTPInfo, MainArtInfo As Array
        Try
            MainArtInfo = Get_Short_Article_Param(TreeView1.Nodes.Item(0).Tag)

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
        set_Value_From_Cell(Sheetname, CN_Purchated, RowN_First - 1, "Покупное")
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
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Application.DoEvents()
        If TreeView1.Nodes.Count = 0 Then
            MsgBox("Для выгрузки Ведомости, выберите Объект")
            Exit Sub
        End If
        WriteEXCellTitle()

        If NO_Purchated Then addExTitlePurchated() 'создать ВЕДОМОСТЬ ПОКУПНЫХ
        If NO_Parts Then addExTitlePartList() 'создать ВЕДОМОСТЬ ДЕТАЛЕЙ

        read_NEXTLEVELtreeview(TreeView1.Nodes.Item(0))

        last_Rowmun = Get_LastRowInOneColumn(Sheetname, CN_Naim)


        'ровнение текста в шапке
        CellsTextHorisAligment(Sheetname, 1, 1, last_ColNum, RowN_First - 1, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter)
        CellsTextVerticalAligment(Sheetname, 1, 1, last_ColNum, RowN_First - 1, Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter)

        CellsTextHorisAligment(Sheetname, CN_NumPP, RowN_First, last_ColNum, last_Rowmun, Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft)

        SetAutoFIT(Sheetname)
        If NO_Purchated Then SetAutoFIT(ShName_Purchated)
        If NO_Parts Then SetAutoFIT(ShName_PartList)

        'SetCellsCdbl(Sheetname, CN_NumPP, RowN_First, CN_NumPP, last_Rowmun) 'преобразовал в число столбцы с позицией
        'SetCellsCdbl(Sheetname, CN_Mass, RowN_First, CN_Mass, last_Rowmun) 'преобразовал в число столбцы с массой
    End Sub
    Sub addExcelAboutArtStructure(ArtID As Integer)
        If NO_Parts Then addExTitlePartList()
        If NO_Purchated Then addExTitlePurchated()

        With s4
            .OpenArticleStructureExpanded2(ArtID, -1, 2)
            .asFirst()
            Dim GetArtKind As Integer = .asGetArtKind
            While .asEof = 0
                Select Case GetArtKind
                    Case 4, 5, 6, 7
                        Dim tmp_ArtID As Integer = .asGetArtID
                        Dim total_count As Integer = .asGetArtCount
                        Dim Count_MU As String = .asGetArtCount_MU_ID
                        Dim S4_Info, TC_Info As Array
                        S4_Info = Get_Article_Param(tmp_ArtID)
                        If TPsfilter Then
                            'TC_Info = Get_TP_ParmArray(tmp_ArtID)
                        End If
                        If GetArtKind = 4 And NO_Parts Then
                            'процедура длязаполнения листа ПЕРЕЧЕНЬ ДЕТАЛЕЙ
                            'excel_write_aboutPart_in_PartList(tmp_ArtID, S4_Info, TC_Info, total_count, Count_MU)
                        End If
                        If NO_Purchated Then
                            'процедура длязаполнения листа ПЕРЕЧЕНЬ МАТЕРИАЛОВ

                        End If
                End Select
                .asNext()
            End While
            .CloseArticleStructure()
        End With
    End Sub
    Sub read_NEXTLEVELtreeview(myNextNode As TreeNode)
        Application.DoEvents()
        Dim myNode As TreeNode
        For Each myNode In myNextNode.Nodes
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
                TP_Vspom_Mater_Array = Get_Vspom_Mater_Array(ArtID, PRJLINK_ID)
                Try
                    If TP_Vspom_Mater_Array(0, 1) IsNot Nothing Then excel_write_aboutVSPomMater_in_Purchated(TP_Vspom_Mater_Array)
                Catch ex As Exception
                End Try
            End If
            excel_write_about_TreeNode(Positio, ArtID, PRJLINK_ID, Count_Summ, PRJLINK_Param, Art_Param, TP_Array)

            If myNode.Nodes.Count > 0 Then
                read_NEXTLEVELtreeview(myNode)
            Else
                If NO_Purchated Then excel_write_aboutPart_in_Purchated(ArtID, Art_Param, TP_Array, PRJLINK_Param, Count_Summ)
                If NO_Parts Then excel_write_aboutPart_in_PartList(ArtID, PRJLINK_ID, Art_Param, TP_Array, PRJLINK_Param, Count_Summ)
            End If
        Next
    End Sub
    Sub excel_write_aboutPart_in_PartList(ArtID As Integer, PRJLINK_ID As Integer, Art_Info As Array, TC_Info As Array, PRJLINK_Param As Array, Total_Count As String)
        Dim sum, totalsum As Double
        Dim tmp_row As Integer

        Select Case Art_Info(12)
            Case 4
                Dim lastRowNumInPartlist As Integer = Get_LastRowInOneColumn(ShName_PartList, CN_PL_ArtID) + 1
                Try
                    'tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_PL_ArtID, RowN_PL_First, ArtID)
                    'If tmp_row > 0 Then lastRowNumInPartlist = tmp_row

                    set_Value_From_Cell(ShName_PartList, CN_PL_ArtID, lastRowNumInPartlist, ArtID)
                    set_Value_From_Cell(ShName_PartList, CN_PL_Oboz, lastRowNumInPartlist, Art_Info(0))
                    set_Value_From_Cell(ShName_PartList, CN_PL_Naim, lastRowNumInPartlist, Art_Info(1))
                    set_Value_From_Cell(ShName_PartList, CN_PL_PROJ_ID, lastRowNumInPartlist, PRJLINK_Param(0))
                    set_Value_From_Cell(ShName_PartList, CN_PL_Primen9emost, lastRowNumInPartlist, get_OBOZ_by_Part_AID(PRJLINK_Param(0))) 'обозначение применяемого обекта

                    If TC_Info(9) IsNot Nothing Then
                        sum = CDbl(Total_Count) / CDbl(TC_Info(6))
                        totalsum = CDbl(Total_Count)
                    Else
                        sum = CDbl(PRJLINK_Param(2)) '* CDbl(Total_Count)
                        totalsum = CDbl(Total_Count)
                    End If
                    set_Value_From_Cell(ShName_PartList, CN_PL_Count, lastRowNumInPartlist, sum) 'колво

                    set_Value_From_Cell(ShName_PartList, CN_PL_TotalCount, lastRowNumInPartlist, totalsum) 'колво общее

                    'set_Value_From_Cell(ShName_PartList, CN_PL_MU, lastRowNumInPartlist, PRJLINK_Param(11)) 'ед изм
                    set_Value_From_Cell(ShName_PartList, CN_PL_LinkType, lastRowNumInPartlist, PRJLINK_Param(9)) 'тип связи

                    If TC_Info(3) IsNot Nothing Or TC_Info(3) <> "" Then
                        set_Value_From_Cell(ShName_PartList, CN_PL_Sort, lastRowNumInPartlist, TC_Info(3)) 'Сортамент
                        set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, lastRowNumInPartlist, TC_Info(4)) 'ключ имбайз сортамента 
                    Else
                        set_Value_From_Cell(ShName_PartList, CN_PL_Sort, lastRowNumInPartlist, Art_Info(5)) 'Сортамент
                        set_Value_From_Cell(ShName_PartList, CN_PL_Sort_IBKey, lastRowNumInPartlist, Art_Info(6)) 'ключ имбайз сортамента
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

                SetCellsBorderLineStyle2(ShName_PartList, 1, lastRowNumInPartlist, CN_PL_Sort_IBKey, lastRowNumInPartlist, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)

        End Select
    End Sub
    Function get_OBOZ_by_Part_AID(Part_AID As Integer) As String
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
    Sub excel_write_aboutPart_in_Purchated(ArtID As Integer, Art_Info As Array, TC_Info As Array, PRJLINK_Param As Array, Total_Count As String)
        Dim tmp_colors As Color = Color.Khaki
        Dim lastRowNumPurchated As Integer = Get_LastRowInOneColumn(ShName_Purchated, CN_Purchated_Naim) + 1
        Dim tmp_row As Integer
        Dim sum As Double
        Select Case Art_Info(12)
            Case 4 'детали
                Try
                    If TC_Info(4) IsNot Nothing Then
                        tmp_colors = Color.Tan
                        tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, TC_Info(4))
                        If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, TC_Info(3))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, TC_Info(4))
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, TC_Info(10))

                        sum = Math.Round((CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(Total_Count) * CDbl(TC_Info(2))), 3)
                        set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))
                    Else
                        tmp_colors = Color.LightSteelBlue
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
            Case 5, 6, 7 'станд изделия
                Try
                    tmp_colors = Color.LightGreen
                    tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, Art_Info(3))
                    If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, Art_Info(1))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, Art_Info(3))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, PRJLINK_Param(11))

                    sum = (CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(Total_Count))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))
                    'красим вспомогательный материал в свой цвет
                    SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)

                Catch ex As Exception

                End Try
                'Case 6 'прочие

                'Case 7 'материалы

        End Select
        'красим в свой цвет
        SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)
        SetCellsBorderLineStyle2(ShName_Purchated, 1, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)

    End Sub
    Sub excel_write_aboutVSPomMater_in_Purchated(TP_Vspom_Mater_Array As Array)
        Dim lastRowNumPurchated As Integer = Get_LastRowInOneColumn(ShName_Purchated, CN_Purchated_Naim) + 1
        Dim tmp_row As Integer
        Dim sum As Double
        Dim tmp_colors As Color = Color.DarkSalmon
        Try
            If UBound(TP_Vspom_Mater_Array, 1) >= 0 Then
                For q As Integer = 0 To UBound(TP_Vspom_Mater_Array, 1)
                    tmp_row = get_value_bay_FindText_Strong(ShName_Purchated, CN_Purchated_IBKey, RowN_Purchated_First, TP_Vspom_Mater_Array(q, 1))
                    If tmp_row > 0 Then lastRowNumPurchated = tmp_row
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 0))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_IBKey, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 1))
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_MU, lastRowNumPurchated, TP_Vspom_Mater_Array(q, 3))

                    sum = Math.Round((CDbl(get_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated)) + CDbl(TP_Vspom_Mater_Array(q, 2))), 3)
                    set_Value_From_Cell(ShName_Purchated, CN_Purchated_Count, lastRowNumPurchated, sum.ToString.Replace(",", "."))

                    'красим вспомогательный материал в свой цвет
                    SetCellsColor(ShName_Purchated, CN_Purchated_Naim, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, tmp_colors)
                    SetCellsBorderLineStyle2(ShName_Purchated, 1, lastRowNumPurchated, CN_Purchated_MU, lastRowNumPurchated, Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub
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
        set_Value_From_Cell(Sheetname, CN_Purchated, lastRowNum, Art_Param(4))
        set_Value_From_Cell(Sheetname, CN_IBKey_purchated, lastRowNum, Art_Param(3))
        Try
            If Art_Param(2) IsNot Nothing Then set_Value_From_Cell(Sheetname, CN_Mass, lastRowNum, Art_Param(2).ToString.Replace(",", "."))
            set_Value_From_Cell(Sheetname, CN_Mater, lastRowNum, Art_Param(5))
            set_Value_From_Cell(Sheetname, CN_MaterIBKey, lastRowNum, Art_Param(6))
        Catch ex As Exception
        End Try

        set_Value_From_Cell(Sheetname, CN_Kolvo, lastRowNum, PRJLINK_Param(2))
        set_Value_From_Cell(Sheetname, CN_KolvoSbor, lastRowNum, Count_Summ)
        set_Value_From_Cell(Sheetname, CN_EdIzm, lastRowNum, PRJLINK_Param(11))
        set_Value_From_Cell(Sheetname, CN_Type_Article, lastRowNum, Art_Param(7))
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
                    Norma_ras_SUM = ZagSumCount * CDbl(TP_Array(2))
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
                temp_color = Color.Aqua
            Case 2 'комплексы
                temp_color = Color.Aquamarine
            Case 3 'сборки
                temp_color = Color.Yellow
            Case 4 'детали
                temp_color = Color.FromArgb(216, 228, 188) 'Color.DarkKhaki 'GreenYellow (255, 153, 0)
            Case 5 'стандартные
                temp_color = Color.IndianRed
            Case 6 'прочие
                temp_color = Color.LightBlue
            Case 7 'материалы
                temp_color = Color.Khaki
            Case 8 'комплекты
                temp_color = Color.LightPink
        End Select
        SetCellsColor(Sheetname, 1, lastRowNum, last_ColNum, lastRowNum, temp_color)
        If PRJLINK_Param(8) = 2 Then
            temp_color = Color.Gray
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
    Function Get_Vspom_Mater_Array(Art_ID As Integer, Proj_Id As Integer) As Array
        Application.DoEvents()
        Dim array(0, 0), MaterName, MaterIBKey, MaterNorma, MaterMU As String
        Try
            Dim TArt As TPServer.ITArticle = tp.Articles.ByArchCode(Art_ID)
            Dim Tmats As TPServer.ITMaterials = TArt.Materials
            Dim MatCount As Integer = Tmats.Count

            Dim imat As TPServer.ITMaterial
            imat = Tmats.First

            Dim j As Integer = 0
            Try
                For i As Integer = 0 To MatCount - 1
                    If imat.Materials.Count = 0 Then j += 1
                    imat = Tmats.Next
                Next
                j -= 1
            Catch ex As Exception
            End Try
            ReDim array(j, 3)
            Try
                j = 0
                imat = Tmats.First
                For i As Integer = 0 To MatCount - 1
                    If imat.Materials.Count = 0 Then
                        'ReDim Preserve array(j + 1, 3)
                        MaterName = imat.Value("Овсм")
                        MaterIBKey = imat.Value("%MAT")
                        MaterNorma = imat.Norma
                        MaterMU = imat.Value("едНв")

                        array(j, 0) = MaterName
                        array(j, 1) = MaterIBKey
                        array(j, 2) = MaterNorma
                        array(j, 3) = MaterMU
                        j += 1
                    End If
                    imat = Tmats.Next
                Next
            Catch ex As Exception
            End Try
            Return array
        Catch ex As Exception
            Return array
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
                    Dim RefArts As TPServer.ITArticle = Zag.InArts.First
                    For j As Integer = 0 To Zag.InArts.Count - 1
                        'ищем заготовку по применяемости! Если условие не строгое, то нужно дать значение "-1"
                        Try
                            If RefArts.ArchID = Proj_Id Or Proj_Id = -1 Or Proj_Id = 0 Then
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
                        RefArts = Zag.InArts.Next
                    Next
                    Zag = Zags.Next
                Next

                If ReplaceTildaOnSpace Then
                    Sortament = Replace(Sortament, "~", " ")
                End If
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

            Return Array
        End Try
    End Function
    Function Get_Article_Param(Art_ID As Integer) As Array
        Application.DoEvents()
        Dim array(12) As String
        Try
            With s4
                .OpenArticle(Art_ID)
                .ReturnFieldValueWithImbaseKey = 0
                Dim oboz, naim, mass, Imbase_key, purchased, material, materialIMKey, ArtKindName, mass_MU_ID, mass_MU_ID_Str, ArchID, ArchName, ArtKind As String
                oboz = .GetFieldValue_Articles("Обозначение")
                naim = .GetFieldValue_Articles("Наименование")
                mass = .GetArticleMassa
                Imbase_key = .GetFieldValue_Articles("Ключ Imbase")
                purchased = .GetArticlePurchased
                material = .GetArticleMaterial
                If ReplaceTildaOnSpace Then
                    material = Replace(material, "~", " ")
                    naim = Replace(naim, "~", " ")
                End If
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
                Return array
            End With
        Catch ex As Exception

            Return Array
        End Try
    End Function

    Private Sub БезРазделаДокументацияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БезРазделаДокументацияToolStripMenuItem.Click
        БезРазделаДокументацияToolStripMenuItem.Checked = Not (БезРазделаДокументацияToolStripMenuItem.Checked)
        NO_Art_Documentacia = БезРазделаДокументацияToolStripMenuItem.Checked
        CT_ID_in_Query_Change()
    End Sub

    Private Sub БезТехнолическихСвязейToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БезТехнолическихСвязейToolStripMenuItem.Click
        БезТехнолическихСвязейToolStripMenuItem.Checked = Not (БезТехнолическихСвязейToolStripMenuItem.Checked)
        NO_Ru4na9_Sv9z = БезТехнолическихСвязейToolStripMenuItem.Checked
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
    End Sub

    Private Sub СПеречнемМатериаловToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СПеречнемМатериаловToolStripMenuItem.Click
        СПеречнемМатериаловToolStripMenuItem.Checked = Not (СПеречнемМатериаловToolStripMenuItem.Checked)
        NO_Purchated = СПеречнемМатериаловToolStripMenuItem.Checked
    End Sub


    Private Sub ToolStripButton4_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        'Dim values As New List(Of Tuple(Of String, String, String))
        'values.Add(Tuple.Create("a1", "a2", "a2"))
        'values.Add(Tuple.Create("b1", "b2", "a2"))
        'values.Add(Tuple.Create("c1", "c2", "a2"))
        Get_Vspom_Mater_Array(15383, -1)
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