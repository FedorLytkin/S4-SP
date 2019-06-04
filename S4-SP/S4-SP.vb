Imports System
Imports System.Threading
Imports System.IO
Imports System.Text
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports ListView = WinControls.ListView
'Imports 

Public Class Form1
    'индексы столбцов в таблице ВСНРМ
    Public NumPP_col_number As Integer = 1
    Public Oboz_col_number As Integer = 2
    Public Naim_col_number As Integer = 3
    Public KolNaEd_col_number As Integer = 4
    Public Material_col_number As Integer = 8
    Public Tolshina_Diameter_col_number As Integer = 10
    Public Dlina_col_number As Integer = 11
    Public Mass_col_number As Integer = 13
    Public Ed_Ism_col_number As Integer = 18
    Public Prim_col_number As Integer = 19
    Public TypDSE_col_number As Integer = 20
    Public format_col_number As Integer = 21
    Public file_formar_for_doctypes As String
    'индексы столбцов в таблице с материалов
    Public Materials_SheetName_TableMaterialsAndPurchates As String = "Materials"
    Public Purchates_SheetName_TableMaterialsAndPurchates As String = "Purchated articles"
    Public Poln_oboz_VSNRM_col_number As Integer = 1
    Public Sootvetstvie_Imbase_col_number As Integer = 2
    Public IMBASE_Key_col_number As Integer = 3
    Public RazdelSP_col_number As Integer = 4
    'индексы столбцов в таблице с Документация
    Public format_Doc_col_number As Integer = 1
    Public Oboz_Doc_col_number As Integer = 2
    Public Naim_Doc_col_number As Integer = 3

    'параметры опции(настроек)
    Public material_table_path As String ' cсылка на таблицу соответствия материалов
    Public Search_Subdirectories As Boolean
    Public doctypes_array As String
    Public note_list_check_FFm As String
    Public excel_visible As String
    Public Default_work_path As String = "C:\IM\IMWork"
    Public only_Articles As Boolean
    Public Dlina_Componenta_VSNRM As String
    Public ParamSostavVisible As String = get_reesrt_value("ParamSostavVisible", False.ToString)
    Public AddPurchatedArticles As String = get_reesrt_value("AddPurchatedArticles", False.ToString)
    Public PreStartCompareMaterialInVSNRMWithMaterialTable As String = get_reesrt_value("PreStartCompareMaterialInVSNRMWithMaterialTable", False.ToString)
    Public Stop_Process_BOMCreate As String = get_reesrt_value("Stop_Process_BOMCreate", False.ToString)
    Public Stop_Process As Boolean = False
    Public SeparatorObozna4InVSNRMTitle As String = get_reesrt_value("SeparatorObozna4InVSNRMTitle", "  ")
    Public Carto4kaS4SPEnable As Boolean
    Public App_Vers As String = get_reesrt_value("App_Vers", System.Windows.Forms.Application.ProductVersion)
    Public ChekSubAssemly As Boolean

    Public VSNRM_path, Doc_list_path As String
    Public SB_doc_types_file_Format, detal_doc_types_file_Format As String
    Public SB_doc_types, detal_doc_types As Integer
    Public Arch_ID As Integer
    Public ITS4App As S4.TS4App
    Dim search_API As New search_API
    Public DocID, ArtID As Integer
    Dim Doc_list() As String
    Public xlApp As Excel.Application ' = New Microsoft.Office.Interop.Excel.Application()
    Public WB As Excel.Workbook
    Public WS As Excel.Worksheet
    Public WS_DOC As Excel.Worksheet
    Public WS_range As Excel.Range
    Public excFileName As String = "C:\Users\aidarhanov.n.VEZA-SPB\Downloads\таблица.xlsx"
    Public Part_Arr As String = "СПИСОК ДЕТАЛЕЙ"
    Public Docs_arr As String = "ДОКУМЕНТАЦИЯ"
    Public LastInd As Integer
    Public rw As Excel.Range


    'Public DocNames As New ListView.TreeListNode
    Private Sub ВыбратьОбъектToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыбратьОбъектToolStripMenuItem.Click
        select_VSNRM()
    End Sub
    Sub progress_txt_im_ListBox1(progress_txt As String)
        ListBox1.Items.Add(progress_txt)
    End Sub
    Sub VersCange()
        Dim tempApp_Ver As String = System.Windows.Forms.Application.ProductVersion
        If Not chek_reg_exist("App_Vers") Then
            set_reesrt_value("App_Vers", App_Vers)
        End If

        If tempApp_Ver <> App_Vers Then
            Process.Start("http://www.evernote.com/l/AjJ5PRJAb7ZB3raLw5_efcpQ5Zd5Wtz_MWc/")
            set_reesrt_value("App_Vers", tempApp_Ver)
        End If
    End Sub
    Sub select_VSNRM()
        Dim openFileDialog1 As New OpenFileDialog()


        'openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = "excel files (*.xlsx)|*.xls|All files (*.*)|*.*"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                new_doc.ExcelFullFileName = openFileDialog1.FileName
                Open_VSNRM(openFileDialog1.FileName)
            Catch Ex As Exception
                MessageBox.Show("Не удается открыть файл. Код ошибки: " & Ex.Message)
                exc_close()
                xlApp.Quit()
                xlApp = Nothing
                progress_txt_im_ListBox1("Процесс завершился с ошибкой")
            End Try
        End If
    End Sub
    Sub Open_VSNRM(VSNRM_path As String)
        If VSNRM_path IsNot Nothing And VSNRM_path.IndexOf("ВСНРМ") <> 0 And VSNRM_path.IndexOf("xls") <> 0 Then
            progress_txt_im_ListBox1("Подготовка к созданию древовидной структуры")
            xlApp = New Excel.Application()
            xlApp.Visible = False
            WB = xlApp.Workbooks.Open(VSNRM_path)
            WS = WB.Sheets.Item(Part_Arr)

            progress_txt_im_ListBox1("Файл ВСНРМ: " & VSNRM_path)

            Dim str As String = WS.Cells(1, 1).Value()
            Dim naim As String = str.Substring(str.IndexOf(SeparatorObozna4InVSNRMTitle)).Trim(" ")
            Dim oboz As String = str.Substring(0, str.IndexOf(SeparatorObozna4InVSNRMTitle))
            Dim root = New TreeNode(" " & oboz & " - " & naim)

            LastInd = WS.Cells(WS.Rows.Count, Naim_col_number).End(Excel.XlDirection.xlUp).Row
            TreeView1.Nodes.Clear()
            TreeView1.Nodes.Add(root)
            recvAdd_Only_treeview(root, 5, True, LastInd)
            progress_txt_im_ListBox1("Процесс завершен! Древовидная структура построена")
            xlApp.Quit()
            xlApp = Nothing
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = MainFormHeght()
        Me.Width = MainFormWidht()
        tempform = Me
        'VersCange()
        Search_Subdirectories = Convert.ToBoolean(get_reesrt_value("Search_Subdirectories", True.ToString))
        material_table_path = get_reesrt_value("material_table_path", System.Windows.Forms.Application.StartupPath & "\material_with_ImBaseKey.xlsx")
        doctypes_array = get_reesrt_value("doctypes_array", Options.doctypes_array_default)
        note_list_check_FFm = get_reesrt_value("note_list_check_FFm", Options.note_list_check_FFm)
        excel_visible = get_reesrt_value("excel_visible", False.ToString)
        Default_work_path = get_reesrt_value("Default_work_path", Default_work_path)
        Dlina_Componenta_VSNRM = get_reesrt_value("Dlina_Componenta_VSNRM", Options.Dlina_Componenta_VSNRM)
        ParamSostavVisible = get_reesrt_value("ParamSostavVisible", False.ToString)
        SeparatorObozna4InVSNRMTitle = get_reesrt_value("SeparatorObozna4InVSNRMTitle", "  ")
        ChekSubAssemly = get_reesrt_value("ChekSubAssemly", False.ToString)

        AddPurchatedArticles = get_reesrt_value("AddPurchatedArticles", False.ToString)
        PreStartCompareMaterialInVSNRMWithMaterialTable = get_reesrt_value("PreStartCompareMaterialInVSNRMWithMaterialTable", False.ToString)
        Stop_Process_BOMCreate = get_reesrt_value("Stop_Process_BOMCreate", False.ToString)


        Dim err_massage As String
        Try
            ITS4App = CreateObject("S4.TS4App")
            ITS4App.Login()
            ToolStripTextBox1.Text = ITS4App.GetUserFullName_ByUserID(ITS4App.GetUserID)
        Catch ex As Exception
            err_massage = "У вас на компьютере не установлен Search"
            MsgBox(err_massage)
            ListBox1.Items.Add(err_massage)
            СоздатьДеревоToolStripMenuItem.Enabled = False
        End Try
        'проверка на наличие прав отправки запросов sql
        Dim check_right_SQL As String = query_s4("USERS", "fullname", "LOGINNAME", ITS4App.GetUserFullName_ByUserID(ITS4App.GetUserID))
        If check_right_SQL Is Nothing Then
            err_massage = "У вас нет доступа к расширенным функциям API. Обратитесь к Администратору БД, за помощью"
            MsgBox(err_massage)
            ListBox1.Items.Add(err_massage)
            СоздатьДеревоToolStripMenuItem.Enabled = False
            Carto4kaS4SPEnable = False
            ЖурналОТДToolStripMenuItem.Enabled = False
        End If
        C()
        'Chek_licenc()
        'VSNRM2.ShowDialog()
    End Sub
    Function MainFormHeght() As Double
        Dim FormHeght As Double = get_reesrt_value("MainFormHeght", Me.Height)
        If FormHeght = 0 Then
            set_reesrt_value("MainFormHeght", Me.Height)
            FormHeght = get_reesrt_value("MainFormHeght", Me.Height)
        End If
        Return FormHeght
    End Function
    Function MainFormWidht() As Double
        Dim FormWidht As Double = get_reesrt_value("MainFormWidht", Me.Width)
        If FormWidht = 0 Then
            set_reesrt_value("MainFormWidht", Me.Width)
            FormWidht = get_reesrt_value("MainFormWidht", Me.Width)
        End If
        Return FormWidht
    End Function
    Public Sub MainFormSize()
        set_reesrt_value("MainFormHeght", Me.Height)
        set_reesrt_value("MainFormWidht", Me.Width)
    End Sub
    Public Sub createDocs_in_Documents()
        Dim LastInd_InDoc = WS_DOC.Cells(WS_DOC.Rows.Count, Naim_Doc_col_number).End(Excel.XlDirection.xlUp).Row
        For q As Integer = 3 To LastInd_InDoc
            add_document_InDoc(q)
        Next
    End Sub
    Public Sub start_create_BOM()
        progress_txt_im_ListBox1("Подготовка к созданию древовидной структуры")
        progress_txt_im_ListBox1("Файл ВСНРМ: " & VSNRM_path)
        progress_txt_im_ListBox1("Архив: " & new_doc.TextBox3.Text)
        If only_Articles = False Then progress_txt_im_ListBox1("Папка с документацией: " & Doc_list_path)
        If Convert.ToBoolean(PreStartCompareMaterialInVSNRMWithMaterialTable) Then
            progress_txt_im_ListBox1("Идет сверка содержимого ВСНРМ с Таблицей соответствия материалов и покупных изделий")
            CompareMaterialInVSNRMWithMaterialTable()
            If Convert.ToBoolean(Stop_Process_BOMCreate) = True And Stop_Process = True Then
                progress_txt_im_ListBox1("Процесс созданию древовидной структуры экстренно завершен! При сверке обнаружились недостающие элементы")
                Exit Sub
            End If
        End If

        Try
            xlApp = New Excel.Application()
            xlApp.Visible = excel_visible
            WB = xlApp.Workbooks.Open(VSNRM_path)
            WS = WB.Sheets.Item(Part_Arr)
            Try
                WS_DOC = WB.Sheets.Item(Docs_arr)
            Catch ex As Exception
                'MessageBox.Show("В ВСНРМ нет листа " & Docs_arr & vbNewLine & "Продолжить?", ,, vbYesNo)
                Dim result As Integer = MessageBox.Show("В ВСНРМ нет листа " & Docs_arr & vbNewLine & "Продолжить?", "S4-SP", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    progress_txt_im_ListBox1("В ВСНРМ не найден лист " & Docs_arr & "! Остановка процедуры")
                    excel_docs_close()
                    exc_close()
                    Exit Sub
                ElseIf result = DialogResult.Yes Then
                    progress_txt_im_ListBox1("В ВСНРМ не найден лист " & Docs_arr & "! Продолжение процедуры")
                    Exit Try
                End If
            End Try
            Open_EX_Doc(False, material_table_path)

            Dim str As String = WS.Cells(1, 1).Value()
            Dim naim As String = str.Substring(str.IndexOf(SeparatorObozna4InVSNRMTitle)).Trim(" ")
            Dim oboz As String = str.Substring(0, str.IndexOf(SeparatorObozna4InVSNRMTitle))
            Dim root = New TreeNode(" " & oboz & " - " & naim)

            'запустить создание объекта в с4 (Сборочная единица!). Вернет ее PartAID (понадобится следующем уровне)
            'PartAID записать в тегах текущей ветки
            Dim MainPartID As Integer = ITS4App.AddNewArticle2(oboz.Replace("СБ", ""), "", naim, 3)
            root.Tag = MainPartID
            If root.Tag > 0 Then
                progress_txt_im_ListBox1("Создан объект: ArtID:" & MainPartID & " * " & oboz & " - " & naim)
            End If
            If only_Articles = False Then
                Dim doc_type As String = doctypes(5)
                Dim doc_filename As String = Doc_list_path & "\" & oboz & " " & naim & "." & fileformat_for_doctypes(5)
                'Dim doc_type As Integer = query_s4("doctypes", "DOC_NAME", "DOC_TYPE", detal_doc_types)
                'If File.Exists(doc_filename) Then
                '    DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, SB_doc_types, Arch_ID, oboz, naim)
                'End If
                If Not File.Exists(doc_filename) Then
                    doc_filename = find_file_by_ObozNainFformat(oboz, naim, fileformat_for_doctypes(5))
                    If doc_filename Is Nothing Then doc_filename = find_file_by_ObozNainFformat_In_PDF(oboz, naim, fileformat_for_doctypes(5))
                    If doc_filename IsNot Nothing Then
                        Dim find_txt_inSQL_Quere As String = find_txt_inSQL_Quere_function(WS.Cells(5, TypDSE_col_number).value)
                        doc_type = query_s4_2_answer("doctypes", "doc_ext", "dt_name", "doc_type", "pdf", find_txt_inSQL_Quere)
                        'ITS4App.OpenQuery("select * from doctypes where doc_ext = 'pdf' And dt_name = '" & WS.Cells(5, TypDSE_col_number).value & "'")
                        'doc_type = ITS4App.QueryFieldByName("doc_type")
                        'ITS4App.CloseQuery()
                    End If
                End If
                DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, doc_type, Arch_ID, oboz, naim)
                If DocID > 0 Then
                    progress_txt_im_ListBox1("Создан документ: DocID:" & DocID & " * " & oboz & " - " & naim)
                    create_New_Version_InDoc(doc_filename, DocID)
                End If
            End If
            Call add_Bom_Node_In_Documentacia2(oboz, MainPartID)
            If WS_DOC IsNot Nothing Then
                'тут будет процедуры создающая документы из листа ДОКУМЕНТАЦИЯ в серч
                Documentacia(3, WS_DOC.Cells(WS_DOC.Rows.Count, Naim_Doc_col_number).End(Excel.XlDirection.xlUp).Row, MainPartID)
                'MsgBox("Documentacia")
            End If
            TreeView1.Nodes.Clear()
            TreeView1.Nodes.Add(root)
            LastInd = WS.Cells(WS.Rows.Count, Naim_col_number).End(Excel.XlDirection.xlUp).Row

            'прогрессбар появление 
            'Dim SelectedCount As Integer = LastInd - 5
            'Dim textProgress As String = "Обработано 0 из " & SelectedCount
            ShowProgressBarForm(LastInd - 4, "Создание Дерево состава и Занесение в архив", "Занесение в архив документов из ВСНРМ")

            recvAdd(root, 5, True, LastInd)
            ITS4App.CloseProgressBarForm()
            If only_Articles = False Then add_SP_InBOM(root.Tag) 'процедура по созданию документов SP на сборочные единицы

            progress_txt_im_ListBox1("Процесс завершен! Создано изделие: ArtID: " & root.Tag & " * " & oboz & " - " & naim)
            excel_docs_close()
        Catch ex As Exception
            MsgBox("Не удалось запустить процедуру! Описание ошибки: " & ex.Message)
            progress_txt_im_ListBox1("Процесс прерван! Ошибка! " & ex.Message)
            excel_docs_close()
            ITS4App.CloseProgressBarForm()
        End Try
    End Sub

    Public Sub Documentacia(rownum As Integer, totalNum As Integer, ProjAID As Integer)
        progress_txt_im_ListBox1("Начало обработки данных в листе ДОКУМЕНТАЦИЯ. Всего позиций: " & (totalNum + 1) - rownum)
        Try
            For i As Integer = rownum To totalNum
                Dim Doc_Format, Doc_Oboz, Doc_Naim As String
                Dim ArtIDinDoc, PRJLINK_ID, razdel As Integer : razdel = 1
                Doc_Format = WS_DOC.Cells(i, format_Doc_col_number).value
                Try
                    Doc_Format = Doc_Format.Replace("А", "A") ' замена киррилицы на латиницу
                Catch ex As Exception

                End Try
                Doc_Oboz = WS_DOC.Cells(i, Oboz_Doc_col_number).value
                Doc_Naim = WS_DOC.Cells(i, Naim_Doc_col_number).value
                With ITS4App
                    Try
                        ArtIDinDoc = .AddNewArticle2(Doc_Oboz, "", Doc_Naim, razdel)
                        progress_txt_im_ListBox1("Создан объект: ArtID: " & ArtIDinDoc & " * " & Doc_Oboz & " - " & Doc_Naim)
                    Catch ex As Exception
                        progress_txt_im_ListBox1("Процесс прерван! Ошибка! " & ex.Message)
                    End Try

                    Try
                        PRJLINK_ID = .AddBOMItem(ProjAID, ArtIDinDoc, 0, 0, razdel, "", "")
                        .OpenArticle(ProjAID)
                        progress_txt_im_ListBox1("Объект: ArtID: " & ArtIDinDoc & " * " & Doc_Oboz & " - " & Doc_Naim & ", добавлен в состав объекта ProjAID: " & ProjAID & " * " & .GetArticleDesignation & " - " & .GetArticleName)
                        .CloseArticle()
                        .OpenBOMItem(PRJLINK_ID)
                        If .GetFieldValue_BOM("Формат") <> "" Then
                            .SetFieldValue_BOM("Формат", Doc_Format)
                            .AddArticleUserEventToLog(ProjAID, 1, "Изменен формат с '" & .GetFieldValue_BOM("Формат") & "' на '" & Doc_Format & "'")
                        End If
                        .CloseBOMItem()
                    Catch ex As Exception
                        progress_txt_im_ListBox1("Процесс прерван! Ошибка! " & ex.Message)
                    End Try
                End With
            Next
        Catch ex As Exception
            progress_txt_im_ListBox1("Процесс прерван! Ошибка! " & ex.Message)
        End Try
        progress_txt_im_ListBox1("Завершение обработки данных в листе ДОКУМЕНТАЦИЯ")
    End Sub
    Function find_txt_inSQL_Quere_function(typeDSEInVSNRM As String)
        Select Case typeDSEInVSNRM
            Case "Деталь"
                find_txt_inSQL_Quere_function = "Чертеж"
            Case "Сборочная единица"
                find_txt_inSQL_Quere_function = "Сборочный чертеж"
        End Select
    End Function
    Sub excel_docs_close()
        'закрытие excel Документов
        xlApp.Quit()
        xlApp = Nothing
        exc_close()
    End Sub
    Public Sub ShowProgressBarForm(SelectedCount As Integer, title_message As String, description_text As String)
        Dim textProgress As String = "Обработано 0 из " & SelectedCount
        ITS4App.ShowProgressBarForm("Создание Дерева состава и Занесение в архив", "Занесение в архив документов из ВСНРМ", textProgress, SelectedCount)
    End Sub
    Public Sub SetProgressBarData_and_CheckUserBreak(partsComplete As Integer, SelectedCount As Integer)
        Dim textProgress As String = "Обработано " & partsComplete & " из " & SelectedCount
        Dim UserBreak As Integer = ITS4App.SetProgressBarData_and_CheckUserBreak("", textProgress, partsComplete)
        If UserBreak > 0 Then
            ITS4App.CloseProgressBarForm()
        End If
    End Sub
    Private Function recvAdd(node As TreeNode, rownum As Integer, isFirst As Boolean, totalNum As Integer)
        'прогресс бар заполнение 
        SetProgressBarData_and_CheckUserBreak(rownum - 5, LastInd - 4)

        Dim numpp As String = WS.Cells(rownum, NumPP_col_number).Value()

        Dim ptrn1
        If isFirst Then
            ptrn1 = New Regex("^.+")
        Else
            ptrn1 = New Regex("^" + node.Text.Substring(0, node.Text.IndexOf(" ")) + ".+")
        End If
        While (rownum <= totalNum) ' (numpp IsNot Nothing)
            If (numpp Is Nothing) Then
                numpp = node.Text.Substring(0, node.Text.IndexOf(" ")) + ".1"
            End If
            If Not ptrn1.IsMatch(numpp) Then
                Exit While
            End If

            Dim part_type As String = WS.Cells(rownum, TypDSE_col_number).value

            'здесь будет процедура, которая будет проверять существование сборочной единицы в архиве, и если такой объет есть, будет пропускать его
            If part_type = "Сборочная единица" And check_exist_SB_in_S4(WS.Cells(rownum, Oboz_col_number).Value()) And ChekSubAssemly Then
                Exit While
            End If

            'запустить создание объекта в с4 (Тип объекта брать из ТИПА ДСЕ). Вернет ее PartAID (понадобится следующем уровне)
            'Dim chekc_art_exist = ITS4App.GetArtID_ByDesignation(WS.Cells(rownum, Oboz_col_number).value)

            Dim newNode As New TreeNode(numpp & " " & WS.Cells(rownum, Oboz_col_number).Value() & " - " & WS.Cells(rownum, Naim_col_number).Value())
            If (part_type = "Деталь" Or part_type = "Сборочная единица") Then
                newNode.Tag = add_article(rownum)
                If only_Articles = False Then DocID = add_document(rownum)
                Call add_Bom_Node(rownum, node.Tag, newNode.Tag)
                'запустить создание ветки в с4 AddBOMItem
                'PartAID записать в тегах текущей ветки 

                'записать в раздел документация сборочный чертеж
                If part_type = "Сборочная единица" Then
                    Call add_Bom_Node_In_Documentacia(rownum, newNode.Tag)
                End If
            End If
            If part_type = "Покупное изделие" Then
                If Convert.ToBoolean(AddPurchatedArticles) = True Then
                    newNode.Tag = add_materialInBOM_S4(rownum)
                    Call add_Bom_Node(rownum, node.Tag, newNode.Tag)
                End If
            End If
            node.Nodes.Add(newNode)

            rownum += 1
            'прогресс бар заполнение 
            SetProgressBarData_and_CheckUserBreak(rownum - 4, LastInd - 4)

            Dim nextnumpp As String = WS.Cells(rownum, NumPP_col_number).Value()
            If nextnumpp Is Nothing Then
                If (isFirst) Then
                    nextnumpp = CInt(newNode.Text) + 1
                Else
                    Dim prevNum As Integer = CInt(Regex.Match(newNode.Text, "[0-9]+$").Value)
                    nextnumpp = node.Text.Substring(0, node.Text.IndexOf(" ")) + "." + CStr(prevNum + 1)
                End If
                'Exit While
            End If
            On Error Resume Next

            Dim ptrn2 As New Regex("^" + numpp + "\.[0-9]+$")
            If ptrn2.IsMatch(nextnumpp) Then
                rownum = recvAdd(newNode, rownum, False, totalNum)
                numpp = WS.Cells(rownum, NumPP_col_number).Value()
            Else
                numpp = nextnumpp
            End If
        End While
        Return rownum
    End Function
    Function check_exist_SB_in_S4(Designatio As String) As Boolean
        With ITS4App
            Dim search_Art_ID As Integer = .GetArtID_ByDesignation(Designatio)
            If search_Art_ID > 0 Then Return True Else Return False
        End With
    End Function
    Private Function recvAdd_Only_treeview(node As TreeNode, rownum As Integer, isFirst As Boolean, totalNum As Integer)
        Dim numpp As String = WS.Cells(rownum, NumPP_col_number).Value()

        Dim ptrn1
        If isFirst Then
            ptrn1 = New Regex("^.+")
        Else
            ptrn1 = New Regex("^" + node.Text.Substring(0, node.Text.IndexOf(" ")) + ".+")
        End If
        While (rownum <= totalNum)
            If (numpp Is Nothing) Then
                numpp = node.Text.Substring(0, node.Text.IndexOf(" ")) + ".1"
            End If
            If Not ptrn1.IsMatch(numpp) Then
                Exit While
            End If
            'Dim newNode As New TreeNode(numpp & " " & WS.Cells(rownum, 2).Value() & " - " & WS.Cells(rownum, 3).Value())
            Dim newNode As New TreeNode(numpp & " " & WS.Cells(rownum, 2).Value() & " - " & WS.Cells(rownum, 3).Value())
            node.Nodes.Add(newNode)
            rownum += 1
            Dim nextnumpp As String = WS.Cells(rownum, 1).Value()
            If (nextnumpp Is Nothing) Then
                If (isFirst) Then
                    nextnumpp = CInt(newNode.Text) + 1
                Else
                    Dim prevNum As Integer = CInt(Regex.Match(newNode.Text, "[0-9]+$").Value)
                    nextnumpp = node.Text.Substring(0, node.Text.IndexOf(" ")) + "." + CStr(prevNum + 1)
                End If
            End If
            On Error Resume Next
            Dim ptrn2 As New Regex("^" + numpp + "\.[0-9]+$")
            If ptrn2.IsMatch(nextnumpp) Then
                rownum = recvAdd_Only_treeview(newNode, rownum, False, totalNum)
                numpp = WS.Cells(rownum, 1).Value()
            Else
                numpp = nextnumpp
            End If
            If Regex.IsMatch(numpp, "3.3.6$") Then
                totalNum += 1
                totalNum -= 1
            End If
        End While
        Return rownum
    End Function
    Function add_materialInBOM_S4(rownum As Integer) As Integer
        Dim NaimVSNRM As String = WS.Cells(rownum, Naim_col_number).value ' "Сталь EN~10088-4~- X2CrNiMoN22-5-3&^i6000345087E14000003" '= WS.Cells(rownum, Naim_col_number).value
        Dim ReSpomsePolnoeNaimByPauchatedTable As String = PurchatedAndMaterial_with_IMBKey(NaimVSNRM) 'сюда передается строка в виде "ПОлное обозначение" &^ "Ключ ИМбэйз" &^ "Раздел СП". нужен парсер(сплитер)
        Dim Naim, ImbaseKeys As String
        Dim new_ArtID, Razdel As Integer
        Dim delimiter As Char = "&"

        Dim ReSpomsePolnoeNaimByPauchatedTable_Array() As String = ReSpomsePolnoeNaimByPauchatedTable.Split(delimiter)
        Naim = ReSpomsePolnoeNaimByPauchatedTable_Array(0)
        ImbaseKeys = ReSpomsePolnoeNaimByPauchatedTable_Array(1)
        Razdel = ReSpomsePolnoeNaimByPauchatedTable_Array(2)
        If Naim Is Nothing Or ImbaseKeys Is Nothing Or Razdel = 0 Then
            MsgBox("Покупное изделие " & NaimVSNRM & " не может быть создано!" & vbNewLine & "В таблице соответсвия для него не заполнены параметры: Полное обозначение; Ключ IMBASE; Раздел СП",, "Ошибка!!!")
            Exit Function
        End If
        new_ArtID = ITS4App.AddNewArticle2("", "", Naim, Razdel)
        ITS4App.OpenArticle(new_ArtID)
        ITS4App.SetFieldValue_Articles("Ключ Imbase", ImbaseKeys)
        ITS4App.AddArticleUserEventToLog(new_ArtID, 1, "Изменен Ключ Imbase на '" & ImbaseKeys & "'")
        If new_ArtID > 0 Then progress_txt_im_ListBox1("Создан объект: ArtID: " & new_ArtID & " * " & Naim & " - Ключ IMBase: " & ImbaseKeys)
        Return new_ArtID
    End Function
    Function add_article(rownum As Integer) As Integer
        Dim designation As String = WS.Cells(rownum, Oboz_col_number).value.ToString.Replace("СБ", "")
        Dim Naim As String = WS.Cells(rownum, Naim_col_number).value
        Dim material As String = WS.Cells(rownum, Material_col_number).value
        Dim tolshina_diameter As String = WS.Cells(rownum, Tolshina_Diameter_col_number).value
        Dim dlina As String = WS.Cells(rownum, Dlina_col_number).value
        Dim mass As String = WS.Cells(rownum, Mass_col_number).value

        'ITS4App.OpenQuery("select * from ssections where art_kind = '" & WS.Cells(rownum, TypDSE_col_number).Value() & "'")
        Dim Razdel As Integer '= UCase(ITS4App.QueryFieldByName("SECTION_ID"))
        'ITS4App.CloseQuery()
        Razdel = query_s4("ssections", "art_kind", "SECTION_ID", WS.Cells(rownum, TypDSE_col_number).Value())
        Dim new_ArtID As Integer
        If only_Articles Then GoTo ifDocID_is_null
        If designation.IndexOf("-") > 0 Then
            Dim isp As String = designation.Substring(designation.LastIndexOf("-") + 1)
            Dim docid As Integer = ITS4App.GetDocID_ByDesignation(designation.Substring(0, designation.LastIndexOf("-")))
            If docid > 0 Then
                ITS4App.OpenDocument(docid)
                ITS4App.CheckOut()
                new_ArtID = ITS4App.AddArticleIsp(0, docid, ITS4App.GetFieldValue("Обозначение") & "-" & isp, ITS4App.GetFieldValue("Наименование"), "", isp, mass, 2)
                ITS4App.CheckIn()
            Else
                'если исполнение встретилось раньше базовой версии файла, эта фунция найдет в ВСНРМ базовое исполнение, создаст на нее документ
                If add_document_on_Isponenie(designation) > 0 Then
                    docid = ITS4App.GetDocID_ByDesignation(designation.Substring(0, designation.LastIndexOf("-")))
                    ITS4App.OpenDocument(docid)
                    ITS4App.CheckOut()
                    new_ArtID = ITS4App.AddArticleIsp(0, docid, ITS4App.GetFieldValue("Обозначение") & "-" & isp, ITS4App.GetFieldValue("Наименование"), "", isp, mass, 2)
                    ITS4App.CheckIn()
                Else
                    GoTo ifDocID_is_null
                End If
            End If
        Else
ifDocID_is_null:
            new_ArtID = ITS4App.AddNewArticle2(designation, "", Naim, Razdel)

        End If

        If new_ArtID = -1 Then
            new_ArtID = ITS4App.GetArtID_ByDesignation(designation)
            Return new_ArtID
            progress_txt_im_ListBox1("Объект: ArtID: " & new_ArtID & " * " & designation & " - " & Naim & " уже существует в S4")
            Exit Function
        End If
        ITS4App.OpenArticle(new_ArtID)
        If material IsNot Nothing Then
            Dim material_IMB As String = material_with_IMBKey(material)
            Call ITS4App.SetFieldValue_Articles("Материал", material_IMB)
            ITS4App.AddArticleUserEventToLog(new_ArtID, ITS4App.ErrorCode(), "Изменен Материал с '" & ITS4App.GetArticleMaterial & "' на '" & material_IMB & "'")
        End If
        If tolshina_diameter IsNot Nothing Then
            Call ITS4App.SetFieldValue_Articles("Толщина/Диаметр", tolshina_diameter)
        End If
        If dlina IsNot Nothing And Razdel <> 3 Then
            Call ITS4App.SetFieldValue_Articles("Примечание", Dlina_Componenta_VSNRM & "=" & dlina)
            ITS4App.AddArticleUserEventToLog(new_ArtID, ITS4App.ErrorCode(), "Примечание" & Dlina_Componenta_VSNRM & "=" & dlina)
        End If
        Call ITS4App.SetFieldValue_Articles("Масса", mass)
        progress_txt_im_ListBox1("Создан объект: ArtID: " & new_ArtID & " * " & designation & " - " & Naim)
        ITS4App.AddArticleUserEventToLog(new_ArtID, ITS4App.ErrorCode(), "Создан объект: ArtID: " & new_ArtID & " * " & designation & " - " & Naim)
        Return new_ArtID
    End Function
    Function add_document(rownum As Integer) As Integer
        'Dim note_list_check_FFm As String = "изм.; зам."
        Dim designation As String = WS.Cells(rownum, Oboz_col_number).value
        Dim Naim As String = WS.Cells(rownum, Naim_col_number).value
        Dim fileformat As String = fileformat_for_doctypes(rownum)
        Dim doc_filename As String = Doc_list_path & "\" & designation & " " & Naim & "." & fileformat

        Dim doc_type As Integer = doctypes(rownum)

        If Not File.Exists(doc_filename) Then
            doc_filename = find_file_by_ObozNainFformat(designation, Naim, fileformat)
            If doc_filename Is Nothing Then
                doc_filename = find_file_by_ObozNainFformat_In_PDF(designation, Naim, fileformat)
                If doc_filename IsNot Nothing Then
                    Dim find_txt_inSQL_Quere As String = find_txt_inSQL_Quere_function(WS.Cells(rownum, TypDSE_col_number).value)
                    doc_type = query_s4_2_answer("doctypes", "doc_ext", "dt_name", "doc_type", "pdf", find_txt_inSQL_Quere)
                    'ITS4App.OpenQuery("select * from doctypes where doc_ext = 'pdf' And dt_name = '" & WS.Cells(rownum, TypDSE_col_number).value & "'")
                    'doc_type = ITS4App.QueryFieldByName("doc_type")
                    'ITS4App.CloseQuery()
                End If
            End If
        End If
        DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, doc_type, Arch_ID, designation, Naim)
        'формат страницы документа заполнится тут
        If DocID > 0 Then
            Dim format_doc As String = WS.Cells(rownum, format_col_number).value
            format_doc = UCase(format_doc).Replace("А", "A") 'проверка и исправление на маленькие буквы и замена с русского на анл "A" в ячейке Формат
            If format_doc IsNot Nothing Then
                ITS4App.OpenDocument(DocID)
                ITS4App.SetFieldValue("Формат", format_doc)
                ITS4App.AddDocumentUserEventToLog(DocID, ITS4App.ErrorCode(), -1, "Формат изменен с '" & ITS4App.GetFieldValue("Формат") & "' на '" & format_doc & "'")
                ITS4App.CloseDocument()
            End If
            progress_txt_im_ListBox1("Создан документ: DocID: " & DocID & " * " & designation & " - " & Naim)
            'здесь будет процедура по созданию/изменению версии документа 
            create_New_Version_InDoc(doc_filename, DocID)
        End If

        ''приведенный ниже код создает документ спецификации. НО! он не работает, потому что после его создания невозможна правка состава изделия
        ''создание спецификаций делается в самом конце
        'If WS.Cells(rownum, TypDSE_col_number).value = "Сборочная единица" Then
        '    ITS4App.CreateNewDocument(Arch_ID, 1, designation.Replace("СБ", ""), Naim, 0)
        'End If
        Return DocID
    End Function

    Sub create_New_Version_InDoc(doc_filename As String, docId As String)
        For Each note_check_FFm As String In note_list_check_FFm.Split(";")
            note_check_FFm = note_check_FFm.Trim(" ")
            If doc_filename.IndexOf("(" & note_check_FFm) > 0 Then
                Dim new_vers As String = doc_filename.Substring(doc_filename.IndexOf(note_check_FFm) + note_check_FFm.Length)
                new_vers = new_vers.Substring(0, new_vers.LastIndexOf(")"))
                ITS4App.OpenDocument(docId)
                Dim old_vers As String = ITS4App.GetFieldValue("Изменение")
                ITS4App.SetDocVersionCode(new_vers)
                ITS4App.CloseDocument()
                progress_txt_im_ListBox1("В док-те DocId: " & docId & " * " & " изменено Изм. с " & old_vers & " на " & note_check_FFm & new_vers)
                Exit For
            End If
        Next
    End Sub
    Function add_document_InDoc(rownum As Integer) As Integer
        Dim note_list_check_FFm As String = "изм.; зам."
        Dim designation As String = WS_DOC.Cells(rownum, Oboz_Doc_col_number).value
        Dim Naim As String = WS_DOC.Cells(rownum, Naim_Doc_col_number).value
        Dim fileformat As String = fileformat_for_doctypes(rownum)
        Dim doc_filename As String = Doc_list_path & "\" & designation & " " & Naim & "." & fileformat

        Dim doc_type As Integer = doctypes(rownum)

        If Not File.Exists(doc_filename) Then
            doc_filename = find_file_by_ObozNainFformat(designation, Naim, fileformat)
            If doc_filename Is Nothing Then doc_filename = find_file_by_ObozNainFformat_In_PDF(designation, Naim, fileformat)
            If doc_filename IsNot Nothing Then
                Dim find_txt_inSQL_Quere As String = find_txt_inSQL_Quere_function(WS.Cells(rownum, TypDSE_col_number).value)
                doc_type = query_s4_2_answer("doctypes", "doc_ext", "dt_name", "doc_type", "pdf", find_txt_inSQL_Quere)
                'ITS4App.OpenQuery("select * from doctypes where doc_ext = 'pdf' And dt_name = '" & WS.Cells(rownum, TypDSE_col_number).value & "'")
                'doc_type = ITS4App.QueryFieldByName("doc_type")
                'ITS4App.CloseQuery()
            End If
        End If
            DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, doc_type, Arch_ID, designation, Naim)
        'формат страницы документа заполнится тут
        If DocID > 0 Then
            Dim format_doc As String = WS_DOC.Cells(rownum, format_Doc_col_number).value
            If format_doc IsNot Nothing Then
                ITS4App.OpenDocument(DocID)
                ITS4App.SetFieldValue("Формат", format_doc)
                ITS4App.AddDocumentUserEventToLog(DocID, ITS4App.ErrorCode(), -1, "Формат изменен с '" & ITS4App.GetFieldValue("Формат") & "' на '" & format_doc & "'")
                ITS4App.CloseDocument()
            End If
            'здесь будет процедура по созданию/изменению версии документа 
            For Each note_check_FFm As String In note_list_check_FFm.Trim(";")
                If doc_filename.IndexOf(note_check_FFm) > 0 Then
                    Dim new_vers As String = doc_filename.Substring(doc_filename.IndexOf(note_check_FFm) + note_check_FFm.Length)
                    new_vers = new_vers.Substring(0, new_vers.LastIndexOf(")"))
                    ITS4App.OpenDocument(DocID)
                    ITS4App.SetDocVersionCode(new_vers)
                    ITS4App.CloseDocument()
                End If
            Next
        End If

        ''приведенный ниже код создает документ спецификации. НО! он не работает, потому что после его создания невозможна правка состава изделия
        'If WS.Cells(rownum, TypDSE_col_number).value = "Сборочная единица" Then
        '    ITS4App.CreateNewDocument(Arch_ID, 1, designation.Replace("СБ", ""), Naim, 0)
        'End If
        Return DocID
    End Function
    Function add_document_on_Isponenie(Full_designatio As String)
        Dim designatio As String = Full_designatio.Substring(0, Full_designatio.LastIndexOf("-"))
        Dim search_rownum As Integer = get_value_bay_FindText_Strong(Part_Arr, Oboz_col_number, 5, designatio)
        If search_rownum > 5 Then
            add_document_on_Isponenie = add_document(search_rownum)
            With ITS4App
                Dim ArtID_Isp As Integer = .GetArtID_ByDocID(add_document_on_Isponenie)
                Dim massInVSNRM As String = WS.Cells(Mass_col_number, search_rownum).value
                Dim materialIB As String = material_with_IMBKey(WS.Cells(Material_col_number, search_rownum).value)
                .OpenArticle(ArtID_Isp)
                .SetFieldValue_Articles("Масса", massInVSNRM)
                .AddArticleUserEventToLog(ArtID_Isp, .ErrorCode(), "Масса изменена с '" & .GetArticleMassa & "' на '" & massInVSNRM & "'")
                .SetFieldValue_Articles("Толщина/Диаметр", WS.Cells(Tolshina_Diameter_col_number, search_rownum).value)
                .SetFieldValue_Articles("Материал", materialIB)
                .AddArticleUserEventToLog(ArtID_Isp, .ErrorCode(), "Материал изменен с '" & .GetArticleMaterial & "' на '" & materialIB & "'")
                progress_txt_im_ListBox1("Создан Объект: ArtID:" & .GetArtID_ByDocID(add_document_on_Isponenie) & " * " & designatio & " - " & .GetFieldValue_Articles("Наименование"))
            End With
        End If
    End Function
    Public Function get_value_bay_FindText_Strong(SheetName As String, columnIndex As Integer, row_index As Integer, FindText As String)
        Dim ActiveSheet As Worksheet = WS
        Dim lastrow As Integer = LastInd
        Try
            With ActiveSheet
                While row_index < lastrow
                    If .Cells(row_index, columnIndex).Value = FindText Then
                        Return row_index
                        Exit While
                    End If
                    row_index = row_index + 1
                End While
            End With
        Catch ex As Exception

        End Try
    End Function
    Function find_file_by_ObozNainFformat(designation As String, Naim As String, doc_filename As String)
        Dim search_option As SearchOption = SearchOption.TopDirectoryOnly
        If Search_Subdirectories Then search_option = SearchOption.AllDirectories
        Dim lik_oboz, lik_naim As Boolean

        Dim file_list() As String = System.IO.Directory.GetFiles(Doc_list_path, "*." & doc_filename, search_option)
        For Each file As String In file_list
            lik_oboz = file Like "*" & designation & "*"
            lik_naim = file Like "*" & Naim & "*"
            If lik_oboz And lik_naim Then
                Return file
                Exit Function
            End If
        Next
    End Function
    Function find_file_by_ObozNainFformat_In_PDF(designation As String, Naim As String, doc_filename As String)
        Dim search_option As SearchOption = SearchOption.TopDirectoryOnly
        If Search_Subdirectories Then search_option = SearchOption.AllDirectories
        Dim lik_oboz, lik_naim As Boolean

        Dim file_list() As String = System.IO.Directory.GetFiles(Doc_list_path, "*." & "PDF", search_option)
        For Each file As String In file_list
            lik_oboz = file Like "*" & designation & "*"
            lik_naim = file Like "*" & Naim & "*"
            If lik_oboz And lik_naim Then
                Return file
                Exit Function
            End If
        Next
    End Function
    Function doctypes(rownum As Integer)
        'не эффективно каждый раз посылать sql-запрос, отправить один запрос в начале и переменные записать сюда
        'РЕШЕНО! Запрос отправляется в при сохранении параметров в форме New_doc
        Select Case WS.Cells(rownum, TypDSE_col_number).value
            Case "Деталь"
                doctypes = detal_doc_types
            Case "Сборочная единица"
                doctypes = SB_doc_types
        End Select
    End Function
    Function doctypes_When_PDF_Find(rownum As Integer)
        Select Case WS.Cells(rownum, TypDSE_col_number).value
            Case "Деталь"
                doctypes_When_PDF_Find = detal_doc_types
            Case "Сборочная единица"
                doctypes_When_PDF_Find = SB_doc_types
        End Select
    End Function
    Function fileformat_for_doctypes(rownum As Integer)
        Select Case WS.Cells(rownum, TypDSE_col_number).value
            Case "Деталь"
                fileformat_for_doctypes = detal_doc_types_file_Format
            Case "Сборочная единица"
                fileformat_for_doctypes = SB_doc_types_file_Format
        End Select
    End Function
    Public Function material_with_IMBKey(material_VSNRM As String) As String
        Dim material_imbase, imb_kase As String
        Dim sheetName As String = Materials_SheetName_TableMaterialsAndPurchates
        Dim material_imbase_index As Integer
        material_imbase_index = get_value_bay_FindText(sheetName, "A2", "A" & Get_LastRowInOneColumn(sheetName, Poln_oboz_VSNRM_col_number), material_VSNRM)
        If material_imbase_index <> 0 Then
            material_imbase = get_Value_From_Cell(sheetName, Sootvetstvie_Imbase_col_number, material_imbase_index)
            imb_kase = get_Value_From_Cell(sheetName, IMBASE_Key_col_number, material_imbase_index)
            material_with_IMBKey = material_imbase & "&^" & imb_kase
        Else
            set_Value_From_Cell(sheetName, Poln_oboz_VSNRM_col_number, Get_LastRowInOneColumn(sheetName, Poln_oboz_VSNRM_col_number) + 1, material_VSNRM)
            exc_WB1_save()
        End If
    End Function
    Function check_in_BOM_exist(ProjAID As Integer, PartAID As Integer, position As String) As Boolean
        'процедура проверки существования позиции в подсборке
        'сравнивается инфентарный номер и номер позиции. Если они совпадают - создается объект и вложенность
        Dim IBKey As String
        With ITS4App
            .OpenArticle(PartAID)
            IBKey = .GetArticleImbaseKey
            .CloseArticle()
        End With
        Try
            ITS4App.OpenArticleStructure(ProjAID)
            ITS4App.asFirst()
            While ITS4App.asEof = 0
                If ITS4App.asGetArtID = PartAID And ITS4App.asGetPosition = position Then
                    ITS4App.CloseArticleStructure()
                    Return True
                    Exit While
                End If
                If ITS4App.asGetArtImbaseKey = IBKey Then
                    ITS4App.CloseArticleStructure()
                    Return True
                    Exit While
                End If
                ITS4App.asNext()
            End While
            ITS4App.CloseArticleStructure()
            Return False
        Catch ex As Exception
            MsgBox("Ошибка в процедуре check_in_BOM_exist" & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Public Sub add_SP_InBOM(ProjAID As Integer)
        Try
            Dim doc_type As Integer = 1 ' тип документа - СПЕЦИФИКАЦИЯ
            Dim defolt_save_folder_path As String = "C:\IM\IMWork"
            With ITS4App
                .OpenArticle(ProjAID)
                Dim designation_main As String = .GetFieldValue_Articles("Обозначение")
                Dim naim_main As String = .GetFieldValue_Articles("Наименование")
                DocID = .CreateFileDocumentWithDocType1(defolt_save_folder_path & "\" & designation_main & ".SP", doc_type, Arch_ID, designation_main, naim_main)
                If DocID > 0 Then progress_txt_im_ListBox1("Создана СП: DocID: " & DocID & " * " & designation_main & " - " & naim_main)

                .OpenArticleStructureExpanded(ProjAID)
                .asFirst()
                Dim position_cout As Integer = .asCount
                Dim SelectedCount As Integer = 0
                Dim text2 As String = "Обработано 0 из " & position_cout
                .ShowProgressBarForm("Создание СП для Сборочный единиц", "Создание СП для Сборочный единиц в текущем дереве состава", text2, position_cout)
                'ShowProgressBarForm(position_cout, "Создание СП для Сборочный единиц", "Обработано " & position_cout)
                While .asEof = 0
                    If .asGetArtKind = 3 Then
                        Dim ArtID As Integer = .asGetArtID()
                        Dim designation As String = .asGetArtDesignation
                        Dim Naim As String = .asGetArtName
                        Dim doc_filename As String = defolt_save_folder_path & "\" & designation & ".SP" ' Имя файла в создаваемой спецификации

                        DocID = .CreateFileDocumentWithDocType1(doc_filename, doc_type, Arch_ID, designation, Naim)
                        If DocID > 0 Then progress_txt_im_ListBox1("Создана СП: DocID: " & DocID & " * " & designation & " - " & Naim)
                    End If
                    SelectedCount += SelectedCount
                    text2 = "Обработано " & SelectedCount & " из " & position_cout
                    Dim UserBreak As Integer = .SetProgressBarData_and_CheckUserBreak("", text2, SelectedCount)
                    If UserBreak > 0 Then
                        .CloseProgressBarForm()
                    End If
                    .asNext()
                End While
                .CloseArticleStructure()
                .CloseProgressBarForm()
            End With
        Catch ex As Exception
            ITS4App.CloseProgressBarForm()
            progress_txt_im_ListBox1("Сбой при создании СП! Описание ошибки: " & ex.Message)
        End Try
    End Sub
    Function add_Bom_Node(rownum As Integer, ProjAID As Integer, PartAID As Integer)
        On Error Resume Next
        Dim MU_part As String
        MU_part = WS.Cells(rownum, Ed_Ism_col_number).Value.ToString.Trim(".")
        If MU_part Is Nothing Then
            MU_part = 0
        End If
        Dim CountPC As Integer = WS.Cells(rownum, KolNaEd_col_number).value
        Dim MuID As Integer = query_s4("MU", "MU_SHORT_NAME", "MI_ID", MU_part)
        Dim TypDSEInVSNRM As String = WS.Cells(rownum, TypDSE_col_number).value
        Dim Razdel As Integer
        If TypDSEInVSNRM = "Покупное изделие" Then
            Razdel = query_s4("articles", "art_id", "SECTION_ID", PartAID)
        Else
            Razdel = query_s4("ssections", "art_kind", "SECTION_ID", TypDSEInVSNRM)
        End If

        Dim position As Integer = CInt(WS.Cells(rownum, NumPP_col_number).value.ToString.Substring(WS.Cells(rownum, NumPP_col_number).value.ToString.LastIndexOf(".") + 1))
        Dim Note As String
        Dim format_L As String = WS.Cells(rownum, format_col_number).value
        If Razdel <> 3 Then
            Note = WS.Cells(rownum, Prim_col_number).value
        Else
            format_L = "A4"
        End If
        'Dim DlinaL As String = WS.Cells(rownum, Dlina_col_number).value
        If format_L = "БЧ" And Note IsNot Nothing Then
            Note = B4_Note_Parser(Note)
        ElseIf Note.IndexOf("^2") > 0 Then 'индекс 2 уходит вверх, для случае, когда в сборочных единицах пишется площадь
            Note = Note.Replace(" = ", "=")
            Note = Note.Replace("^2", "\S2^;")
        End If
        If check_in_BOM_exist(ProjAID, PartAID, position) = False Then
            With ITS4App
                Dim pojID As Integer = .AddBOMItem(ProjAID, PartAID, CountPC, MuID, Razdel, position, Note)
                .AddArticleUserEventToLog(ProjAID, .ErrorCode(), "В состав добавлен '" & PartAID & "'")
                If format_L IsNot Nothing Or format_L <> "" Then
                    .OpenBOMItem(pojID)
                    .SetFieldValue_BOM("Формат", format_L)
                    .AddArticleUserEventToLog(pojID, .ErrorCode(), "Формат изменен с '" & .GetFieldValue_BOM("Формат") & "' на '" & format_L & "'")
                    'If DlinaL IsNot Nothing Or format_L <> "" Then .SetFieldValue_BOM("Размеры", DlinaL)
                    '.GetFieldValue_BOM("Формат")
                    .CloseBOMItem()
                End If
                'exc_rprt_edit(PartAID, ITS4App.GetDocID_ByArtID(PartAID), ITS4App.GetDocumentParams() 
            End With
        End If
    End Function
    Function B4_Note_Parser(Note As String)
        Dim thick_Char As String = "s"
        Dim pref_Char As String = "\S"
        Dim post_Char As String = "^"
        Dim perenos_Char As String = "?"
        Dim separator As String = vbLf
        Dim thick_Val, thick_Error As String

        Dim note_Split() As String
        If Note.IndexOf(thick_Char) >= 0 Then
            If Note.IndexOf(post_Char) >= 0 Then
                If Note.IndexOf(separator) >= 0 Then
                    note_Split = Note.Split(separator)
                    B4_Note_Parser = note_Split(0) '.ToString.Replace(thick_Char, "")
                Else
                    B4_Note_Parser = Note.ToString '.Replace(thick_Char, "")
                End If
            Else
                If Note.IndexOf(separator) >= 0 Then
                    note_Split = Note.Split(separator)
                    thick_Val = note_Split(0).Substring(0, note_Split(0).LastIndexOf("+"))
                    thick_Error = note_Split(0).Substring(note_Split(0).LastIndexOf("+"))
                    B4_Note_Parser = thick_Val + pref_Char + thick_Error + post_Char + perenos_Char + note_Split(1)
                Else
                    thick_Val = Note.Substring(0, Note.LastIndexOf("+"))
                    thick_Error = Note.Substring(Note.LastIndexOf("+"))
                    B4_Note_Parser = thick_Val + pref_Char + thick_Error + post_Char
                End If
            End If
        Else
            B4_Note_Parser = Note
        End If
    End Function
    Function add_Bom_Node_In_Documentacia(rownum As Integer, ProjAID As Integer)
        Dim PartAID As Integer = ITS4App.GetArtID_ByDesignation(WS.Cells(rownum, Oboz_col_number).value)

        On Error Resume Next
        Dim MU_part As String
        MU_part = WS.Cells(rownum, Ed_Ism_col_number).Value
        If MU_part Is Nothing Then
            MU_part = 0
        End If
        Dim CountPC As String = "" ' WS.Cells(rownum, KolNaEd_col_number).value
        Dim MuID As Integer = query_s4("MU", "MU_SHORT_NAME", "MI_ID", MU_part)
        Dim Razdel As Integer = 1
        Dim position As String = ""
        Dim Note As String = ""
        Dim format_L As String = WS.Cells(rownum, format_col_number).value
        If check_in_BOM_exist(ProjAID, PartAID, position) = False Then
            With ITS4App
                Dim pojID As Integer = .AddBOMItem(ProjAID, PartAID, CountPC, MuID, Razdel, position, Note)
                If format_L IsNot Nothing Or format_L <> "" Then
                    .OpenArticleStructure(pojID)
                    .SetFieldValue_BOM("Формат", format_L)
                    .AddArticleUserEventToLog(pojID, .ErrorCode(), "Формат изменен с '" & .GetFieldValue_BOM("Формат") & "' на '" & format_L & "'")
                    .CloseArticleStructure()
                End If
            End With
        End If
    End Function
    Function add_Bom_Node_In_Documentacia2(Oboz As String, ProjAID As Integer)
        Dim PartAID As Integer = ITS4App.GetArtID_ByDesignation(Oboz)

        On Error Resume Next
        Dim MU_part As String = 0
        Dim CountPC As Integer = 0
        Dim MuID As Integer = query_s4("MU", "MU_SHORT_NAME", "MI_ID", MU_part)
        Dim Razdel As Integer = 1
        Dim position As String = ""
        Dim Note As String = ""
        If check_in_BOM_exist(ProjAID, PartAID, position) = False Then
            ITS4App.AddBOMItem(ProjAID, PartAID, CountPC, MuID, Razdel, position, Note)
        End If
    End Function
    Function add_Bom_Node_In_Documentacia3(rownum As Integer, ProjAID As Integer)
        Dim PartAID As Integer = ITS4App.GetArtID_ByDesignation(WS.Cells(rownum, Oboz_col_number).value)

        On Error Resume Next
        Dim MU_part As String
        MU_part = WS_DOC.Cells(rownum, Ed_Ism_col_number).Value
        If MU_part Is Nothing Then
            MU_part = 0
        End If
        Dim CountPC As Double
        Dim MuID As Integer = query_s4("MU", "MU_SHORT_NAME", "MI_ID", MU_part)
        Dim Razdel As Integer = 1
        Dim position As String = ""
        Dim Note As String = ""
        If check_in_BOM_exist(ProjAID, PartAID, position) = False Then
            ITS4App.AddBOMItem(ProjAID, PartAID, CountPC, MuID, Razdel, position, Note)
        End If
    End Function
    Function query_s4(table_name As String, column_name As String, find_column_name As String, find_text As String)
        ITS4App.OpenQuery("select * from " & table_name & " where " & column_name & " = '" & find_text & "'")
        query_s4 = ITS4App.QueryFieldByName(find_column_name)
        ITS4App.CloseQuery()
    End Function
    Function query_s4_2_answer(table_name As String, column_name1 As String, column_name2 As String, find_column_name As String, find_text1 As String, find_text2 As String)
        ITS4App.OpenQuery("select * from " & table_name & " where " & column_name1 & " = '" & find_text1 & "' And " & column_name2 & " = '" & find_text2 & "'")
        query_s4_2_answer = ITS4App.QueryFieldByName(find_column_name)
        ITS4App.CloseQuery()
    End Function

    Function query_s4_readonly(table_name As String, column_name As String, find_column_name As String, find_text As String)
        ITS4App.OpenQueryEx("select * from " & table_name & " where " & column_name & " = '" & find_text & "'")
        query_s4_readonly = ITS4App.QueryFieldByName(find_column_name)
        ITS4App.CloseQuery()
    End Function
    Private Sub СправкаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СправкаToolStripMenuItem.Click
        INFOrmation()
    End Sub
    Sub INFOrmation()
        Process.Start("http://www.evernote.com/l/AjJptzkrFONOHIrynbheNwbKU-OJ40zG0Nw/")
    End Sub

    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        App_info.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        'find_file_by_ObozNainFformat("АБВГ.123456.00.075", "Проушина", "tif")
        'Dim filename As String = "ТНВП.011752.01.000СБ Аппарат направляющий (изм.1488).tif"
        'Dim izm As String = filename.Substring(filename.IndexOf("изм.") + "изм.".Length)
        'izm = izm.Substring(0, izm.LastIndexOf(")"))
        'ITS4App.OpenDocument(4952)
        'ITS4App.GetFieldValue("Изменение")
        'ITS4App.SetDocVersionCode("4")
        'ITS4App.GetFieldValue("Изменение") 
        'Dim first_Name As String = query_s4("USERS", "fullname", "LOGINNAME", ITS4App.GetUserFullName_ByUserID(ITS4App.GetUserID))
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        'GroupBox3.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        'GroupBox3.Enabled = False
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Control Or e.Alt Or e.Shift Or e.KeyCode = 13 Or e.KeyCode = 91 Or e.KeyCode = 92 Then
            e.SuppressKeyPress = True
        End If

        If e.Control And e.KeyCode.ToString = "O" Then
            Me.KeyPreview = False
            new_doc.ShowDialog()
        ElseIf e.Control And e.KeyCode.ToString = "E" Then
            select_VSNRM()
        ElseIf e.Control And e.KeyCode.ToString = "F3" Then
            Me.KeyPreview = False
            Options.ShowDialog()
        ElseIf e.Control And e.Shift And e.KeyCode.ToString = "K" Then
            LicLD()
        ElseIf e.Control And e.Shift And e.KeyCode.ToString = "L" Then
            licCFG()
        End If

        Select Case e.KeyCode
            Case Keys.F1
                Me.KeyPreview = False
                Call INFOrmation()
            Case Keys.F2
                Me.KeyPreview = False
                App_info.ShowDialog()
        End Select
        Me.KeyPreview = True
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        MainFormSize()
        exc_close()
    End Sub

    Private Sub НастройкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НастройкиToolStripMenuItem.Click
        Options.ShowDialog()
    End Sub

    'Private Sub TreeView1_Click(sender As Object, e As EventArgs) Handles TreeView1.Click
    '    'Dim node_position As String = TreeView1.SelectedNode.Text

    'End Sub

    Private Sub СоздатьДеревоToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СоздатьДеревоToolStripMenuItem.Click
        new_doc.ShowDialog()
    End Sub

    Private Sub СравнитьВСНРМСИзделиемВS4ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Compare_BOM.ShowDialog()
    End Sub

    Public Function SelectArticl()
        Dim selArtID As Integer
        ITS4App.StartSelectArticles()
        ITS4App.SelectArticles()
        If ITS4App.SelectedArticlesCount = 0 Then
            Exit Function
        End If
        selArtID = ITS4App.GetSelectedArticleID(0)
        ITS4App.EndSelectArticles()
        Return selArtID
    End Function

    Private Sub ОчиститьИсториюToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОчиститьИсториюToolStripMenuItem.Click
        clear_command_listbox()
    End Sub
    Sub clear_command_listbox()
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        Dim s1 As String = "ВСНРМ ТНВП.010749.00.000СБ - Вентилятор.xlsx"
        Dim f As Integer = s1.IndexOf("ВСНРМ")
        If s1.IndexOf("ВСНРМ") > 0 Then
            MsgBox(f)
        Else
            MsgBox(f)
        End If
        'Dim obozn As String = s1.Substring(0, s1.LastIndexOf(" - "))
        'obozn = obozn.Substring(obozn.LastIndexOf(" ")).Trim(" ")
        'Dim selArtID As Integer = ITS4App.GetArtID_ByDesignation(obozn)
        'ITS4App.OpenArticle(selArtID)
        'Dim naim As String = ITS4App.GetFieldValue_Articles("Наименование")
        'Dim mass As String = ITS4App.GetFieldValue_Articles("Масса")
        'Dim ob As String = ITS4App.GetFieldValue_Articles("Обозначение")
        'Dim material As String = ITS4App.GetFieldValue_Articles("Материал")
        'ITS4App.CloseArticle()
        'ITS4App.OpenVArticle(10268)
        'ITS4App.VArticleCard()
        'ITS4App.CloseVArticle()
    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As Object, e As ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening
        If TreeView1.Nodes.Count = 0 Then
            ToolStripMenuItem1.Enabled = False
            ОткрытьКарточкуДокументаToolStripMenuItem.Enabled = False
            'ОткрытьВСНРМToolStripMenuItem.Enabled = False
        Else
            ToolStripMenuItem1.Enabled = True
            ОткрытьКарточкуДокументаToolStripMenuItem.Enabled = True
            'ОткрытьВСНРМToolStripMenuItem.Enabled = True
        End If
        Dim ExcelFullFileName As String = new_doc.ExcelFullFileName
        If ExcelFullFileName <> Nothing Or ExcelFullFileName <> "" Then
            ОткрытьВСНРМToolStripMenuItem1.Enabled = True
        Else
            ОткрытьВСНРМToolStripMenuItem1.Enabled = False
        End If
        If Carto4kaS4SPEnable = False Then
            ToolStripMenuItem1.Enabled = False
        Else
            ToolStripMenuItem1.Enabled = True
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        pre_Open_Art_Properety()
    End Sub
    Sub pre_Open_Art_Properety()
        Try
            Dim selArtID As Integer = getArtIDBySelectedNodeInTreeView()
            If selArtID > 0 Then
                Art_property.ArtID = selArtID
                Art_property.ShowDialog()
            Else
                MsgBox("У данного компонента ВСНРМ пока нет Объекта в S4")
            End If
        Catch ex As Exception
            MsgBox("Ошибка!" & vbNewLine & ex.Message)
        End Try
    End Sub
    Function getArtIDBySelectedNodeInTreeView()
        Try
            If TreeView1.Nodes.Count = 0 Then Exit Function
            Dim s1 As String = TreeView1.SelectedNode.Text
            Dim obozn As String = s1.Substring(0, s1.LastIndexOf(" - "))
            obozn = obozn.Substring(obozn.LastIndexOf(" ")).Trim(" ")
            Dim selArtID As Integer = ITS4App.GetArtID_ByDesignation(obozn)
            If selArtID < 0 Then
                obozn = obozn.Replace("СБ", "")
                selArtID = ITS4App.GetArtID_ByDesignation(obozn)
            End If
            'selArtID = TreeView1.SelectedNode.Tag
            Return selArtID
        Catch ex As Exception
            MsgBox("Ошибка!" & vbNewLine & ex.Message)
        End Try
    End Function
    Private Sub TreeView1_KeyDown(sender As Object, e As KeyEventArgs) Handles TreeView1.KeyDown
        Try
            If e.KeyCode = Keys.F4 Then
                If TreeView1.Nodes.Count > 0 Then
                    pre_Open_Art_Properety()
                End If
            End If
        Catch ex As Exception
            MsgBox("Ошибка!" & vbNewLine & ex.Message)
        End Try
    End Sub

    Private Sub TreeView1_DragDrop(sender As Object, e As DragEventArgs) Handles TreeView1.DragDrop
        Dim file, fileformat As String
        For Each file In e.Data.GetData(DataFormats.FileDrop)
            fileformat = file.Substring(file.LastIndexOf(".") + 1)
            If UCase(fileformat) = "XLSX" And file.IndexOf("ВСНРМ") <> 0 Then
                Drag_Drop_Form.VSNRM_path = file
                Drag_Drop_Form.ShowDialog()
                Exit For
            End If
        Next
    End Sub

    Private Sub TreeView1_DragEnter(sender As Object, e As DragEventArgs) Handles TreeView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub ПроверитьДеревоСоставаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПроверитьДеревоСоставаToolStripMenuItem.Click
        Checking_SOSTAV_Form.ShowDialog()
    End Sub
    Public Sub check_article_SOSTAV()
        Checking_SOSTAV_Form.ListBox1.Items.Clear()
        Dim cmd As String = "Начата проверка структуры"
        Dim str_DocID As String = "DocID: "
        Dim str_ArdID As String = "ArtID: "
        Dim separator As String = " - "

        With ITS4App
            Dim progId As Integer
            .StartSelectArticles()
            .SelectArticles()
            If .SelectedArticlesCount = 0 Then
                Exit Sub
            End If
            progId = .GetSelectedArticleID(0)
            .EndSelectArticles()
            cmd = str_ArdID & progId & separator & "Выбран объект"
            report_in_LB_checkingSOSTAV(cmd)
            If progId <= 0 Then Exit Sub
            check_article_SOSTAV_recursuive(progId)
            '.OpenArticleStructure2(ArtID, -1, 2)
            '.OpenArticleStructure2(progId, -1, 1) '.OpenArticleStructureExpanded(progId)
            '.asFirst()
            'While .asEof = 0
            '    Dim razdel As Integer = .asGetArtKind
            '    '     1 'документация
            '    '     2 'комплексы
            '    '     3 'сборочная единица
            '    '     4 'деталь
            '    '     5 'стандартное изделие
            '    '     6 'прочие изделия
            '    '     7 'материалы
            '    '     8 'комплекты
            '    Dim tmp_artID As Integer = .asGetArtID
            '    Dim naim_BOM As String = .asGetArtName
            '    If naim_BOM = "" Or naim_BOM Is Nothing Then
            '        cmd = "не заполнено поле НАИМЕНОВНИЕ"
            '        report_in_LB_checkingSOSTAV(cmd)
            '    End If
            '    Select Case razdel
            '        Case 1, 2, 3, 4 'деталь
            '            Dim PrjLinkID As String = .asGetPrjLinkID
            '            Dim tmp_Doc_ID As Integer = .GetDocID_ByArtID(tmp_artID)

            '            'проверка заполненности ФОРМАТА
            '            .OpenBOMItem(PrjLinkID)
            '            Dim format_BOM As String = .GetFieldValue_BOM("Формат")
            '            If format_BOM = "" Or format_BOM Is Nothing Then
            '                cmd = str_ArdID & tmp_artID & separator & "не заполнено поле Формат"
            '                report_in_LB_checkingSOSTAV(cmd)
            '            End If
            '            .CloseBOMItem()
            '            If format_BOM <> "БЧ" Then
            '                'проверка на наличие Документа у объекта (не забыть про случай с БЧ)
            '                If tmp_Doc_ID < 1 Then
            '                    cmd = str_ArdID & tmp_artID & separator & "нет документа!"
            '                    report_in_LB_checkingSOSTAV(cmd)
            '                End If
            '            End If
            '            'проверка заполненности МАТЕРИАЛА
            '            If razdel > 3 Then
            '                Dim tmp_material As String = .GetFieldValue_Articles("Материал")
            '                If tmp_material Is Nothing Or tmp_material = "" Then
            '                    cmd = str_ArdID & tmp_artID & separator & "не заполнен материал"
            '                    report_in_LB_checkingSOSTAV(cmd)
            '                End If
            '            End If
            '            'проверка заполненности МАССЫ
            '            Dim tmp_mass As String = .GetFieldValue_Articles("Масса")
            '            If tmp_mass Is Nothing Or tmp_mass = "" Then
            '                cmd = str_ArdID & tmp_artID & separator & "не заполнена масса"
            '                report_in_LB_checkingSOSTAV(cmd)
            '            End If
            '        Case Else

            '    End Select
            '    'проверка заполненности КОЛ-ВО
            '    Dim count_BOM As Double = .asGetArtCount
            '    If count_BOM = 0 Then
            '        cmd = str_ArdID & tmp_artID & separator & "не указано КОЛ-ВО"
            '        report_in_LB_checkingSOSTAV(cmd)
            '    End If
            '    .asNext()
            'End While
            '.CloseArticleStructure()
            cmd = "Проверка завершена!"
            report_in_LB_checkingSOSTAV(cmd)
        End With
    End Sub
    Function check_article_SOSTAV_recursuive(progId As Integer)
        Dim cmd As String
        Dim str_DocID As String = "DocID: "
        Dim str_ArdID As String = "ArtID: "
        Dim separator As String = " - "
        Dim razdel, tmp_artID As Integer
        With ITS4App
            .OpenArticleStructure2(progId, -1, 1) '.OpenArticleStructureExpanded(progId)
            .asFirst()
            While .asEof = 0
                Dim checkasEof As Integer = .asEof
                razdel = .asGetArtKind
                '     1 'документация
                '     2 'комплексы
                '     3 'сборочная единица
                '     4 'деталь
                '     5 'стандартное изделие
                '     6 'прочие изделия
                '     7 'материалы
                '     8 'комплекты
                tmp_artID = .asGetArtID

                Dim naim_BOM As String = .asGetArtName
                If naim_BOM = "" Or naim_BOM Is Nothing Then
                    cmd = str_ArdID & tmp_artID & separator & "не заполнено поле Наименование"
                    report_in_LB_checkingSOSTAV(cmd)
                End If
                Select Case razdel
                    Case 1, 2, 3, 4 'деталь
                        Dim PrjLinkID As String = .asGetPrjLinkID
                        Dim tmp_Doc_ID As Integer = .GetDocID_ByArtID(tmp_artID)

                        'проверка заполненности ФОРМАТА
                        .OpenBOMItem(PrjLinkID)
                        Dim format_BOM As String = .GetFieldValue_BOM("Формат")
                        If format_BOM = "" Or format_BOM Is Nothing Then
                            cmd = str_ArdID & tmp_artID & separator & "не заполнено поле Формат"
                            report_in_LB_checkingSOSTAV(cmd)
                        End If
                        .CloseBOMItem()
                        .OpenArticle(tmp_artID)
                        If format_BOM <> "БЧ" Then
                            'проверка на наличие Документа у объекта (не забыть про случай с БЧ)
                            If tmp_Doc_ID < 1 Then
                                cmd = str_ArdID & tmp_artID & separator & "нет документа!"
                                report_in_LB_checkingSOSTAV(cmd)
                            End If
                        End If
                        'проверка заполненности МАТЕРИАЛА
                        If razdel > 3 Then
                            Dim tmp_material As String = .GetFieldValue_Articles("Материал")
                            If tmp_material Is Nothing Or tmp_material = "" Then
                                cmd = str_ArdID & tmp_artID & separator & "не заполнено поле Материал"
                                report_in_LB_checkingSOSTAV(cmd)
                            End If
                        End If
                        'проверка заполненности МАССЫ
                        If razdel <> 3 Then
                            Dim tmp_mass As String = .GetFieldValue_Articles("Масса")
                            If tmp_mass Is Nothing Or tmp_mass = "" Then
                                cmd = str_ArdID & tmp_artID & separator & "не заполнено поле Масса"
                                report_in_LB_checkingSOSTAV(cmd)
                            End If
                        End If
                        .CloseArticle()
                    Case Else

                End Select
                'проверка заполненности КОЛ-ВО
                If razdel > 1 Then
                    Dim count_BOM As Double = .asGetArtCount
                    If count_BOM = 0 Then
                        cmd = str_ArdID & tmp_artID & separator & "не заполнено поле Кол-во"
                        report_in_LB_checkingSOSTAV(cmd)
                    End If
                End If

                If razdel = 3 Then
                    check_article_SOSTAV_recursuive(tmp_artID)
                    .OpenArticleStructure2(progId, -1, 1)
                    .asFirst()
                    While .asEof = 0
                        If .asGetArtID = tmp_artID Then
                            Exit While
                        End If
                        .asNext()
                    End While
                End If

                .asNext()
            End While
            .CloseArticleStructure()
        End With
    End Function
    Public Sub report_in_LB_checkingSOSTAV(comand As String)
        Checking_SOSTAV_Form.ListBox1.Items.Add(comand)
    End Sub

    Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestToolStripMenuItem.Click
        whatsNew()
    End Sub
    Public Sub whatsNew()
        Try
            Process.Start(System.Windows.Forms.Application.StartupPath & "\what's new.txt")
            'Process.Start("http://www.evernote.com/l/AjJ5PRJAb7ZB3raLw5_efcpQ5Zd5Wtz_MWc/")
        Catch ex As Exception
            MsgBox("Не могу открыть файл: " & System.Windows.Forms.Application.StartupPath & "\what's new.txt" & vbNewLine & "Проверьте существование файла в указанной директории")
        End Try
    End Sub

    Private Sub ShowArtDocumentationToolStripMenuItem_Click(sender As Object, e As EventArgs)
        With ITS4App
            .OpenArticle(11788)
            .EditParameters2_Article()
        End With
    End Sub

    Private Sub ПросмотрКарточкиДокументаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПросмотрКарточкиДокументаToolStripMenuItem.Click
        Try
            With ITS4App
                Dim thisArtID As Integer = getArtIDBySelectedNodeInTreeView()
                If thisArtID > 0 Then
                    .OpenArticle(thisArtID)
                    .EditParameters_Article()
                    .CloseArticle()
                Else
                    MsgBox("У данного компонента ВСНРМ пока нет Объекта в S4")
                End If
            End With
        Catch ex As Exception
            MsgBox("Ошибка!" & vbNewLine & ex.Message)
        End Try
    End Sub

    Private Sub РедактированиеКарточкиДокументаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles РедактированиеКарточкиДокументаToolStripMenuItem.Click
        Try
            With ITS4App
                Dim thisArtID As Integer = getArtIDBySelectedNodeInTreeView()
                If thisArtID > 0 Then
                    Dim thisDocID As Integer = .GetDocID_ByArtID(thisArtID)
                    If thisDocID > 0 Then
                        .OpenDocument(thisDocID)
                        .CheckOut()
                        .EditParameters()
                    Else
                        .OpenArticle(thisArtID)
                        .EditParameters_Article()
                        .CloseArticle()
                    End If
                Else
                    MsgBox("У данного компонента ВСНРМ пока нет Объекта в S4")
                End If
            End With
        Catch ex As Exception
            MsgBox("Ошибка!" & vbNewLine & ex.Message)
        End Try
    End Sub

    Private Sub FormToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ITS4App.ShowCustomFormDlg_Articles(11787, 10, "Hello, World", 100, 100)
    End Sub

    Private Sub FindBOMItemToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim prj_link_ID As Integer = ITS4App.FindBOMItem(11788, 11791)
    End Sub

    Private Sub ОткрытьВСНРМToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьВСНРМToolStripMenuItem.Click
        OpenExcelFileVSNRM()
    End Sub

    Private Sub ОткрытьВСНРМToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОткрытьВСНРМToolStripMenuItem1.Click
        OpenExcelFileVSNRM()
    End Sub
    Public Sub OpenExcelFileVSNRM()
        Try
            Open_Second_EX_Doc(True, new_doc.ExcelFullFileName)
        Catch ex As Exception
            MessageBox.Show("Не удается открыть файл. Код ошибки: " & ex.Message)
        End Try
    End Sub

    Private Sub ВыбратьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыбратьToolStripMenuItem.Click
        Dim ExcelFullFileName As String = new_doc.ExcelFullFileName
        If ExcelFullFileName <> Nothing Or ExcelFullFileName <> "" Then
            ОткрытьВСНРМToolStripMenuItem.Enabled = True
        Else
            ОткрытьВСНРМToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem.Click
        CompareMaterialInVSNRMWithMaterialTable()
    End Sub
    Private Sub ЖурналПроектовToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЖурналПроектовToolStripMenuItem.Click
        Try
            Dim Copy_directory As String
            If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
                Copy_directory = FolderBrowserDialog1.SelectedPath
            Else
                Exit Sub
            End If

            With ITS4App
                Dim progId As Integer
                .StartSelectArticles()
                .SelectArticles()
                If .SelectedArticlesCount = 0 Then
                    Exit Sub
                End If
                progId = .GetSelectedArticleID(0)
                .EndSelectArticles()
                'cmd = str_ArdID & progId & separator & "Выбран объект"
                'report_in_LB_checkingSOSTAV(cmd)
                If progId <= 0 Then Exit Sub
                CopyToDirProc(progId, Copy_directory)
                '.GetSelectedArticleID()
                .OpenArticleStructureExpanded2(progId, -1, 2)
                .asFirst()
                While .asEof = 0
                    Dim tmp_artID As Integer = .asGetArtID
                    CopyToDirProc(tmp_artID, Copy_directory)
                    .asNext()
                End While
            End With
            MsgBox("Выгрузка КД по выбранному объекту " & "Завершена")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub CopyToDirProc(tmp_artID As Integer, Copy_directory As String)
        Try
            With ITS4App
                Dim tmp_Doc_ID As Integer = .GetDocID_ByArtID(tmp_artID)
                Dim FFN, Oboz As String
                Dim DopFileArray(), DopFileStr As String
                If tmp_Doc_ID > 0 Then
                    .OpenDocument(tmp_Doc_ID)
                    FFN = .GetFieldValue("Имя файла") ' .GetDocFilename(tmp_Doc_ID)
                    Oboz = .GetFieldValue("Обозначение")
                    Dim Fformat_Main As String = Path.GetExtension(FFN)                                                     'расширение
                    Dim FileNameWExt As String = Path.GetFileNameWithoutExtension(FFN)                                      'имя файла без расширения
                    Dim newFFN_Main As String = Copy_directory & "\" & Oboz & " {" & FileNameWExt & "}" & Fformat_Main           'новое(переименованное) имя файла
                    DopFileStr = .GetAdvanFilesList()
                    .CopyToDir(Copy_directory)
                    .CloseDocument()
                    If Options.CopyDir_Method Then
                        If DopFileStr IsNot Nothing Then
                            DopFileArray = DopFileStr.Split(vbCrLf)
                            For j As Integer = 0 To UBound(DopFileArray) - 1
                                Dim tmpDFN As String = DopFileArray(j).Trim(vbLf)
                                Dim newFN As String = Copy_directory & "\" & Oboz & " {" & Path.GetFileNameWithoutExtension(Copy_directory & "\" & tmpDFN) & "}" & Path.GetExtension(tmpDFN)
                                'процедура которая переименоваывает файл на новое значение
                                If File.Exists(newFN) Then
                                    File.Delete(newFN)
                                End If
                                System.IO.File.Move(Copy_directory & "\" & tmpDFN, newFN)
                            Next
                        End If
                        If File.Exists(newFFN_Main) Then
                            File.Delete(newFFN_Main)
                        End If
                        System.IO.File.Move(Copy_directory & "\" & FFN, newFFN_Main)
                        End If
                    End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function SelectArticl_procedure() As Integer
        Dim select_ArtID As Integer
        ITS4App.StartSelectArticles()
        ITS4App.SelectArticles()
        If ITS4App.SelectedArticlesCount = 0 Then
            Exit Function
        End If
        select_ArtID = ITS4App.GetSelectedArticleID(0)
        ITS4App.EndSelectArticles()
        Return select_ArtID
    End Function

    Private Function recvAdd1488(node As TreeNode, rownum As Integer, isFirst As Boolean)
        'прогресс бар заполнение 
        'SetProgressBarData_and_CheckUserBreak(rownum - 5, LastInd - 4)

        Dim numpp As String = WS.Cells(rownum, NumPP_col_number).Value()

        Dim ptrn1
        If isFirst Then
            ptrn1 = New Regex("^.+")
        Else
            ptrn1 = New Regex("^" + node.Text.Substring(0, node.Text.IndexOf(" ")) + ".+")
        End If
        While (numpp IsNot Nothing)
            If Not ptrn1.IsMatch(numpp) Then
                Exit While
            End If
            'запустить создание объекта в с4 (Тип объекта брать из ТИПА ДСЕ). Вернет ее PartAID (понадобится следующем уровне)
            'Dim chekc_art_exist = ITS4App.GetArtID_ByDesignation(WS.Cells(rownum, Oboz_col_number).value)

            Dim newNode As New TreeNode(numpp & " " & WS.Cells(rownum, Oboz_col_number).Value() & " - " & WS.Cells(rownum, Naim_col_number).Value())
            Dim part_type As String = WS.Cells(rownum, TypDSE_col_number).value
            If (part_type = "Деталь" Or part_type = "Сборочная единица") Then
                newNode.Tag = add_article(rownum)
                DocID = add_document(rownum)
                Call add_Bom_Node(rownum, node.Tag, newNode.Tag)
                'запустить создание ветки в с4 AddBOMItem
                'PartAID записать в тегах текущей ветки 

                'записать в раздел документация сборочный чертеж
                If part_type = "Сборочная единица" Then
                    Call add_Bom_Node_In_Documentacia(rownum, newNode.Tag)
                End If
            End If
            node.Nodes.Add(newNode)

            rownum += 1
            'прогресс бар заполнение 
            'SetProgressBarData_and_CheckUserBreak(rownum - 4, LastInd - 4)

            Dim nextnumpp As String = WS.Cells(rownum, NumPP_col_number).Value()
            If nextnumpp Is Nothing Then
                Exit While
            End If
            Dim ptrn2 As New Regex("^" + numpp + "\.[0-9]+$")
            If ptrn2.IsMatch(nextnumpp) Then
                rownum = recvAdd1488(newNode, rownum, False)
                numpp = WS.Cells(rownum, NumPP_col_number).Value()
            Else
                numpp = nextnumpp
            End If
        End While
        Return rownum
    End Function

    Private Sub CreateBoltS4ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'ITS4App.OpenArticle(11650)
        'ITS4App.GetFieldImbaseKey_Articles("Наименование")
        'add_materialInBOM_S4(1, 7)
    End Sub

    Private Sub ЫфвыцуаToolStripMenuItem_Click(sender As Object, e As EventArgs)
        B4_Note_Parser("s8+1,5
Rz23")
    End Sub

    Public Sub Compare_BOM_proc(ArtID As Integer)
        With ITS4App
            .OpenArticleStructureExpanded(ArtID)
            .asFirst()
            While .asEof = 0
                If .asGetArtKind = 3 Then
                    Dim Part_ArtID As Integer = .asGetArtID()
                    Dim S_position As String = .asGetPosition
                    Dim S_designation As String = .asGetArtDesignation
                    Dim S_Naim As String = .asGetArtName
                    Dim S_CountInSostav As String = .asGetArtCount
                    .OpenArticle(Part_ArtID)
                    Dim S_Mass As String = .GetFieldValue_Articles("Масса")
                    Dim S_Material As String = .GetFieldValue_Articles("Материал")
                    Dim S_Tolshina_Diametr As String = .GetFieldValue_Articles("Толщина/Диаметр")
                    Dim S_Dlina As String = .GetFieldValue_Articles("Примечание")

                    'Dim
                End If
                .asNext()
            End While
            .CloseArticleStructure()
            .CloseProgressBarForm()
        End With
    End Sub

    Private Sub DownloadToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim Copy_directory As String
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            Copy_directory = FolderBrowserDialog1.SelectedPath
        Else
            Exit Sub
        End If

        With ITS4App
            .OpenDocument(5755)
            .CopyToDir(Copy_directory)
            .CloseDocument()
        End With
    End Sub

    Public Sub SetOTDRegnum_SelectDocs(OTD_cmd As Integer, OTD_reg As String, OTD_Annul As String)
        Try
            With ITS4App
                .StartSelectDocs()
                .SelectDocsEx2(-5, 0)
                Dim docsCount As Integer = .SelectedDocsCount
                For q As Integer = 0 To docsCount - 1
                    'MsgBox(.GetSelectedDocID(q))
                    SetOTDREGNUM(.GetSelectedDocID(q), OTD_cmd, OTD_reg, OTD_Annul)
                Next
                .EndSelectDocs()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub SetOTDREGNUM(VersDocID As Integer, OTD_cmd As Integer, OTD_reg As String, OTD_Annul As String)
        'процедура регистрирует документ в ОТД и задает ему значение инвентарного номера ОТД
        Try
            Dim progressStatus As String = ""
            Dim str_DocID As String = "DocID: "
            Dim str_ArdID As String = "ArtID: "
            Dim separator As String = " - "
            With ITS4App
                .OpenDocument(VersDocID) '4040
                .SetFieldValue_DocVersion("OTD_STATUS", OTD_cmd)
                .AddDocumentUserEventToLog(VersDocID, -1, .ErrorCode(), "OTD_STATUS изменен на '" & OTD_cmd & "'")
                .SetFieldValue_DocVersion("OTD_reg", OTD_reg)
                .AddDocumentUserEventToLog(VersDocID, -1, .ErrorCode(), "OTD_reg изменен на '" & OTD_reg & "'")
                .SetFieldValue_DocVersion("OTD_Annul", OTD_Annul)
                .AddDocumentUserEventToLog(VersDocID, -1, .ErrorCode(), "OTD_Annul изменен на '" & OTD_Annul & "'")
                Select Case OTD_cmd
                    Case 1 'регистрировать в ОТД
                        'If .GetFieldValue_DocVersion("OTD_STATUS") > 0 Then Exit Sub
                        If .GetFieldValue_DocVersion("OTDREGNUM") = "" Then
                            Dim maxOTDregnum As String = MAXotdregnumSQLQuere()
                            .SetFieldValue_DocVersion("OTDREGNUM", maxOTDregnum)
                            .AddDocumentUserEventToLog(VersDocID, -1, .ErrorCode(), "зарегестрирован в ОТД. OTDRegnum = " & maxOTDregnum)
                            progressStatus = str_DocID & VersDocID & separator & "(" & .GetFieldValue("Обозначение") & ") зарегестрирован в ОТД. OTDRegnum = " & maxOTDregnum
                        Else
                            progressStatus = str_DocID & VersDocID & separator & "(" & .GetFieldValue("Обозначение") & ") уже зарегестрирован в ОТД! " & "Инвентарный номер ОТД: " & .GetFieldValue_DocVersion("OTDREGNUM")
                            MsgBox(progressStatus)
                        End If
                    Case 2 'аннулировать
                        progressStatus = str_DocID & VersDocID & separator & " (" & .GetFieldValue("Обозначение") & ") аннулирован в ОТД."
                        .AddDocumentUserEventToLog(VersDocID, -1, .ErrorCode(), "Документ аннулирован в ОТД")
                    Case 0 'не зарегестрирован в ОТД 
                        .SetFieldValue_DocVersion("OTDREGNUM", "")
                        progressStatus = str_DocID & VersDocID & separator & " (" & .GetFieldValue("Обозначение") & ") снят с учета (не зарегестрирован) в ОТД."
                        .AddDocumentUserEventToLog(VersDocID, -1, .ErrorCode(), "Документ снят с учета (не зарегестрирован) в ОТД")
                End Select
                .CloseDocument()
                OTDEditor.ProcessAdd(progressStatus)
            End With
        Catch ex As Exception
            MsgBox("Ошибка редактировании параметров ОТД!" & vbNewLine & ex.Message)
        End Try
    End Sub
    Public Function MAXotdregnumSQLQuere() As String
        Try
            With ITS4App
                .OpenQuery("SELECT MAX(otdregnum) AS otdregnum FROM doclist")
                Dim maxOTDregnum As String = .QueryFieldByName("otdregnum")
                .CloseQuery()
                If maxOTDregnum Is Nothing Then
                    maxOTDregnum = InputBox("Введите начальное значение Инв.номера ОТД")
                End If
                Dim tmpOTDregnum As Integer = maxOTDregnum + 1
                Dim tmpOTDregnumStr As String = tmpOTDregnum
                For w As Integer = 1 To maxOTDregnum.Length
                    If tmpOTDregnumStr.Length = maxOTDregnum.Length Then
                        Exit For
                    Else
                        tmpOTDregnumStr = "0" & tmpOTDregnumStr
                    End If
                Next
                Return tmpOTDregnumStr
            End With
        Catch ex As Exception
            MsgBox("Ошибка при отправке SQL-запроса!" & vbNewLine & ITS4App.ErrorMessage)
        End Try
    End Function

    Private Sub ЖурналОТДToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЖурналОТДToolStripMenuItem.Click
        'OTDForm.ShowDialog()
        OTDEditor.ShowDialog()
    End Sub

    Private Sub ImbaseToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'Dim IB As ImBase.IImbaseApplication = CreateObject("ImBase.ImbaseApplication")
        'Dim status As ImBase.ImbaseStatus = IB.Status
        'If IB Is Nothing Then
        '    IB = CreateObject("ImBase.ImbaseApplication")
        '    While True
        '        status = IB.Status
        '        System.Windows.Forms.Application.DoEvents()
        '        If status = "IST_READY" Then Exit While
        '    End While
        'End If
        'Dim fname As String = IB.Catalogs.Item(1).Name
        'IB.Blobs.sh
    End Sub

    Private Sub ПроверкаToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim s1 As String = "DocID: 4040 - Пример ВСНРМ"
        Dim result As String = s1.Substring(0, s1.IndexOf(" - "))
        result = result.Substring(s1.IndexOf(": ") + 2).Trim

    End Sub

    Private Sub ConverterToolStripMenuItem_Click(sender As Object, e As EventArgs)
        With ITS4App
            '.ShowCustomFormDlg_Articles(7711, InputBox("ID form"), "тестовая форма", 500, 250)
            .OpenDocument(7711)
            Dim param As String = .GetArtBasicFields()
        End With
    End Sub
    Public Sub SelItems()
        With ITS4App
            Dim SelectedCount, AID, DID As Integer
            Dim SelItems As S4.TS4SelectedItems = .GetSelectedItems
            SelectedCount = SelItems.SelectedCount
            While SelItems.NextSelected <> 0
                DID = SelItems.ActiveDocID
                If DID > 0 Then
                    AID = .GetArtID_ByDocID(DID)
                Else
                    AID = SelItems.ActiveArtID
                End If
                If AID > 0 Then
                    .OpenArticle(AID)
                    If .GetArticleKind = 3 Then
                        change_mass(AID)
                    Else
                        MsgBox("Тип Объекта должен быть Сборочная единица")
                    End If
                End If
                SelItems.InvertCurrent()
            End While
        End With

    End Sub
    Sub change_mass(AID As Integer)
        With ITS4App
            Dim find_AID As Integer
            Dim old_mass, AID_designatio, find_designatio, find_mass As String

            .OpenArticle(AID)
            old_mass = .GetFieldValue_Articles("Масса")
            If old_mass <> "" Or old_mass IsNot Nothing Then Exit Sub
            AID_designatio = .GetFieldValue_Articles("Обозначение")
            find_designatio = AID_designatio & " СБ"
            find_AID = .GetArtID_ByDesignation(find_designatio)
            If find_AID <= 0 Then
                find_designatio = AID_designatio & "СБ"
                find_AID = .GetArtID_ByDesignation(find_designatio)
            End If
            .OpenArticle(find_AID)
            find_mass = .GetFieldValue_Articles("Масса")
            .CloseArticle()
            If find_mass <> "" Or find_mass IsNot Nothing Then
                .OpenArticle(AID)
                .SetFieldValue_Articles("Масса", find_mass)

                progress_txt_im_ListBox1("У объекта " & AID & " (" & .GetFieldValue_Articles("Обозначение") & ") была изменена масса. Новая М = " & find_mass)
                .AddArticleUserEventToLog(AID, 1, "Изменена масса с '" & .GetFieldValue_Articles("Масса") & "' на '" & find_mass & "'")
                .CloseArticle()
            End If
        End With
    End Sub
    Private Sub УтвАрхивРекдакToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SelItems()
    End Sub

    Private Sub ПравитьОбъектToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПравитьОбъектToolStripMenuItem.Click
        Try
            Art_property.ArtID = ITS4App.GetArtID_ByDocID(ITS4App.GetSelectedItems.ActiveDocID)
            Art_property.ShowDialog()
        Catch ex As Exception
            MsgBox("Ошибка!" & vbNewLine & ex.Message)
        End Try
    End Sub

    Private Sub НомерОТДToolStripMenuItem_Click(sender As Object, e As EventArgs)
        With ITS4App
            Dim ArtId, DocID, VerDocID As Integer
            Dim OTDREGNUM As String
            DocID = .GetSelectedItems.ActiveDocID
            ArtId = .GetArtID_ByDocID(DocID)
            .OpenDocument(DocID)
            VerDocID = .GetFieldValue("Version_ID")
            .OpenDocVersion(DocID, VerDocID)
            OTDREGNUM = .GetFieldValue_DocVersion("OTDREGNUM")
            MsgBox(OTDREGNUM)
        End With
    End Sub

    Private Sub ВСНРМ2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВСНРМ2ToolStripMenuItem.Click
        VSNRM2.ShowDialog()
    End Sub

    Private Sub ТетсToolStripMenuItem_Click(sender As Object, e As EventArgs)
        MsgBox("ghbdtn")
    End Sub

    Private Sub ВерсияToolStripMenuItem_Click(sender As Object, e As EventArgs)
        With ITS4App
            .OpenArticle(14284)
            Dim VERS As String = .GetFieldValue_Articles("Art_Ver_ID")
            .CloseArticle()
        End With
    End Sub

    Private Sub TxtToolStripMenuItem_Click(sender As Object, e As EventArgs)
        C()
    End Sub

    Private Sub EventToolStripMenuItem_Click(sender As Object, e As EventArgs)
        With ITS4App
            Dim EvID As Integer = .AddArticleUserEventToLog(14905, 1, "Привет")
        End With
    End Sub

    Private Sub ImbaseToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        With ITS4App
            .OpenArticle(15026)
            Dim IMKey As String = .GetArticleImbaseKey()
        End With
    End Sub

    Private Sub ColortestToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ColorOptions.ShowDialog()
    End Sub

    Private Sub LuctestToolStripMenuItem_Click(sender As Object, e As EventArgs) 
        LicForm2.ShowDialog()
    End Sub

    Public Sub CompareMaterialInVSNRMWithMaterialTable()
        Try
            If VSNRM_path Is Nothing Then select_Excelfile_VSNRM()

            If VSNRM_path Is Nothing Then Exit Sub
            xlApp = New Excel.Application()
            xlApp.Visible = excel_visible
            WB = xlApp.Workbooks.Open(VSNRM_path)
            WS = WB.Sheets.Item(Part_Arr)
            Open_EX_Doc(False, material_table_path)

            progress_txt_im_ListBox1("Запущен процесс сверки " & WB.Name & " с таблицей соответствия материалов и покупных изделий(ТМПИ)")

            Dim MaterialUncnowCount, PurchatedUncnowCount As Integer
            Dim rowEnd As Integer = WS.Cells(WS.Rows.Count, Naim_col_number).End(Excel.XlDirection.xlUp).Row
            For rowNum As Integer = 5 To rowEnd
                Dim art_type As String = WS.Cells(rowNum, TypDSE_col_number).value
                Dim FindTextInVSNRM As String
                If art_type = "Деталь" Then
                    FindTextInVSNRM = WS.Cells(rowNum, Material_col_number).value
                    If Compare_PurchatedAndMaterial_with_IMBKey(FindTextInVSNRM, Materials_SheetName_TableMaterialsAndPurchates) Then MaterialUncnowCount += 1 : progress_txt_im_ListBox1("Материал (" & FindTextInVSNRM & ") не найден!")
                ElseIf art_type = "Покупное изделие" Then
                    FindTextInVSNRM = WS.Cells(rowNum, Naim_col_number).value
                    If Compare_PurchatedAndMaterial_with_IMBKey(FindTextInVSNRM, Purchates_SheetName_TableMaterialsAndPurchates) Then PurchatedUncnowCount += 1 : progress_txt_im_ListBox1("Покупное изделие (" & FindTextInVSNRM & ") не найдено!")
                End If
            Next
            excel_docs_close()
            Dim messagText As String
            If MaterialUncnowCount > 0 Or PurchatedUncnowCount > 0 Then
                Stop_Process = True

                progress_txt_im_ListBox1("Сверка завершена! " & "В таблице соответствия не найдены:" & MaterialUncnowCount & " - материал(ов); " & PurchatedUncnowCount & " - покупных изделий")

                messagText = "Проверка завершена!" & vbNewLine & "В таблице соответствия не найдены:" & vbNewLine & MaterialUncnowCount & " - материал(ов)" & vbNewLine & PurchatedUncnowCount & " - покупных изделий" & vbNewLine & "Открыть Таблицу для внесения изменений?"
                Dim result As Integer = MessageBox.Show(messagText, "Результат сверки ВСНРМ с таблицой соответсвия", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    excel_docs_close()
                    Exit Sub
                ElseIf result = DialogResult.Yes Then
                    Options.openMaterialAndPurchtedTable(True, material_table_path)
                End If
            Else

                progress_txt_im_ListBox1("Сверка завершена! " & "Для указанной ВСНРМ, В таблице соответствия имеются все необходимая ссылочная(из IMBase) информация!")

                messagText = "Проверка завершена!" & vbNewLine & "Для указанной ВСНРМ, В таблице соответствия имеются все необходимая ссылочная(из IMBase) информация!"
                MsgBox(messagText)
                Stop_Process = False
            End If

        Catch ex As Exception
            MsgBox("Не удалось запустить процедуру! Описание ошибки: " & ex.Message)
            progress_txt_im_ListBox1("Процесс прерван! Ошибка! " & ex.Message)
            excel_docs_close()
        End Try
    End Sub
    Public Function PurchatedAndMaterial_with_IMBKey(Material_Or_Purchated_VSNRM As String) As String
        Dim material_imbase, imb_kase, razdel, sheetName As String
        sheetName = Purchates_SheetName_TableMaterialsAndPurchates
        Dim material_imbase_index As Integer
        material_imbase_index = get_value_bay_FindText(sheetName, "A2", "A" & Get_LastRowInOneColumn(sheetName, Poln_oboz_VSNRM_col_number), Material_Or_Purchated_VSNRM)
        If material_imbase_index <> 0 Then
            material_imbase = get_Value_From_Cell(sheetName, Sootvetstvie_Imbase_col_number, material_imbase_index)
            imb_kase = get_Value_From_Cell(sheetName, IMBASE_Key_col_number, material_imbase_index)
            razdel = get_Value_From_Cell(sheetName, RazdelSP_col_number, material_imbase_index)
            PurchatedAndMaterial_with_IMBKey = material_imbase & "&" & imb_kase & "&" & razdel
        Else
            set_Value_From_Cell(sheetName, Poln_oboz_VSNRM_col_number, Get_LastRowInOneColumn(sheetName, Poln_oboz_VSNRM_col_number) + 1, Material_Or_Purchated_VSNRM)
            exc_WB1_save()
        End If
    End Function
    Public Function Compare_PurchatedAndMaterial_with_IMBKey(Material_Or_Purchated_VSNRM As String, sheetName As String) As Boolean
        Compare_PurchatedAndMaterial_with_IMBKey = False
        'Dim material_imbase, imb_kase As String
        Dim material_imbase_index As Integer
        material_imbase_index = get_value_bay_FindText(sheetName, "A2", "A" & Get_LastRowInOneColumn(sheetName, Poln_oboz_VSNRM_col_number), Material_Or_Purchated_VSNRM)
        If material_imbase_index <> 0 Then
            'material_imbase = get_Value_From_Cell(sheetName, Sootvetstvie_Imbase_col_number, material_imbase_index)
            'imb_kase = get_Value_From_Cell(sheetName, IMBASE_Key_col_number, material_imbase_index)
            'Compare_PurchatedAndMaterial_with_IMBKey = True
            'Compare_PurchatedAndMaterial_with_IMBKey = material_imbase & "&^" & imb_kase
        Else
            set_Value_From_Cell(sheetName, Poln_oboz_VSNRM_col_number, Get_LastRowInOneColumn(sheetName, Poln_oboz_VSNRM_col_number) + 1, Material_Or_Purchated_VSNRM)
            exc_WB1_save()
            Compare_PurchatedAndMaterial_with_IMBKey = True
        End If
    End Function
    Sub select_Excelfile_VSNRM()
        Dim openFileDialog1 As New OpenFileDialog()


        'openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = "excel files (*.xlsx)|*.xls|All files (*.*)|*.*"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                new_doc.ExcelFullFileName = openFileDialog1.FileName
                VSNRM_path = openFileDialog1.FileName
            Catch Ex As Exception
                MessageBox.Show("Не удается открыть файл. Код ошибки: " & Ex.Message)
                exc_close()
                xlApp.Quit()
                xlApp = Nothing
                progress_txt_im_ListBox1("Процесс завершился с ошибкой")
            End Try
        End If
    End Sub
End Class
