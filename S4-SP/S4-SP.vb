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
    Public Poln_oboz_VSNRM_col_number As Integer = 1
    Public Sootvetstvie_Imbase_col_number As Integer = 2
    Public IMBASE_Key_col_number As Integer = 3
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


    Public DocNames As New ListView.TreeListNode
    Public Sub hello()
        MsgBox("hello,world!")
    End Sub
    Private Sub ВыбратьОбъектToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыбратьОбъектToolStripMenuItem.Click
        select_VSNRM()
    End Sub
    Sub progress_txt_im_ListBox1(progress_txt As String)
        ListBox1.Items.Add(progress_txt)
    End Sub
    Sub select_VSNRM()
        Dim openFileDialog1 As New OpenFileDialog()

        progress_txt_im_ListBox1("Подготовка к созданию древовидной структуры")

        'openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = "excel files (*.xlsx)|*.xls|All files (*.*)|*.*"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                VSNRM_path = openFileDialog1.FileName
                If VSNRM_path IsNot Nothing And VSNRM_path.IndexOf("ВСНРМ") <> 0 And VSNRM_path.IndexOf("xls") <> 0 Then
                    xlApp = New Excel.Application()
                    xlApp.Visible = False
                    WB = xlApp.Workbooks.Open(VSNRM_path)
                    WS = WB.Sheets.Item(Part_Arr)

                    progress_txt_im_ListBox1("Файл ВСНРМ: " & VSNRM_path)

                    Dim str As String = WS.Cells(1, 1).Value()
                    Dim naim As String = str.Substring(str.LastIndexOf("   ")).Trim(" ")
                    Dim oboz As String = str.Substring(0, str.IndexOf("   "))
                    Dim root = New TreeNode(WS.Cells(1, 1).Value())

                    TreeView1.Nodes.Clear()
                    TreeView1.Nodes.Add(root)
                    recvAdd_Only_treeview(root, 5, True, WS.Cells(WS.Rows.Count, Naim_col_number).End(XlDirection.xlUp).Row)
                    progress_txt_im_ListBox1("Процесс завершен! Древовидная структура построено")
                Else
                    Exit Sub
                End If
            Catch Ex As Exception
                MessageBox.Show("Не удается открыть файл. Код ошибки: " & Ex.Message)
                progress_txt_im_ListBox1("Процесс завершился с ошибкой")
            End Try
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Search_Subdirectories = Convert.ToBoolean(get_reesrt_value("Search_Subdirectories", True.ToString))
        material_table_path = get_reesrt_value("material_table_path", System.Windows.Forms.Application.StartupPath & "\material_with_ImBaseKey.xlsx")
        doctypes_array = get_reesrt_value("doctypes_array", Options.doctypes_array_default)
        note_list_check_FFm = get_reesrt_value("note_list_check_FFm", Options.note_list_check_FFm)
        excel_visible = get_reesrt_value("excel_visible", False.ToString)
        Default_work_path = get_reesrt_value("Default_work_path", Default_work_path)

        Dim err_massage As String
        Try
            ITS4App = CreateObject("S4.TS4App")
            ITS4App.Login()
            Label1.Text = ITS4App.GetUserFullName_ByUserID(ITS4App.GetUserID)
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
        End If
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
                    exc_close()
                    Exit Sub
                ElseIf result = DialogResult.Yes Then
                    progress_txt_im_ListBox1("В ВСНРМ не найден лист " & Docs_arr & "! Продолжение процедуры")
                    Exit Try
                End If
            End Try
            Open_EX_Doc(False, material_table_path)

            Dim str As String = WS.Cells(1, 1).Value()
            Dim naim As String = str.Substring(str.LastIndexOf("   ")).Trim(" ")
            Dim oboz As String = str.Substring(0, str.IndexOf("   "))
            Dim root = New TreeNode(WS.Cells(1, 1).Value())

            'запустить создание объекта в с4 (Сборочная единица!). Вернет ее PartAID (понадобится следующем уровне)
            'PartAID записать в тегах текущей ветки
            root.Tag = ITS4App.AddNewArticle2(oboz.Replace("СБ", ""), "", naim, 3)
            If root.Tag > 0 Then
                progress_txt_im_ListBox1("Создан объект: ArtID:" & root.Tag & " * " & oboz & " - " & naim)
            End If
            If only_Articles = False Then
                Dim doc_filename As String = Doc_list_path & "\" & oboz & " " & naim & "." & fileformat_for_doctypes(5)
                'Dim doc_type As Integer = query_s4("doctypes", "DOC_NAME", "DOC_TYPE", detal_doc_types)
                'If File.Exists(doc_filename) Then
                '    DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, SB_doc_types, Arch_ID, oboz, naim)
                'End If
                If Not File.Exists(doc_filename) Then
                    doc_filename = find_file_by_ObozNainFformat(oboz, naim, fileformat_for_doctypes(5))
                End If
                DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, doctypes(5), Arch_ID, oboz, naim)
                If DocID > 0 Then
                    progress_txt_im_ListBox1("Создан документ: DocID:" & DocID & " * " & oboz & " - " & naim)
                    create_New_Version_InDoc(doc_filename, DocID)
                End If
            End If
            Call add_Bom_Node_In_Documentacia2(oboz, root.Tag)
            If WS_DOC IsNot Nothing Then
                'тут будет процедуры создающая документы из листа ДОКУМЕНТАЦИЯ в серч
                MsgBox("Documentacia")
            End If
            TreeView1.Nodes.Clear()
            TreeView1.Nodes.Add(root)
            LastInd = WS.Cells(WS.Rows.Count, Naim_col_number).End(Excel.XlDirection.xlUp).Row

            'прогрессбар появление 
            'Dim SelectedCount As Integer = LastInd - 5
            'Dim textProgress As String = "Обработано 0 из " & SelectedCount
            ShowProgressBarForm(LastInd - 4, "Создание Дерево состава и Занесение в архив", "Занесение в архив документов из ВСНРМ")

            recvAdd(root, 5, True)
            ITS4App.CloseProgressBarForm()
            If only_Articles = False Then add_SP_InBOM(root.Tag) 'процедура по созданию документов SP на сборочные единицы

            progress_txt_im_ListBox1("Процесс завершен! Создано изделие: ArtID: " & root.Tag & " * " & oboz & " - " & naim)
            excel_docs_close()
        Catch ex As Exception
            MsgBox("Не удалось запустить процедуру! Описание ошибки: " & ex.Message)
            progress_txt_im_ListBox1("Процесс прерван! Ошибка! " & ex.Message)
            ITS4App.CloseProgressBarForm()
        End Try
    End Sub
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
    Private Function recvAdd(node As TreeNode, rownum As Integer, isFirst As Boolean)
        'прогресс бар заполнение 
        SetProgressBarData_and_CheckUserBreak(rownum - 5, LastInd - 4)

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
                If only_Articles = False Then DocID = add_document(rownum)
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
            SetProgressBarData_and_CheckUserBreak(rownum - 4, LastInd - 4)

            Dim nextnumpp As String = WS.Cells(rownum, NumPP_col_number).Value()
            If nextnumpp Is Nothing Then
                Exit While
            End If
            Dim ptrn2 As New Regex("^" + numpp + "\.[0-9]+$")
            If ptrn2.IsMatch(nextnumpp) Then
                rownum = recvAdd(newNode, rownum, False)
                numpp = WS.Cells(rownum, NumPP_col_number).Value()
            Else
                numpp = nextnumpp
            End If
        End While
        Return rownum
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
            progress_txt_im_ListBox1("Объект: ArtID: " & DocID & " * " & designation & " - " & Naim & " уже существует в S4")
            Exit Function
        End If
        ITS4App.OpenArticle(new_ArtID)
        If material IsNot Nothing Then
            Call ITS4App.SetFieldValue_Articles("Материал", material_with_IMBKey(material))
        End If
        If tolshina_diameter IsNot Nothing Then
            Call ITS4App.SetFieldValue_Articles("Толщина/Диаметр", tolshina_diameter)
        End If
        If dlina IsNot Nothing And Razdel <> 3 Then
            Call ITS4App.SetFieldValue_Articles("Примечание", "L=" & dlina)
        End If
        Call ITS4App.SetFieldValue_Articles("Масса", mass)
        progress_txt_im_ListBox1("Создан объект: ArtID: " & DocID & " * " & designation & " - " & Naim)
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
        End If
        DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, doc_type, Arch_ID, designation, Naim)
        'формат страницы документа заполнится тут
        If DocID > 0 Then
            Dim format_doc As String = WS.Cells(rownum, format_col_number).value
            If format_doc IsNot Nothing Then
                ITS4App.OpenDocument(DocID)
                ITS4App.SetFieldValue("Формат", format_doc)
                ITS4App.CloseDocument()
            End If
            progress_txt_im_ListBox1("Создан документ: DocID: " & DocID & " * " & designation & " - " & Naim)
            'здесь будет процедура по созданию/изменению версии документа 
            create_New_Version_InDoc(doc_filename, DocID)
            'For Each note_check_FFm As String In note_list_check_FFm.Split(";")
            '    note_check_FFm = note_check_FFm.Trim(" ")
            '    If doc_filename.IndexOf("(" & note_check_FFm) > 0 Then
            '        Dim new_vers As String = doc_filename.Substring(doc_filename.IndexOf(note_check_FFm) + note_check_FFm.Length)
            '        new_vers = new_vers.Substring(0, new_vers.LastIndexOf(")"))
            '        ITS4App.OpenDocument(DocID)
            '        ITS4App.SetDocVersionCode(new_vers)
            '        ITS4App.CloseDocument()
            '    End If
            'Next
        End If

        ''приведенный ниже код создает документ спецификации. НО! он не работает, потому что после его создания невозможна правка состава изделия
        'If WS.Cells(rownum, TypDSE_col_number).value = "Сборочная единица" Then
        '    ITS4App.CreateNewDocument(Arch_ID, 1, designation.Replace("СБ", ""), Naim, 0)
        'End If
        Return DocID
    End Function
    Sub create_New_Version_InDoc(doc_filename As String, docId As String)
        For Each note_check_FFm As String In note_list_check_FFm.Split(";")
            note_check_FFm = note_check_FFm.Trim(" ")
            If doc_filename.IndexOf("(" & note_check_FFm) > 0 Then
                Dim old_vers As String = ITS4App.GetFieldValue("Изменение")
                Dim new_vers As String = doc_filename.Substring(doc_filename.IndexOf(note_check_FFm) + note_check_FFm.Length)
                new_vers = new_vers.Substring(0, new_vers.LastIndexOf(")"))
                ITS4App.OpenDocument(docId)
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
        End If
        DocID = ITS4App.CreateFileDocumentWithDocType1(doc_filename, doc_type, Arch_ID, designation, Naim)
        'формат страницы документа заполнится тут
        If DocID > 0 Then
            Dim format_doc As String = WS_DOC.Cells(rownum, format_Doc_col_number).value
            If format_doc IsNot Nothing Then
                ITS4App.OpenDocument(DocID)
                ITS4App.SetFieldValue("Формат", format_doc)
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
                .OpenArticle(.GetArtID_ByDocID(add_document_on_Isponenie))
                .SetFieldValue_Articles("Масса", WS.Cells(Mass_col_number, search_rownum).value)
                .SetFieldValue_Articles("Толщина/Диаметр", WS.Cells(Tolshina_Diameter_col_number, search_rownum).value)
                .SetFieldValue_Articles("Материал", material_with_IMBKey(WS.Cells(Material_col_number, search_rownum).value))
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
        Dim sheetName As String = "Materials"
        Dim material_imbase_index As Integer
        material_imbase_index = get_value_bay_FindText(sheetName, "A3", "A" & Get_LastRowInOneColumn(sheetName, 1), material_VSNRM)
        If material_imbase_index <> 0 Then
            material_imbase = get_Value_From_Cell(sheetName, Sootvetstvie_Imbase_col_number, material_imbase_index)
            imb_kase = get_Value_From_Cell(sheetName, IMBASE_Key_col_number, material_imbase_index)
            material_with_IMBKey = material_imbase & "&^" & imb_kase
        Else
            set_Value_From_Cell(sheetName, Poln_oboz_VSNRM_col_number, Get_LastRowInOneColumn(sheetName, 1) + 1, material_VSNRM)
        End If
    End Function
    Function check_in_BOM_exist(ProjAID As Integer, PartAID As Integer, position As String) As Boolean
        ITS4App.OpenArticleStructure(ProjAID)
        ITS4App.asFirst()
        While ITS4App.asEof = 0
            If ITS4App.asGetArtID = PartAID And ITS4App.asGetPosition = position Then
                Return True
                Exit While
            End If
            ITS4App.asNext()
        End While
        ITS4App.CloseArticleStructure()
        Return False
    End Function
    Sub add_SP_InBOM(ProjAID As Integer)
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
        Dim Razdel As Integer = query_s4("ssections", "art_kind", "SECTION_ID", WS.Cells(rownum, TypDSE_col_number).value)
        Dim position As Integer = CInt(WS.Cells(rownum, NumPP_col_number).value.ToString.Substring(WS.Cells(rownum, NumPP_col_number).value.ToString.LastIndexOf(".") + 1))
        Dim Note As String = WS.Cells(rownum, Prim_col_number).value
        If check_in_BOM_exist(ProjAID, PartAID, position) = False Then
            ITS4App.AddBOMItem(ProjAID, PartAID, CountPC, MuID, Razdel, position, Note)
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
        Dim CountPC As Integer = "" ' WS.Cells(rownum, KolNaEd_col_number).value
        Dim MuID As Integer = query_s4("MU", "MU_SHORT_NAME", "MI_ID", MU_part)
        Dim Razdel As Integer = 1
        Dim position As String = ""
        Dim Note As String = ""
        If check_in_BOM_exist(ProjAID, PartAID, position) = False Then
            ITS4App.AddBOMItem(ProjAID, PartAID, CountPC, MuID, Razdel, position, Note)
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
        Dim CountPC As Integer = WS.Cells(rownum, KolNaEd_col_number).value
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        GroupBox3.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GroupBox3.Enabled = False
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
        exc_close()
    End Sub

    Private Sub НастройкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НастройкиToolStripMenuItem.Click
        Options.ShowDialog()
    End Sub

    Private Sub TreeView1_Click(sender As Object, e As EventArgs) Handles TreeView1.Click
        Dim node_position As String = TreeView1.SelectedNode.Text

    End Sub

    Private Sub СоздатьДеревоToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СоздатьДеревоToolStripMenuItem.Click
        new_doc.ShowDialog()
    End Sub

    Public Sub SelectArticl_procedure()
        ITS4App.StartSelectArticles()
        ITS4App.SelectArticles()
        If ITS4App.SelectedArticlesCount = 0 Then
            Exit Sub
        End If
        ArtID = ITS4App.GetSelectedArticleID(0)
        ITS4App.EndSelectArticles()
    End Sub
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
            Dim ptrn2 As New Regex("^" + numpp + "\.[0-9]+$")
            If ptrn2.IsMatch(nextnumpp) Then
                rownum = recvAdd_Only_treeview(newNode, rownum, False, totalNum)
                numpp = WS.Cells(rownum, 1).Value()
            Else
                numpp = nextnumpp
            End If
            If Regex.IsMatch(numpp, "10$") Then
                totalNum += 1
                totalNum -= 1
            End If
        End While
        Return rownum
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
End Class
