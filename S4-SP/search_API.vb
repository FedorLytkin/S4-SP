Imports System
Imports System.IO
Imports System.Text
Public Class search_API
    Public ITS4App As S4.TS4App
    Public SbServer As Object
    Public path As String
    Public sw As StreamWriter
    Sub login() 'ОБЯЗАТЕЛЬНО ВЫЗВАТЬ В НАЧАЛЕ, ЭТУ ПРОЦЕДУРУ, ЧТОБЫ ОПРЕДЕЛИТЬ ПЕРЕМЕННУЮ ITS4App
        ITS4App = CreateObject("S4.TS4App")
        ITS4App.Login()
    End Sub
    Sub loginSBServer() 'ОБЯЗАТЕЛЬНО ВЫЗВАТЬ В НАЧАЛЕ, ЭТУ ПРОЦЕДУРУ, ЧТОБЫ ОПРЕДЕЛИТЬ ПЕРЕМЕННУЮ SbServer
        SbServer = ITS4App.GetSbServer()
    End Sub

    '\\\\\\\\\\\\\\\\\ФУНКЦИИ ПО РАБОТЕ С ПОЛЬЗОВАТЕЛЯМИ\\\\\\\\\\\\\\\\\\
    Function USERS_INFO()  'ВОЗВРАЩАЕТ СПИСОК ПОЛЬЗОВАТЕЛЕЙ И ИХ UserID
        Dim FIO_array(40, 2) As String
        Dim arraycount As Integer = 40
        Dim FIO As String
        For i = 0 To arraycount
            FIO = ITS4App.GetUserFullName_ByUserID(i)
            FIO_array(i, 1) = i
            FIO_array(i, 2) = FIO
        Next
        USERS_INFO = FIO_array
    End Function
    Function getUserID_inUSERS_INFO(UserName As String) As String
        Dim USERS_FIO(,) As String
        USERS_FIO = USERS_INFO()

        For i As Integer = 0 To UBound(USERS_FIO)
            If USERS_FIO(i, 2) = UserName Or Strings.InStr(USERS_FIO(i, 2), UserName) <> 0 Then
                getUserID_inUSERS_INFO = i
                Exit For
            End If
        Next
    End Function
    Function UserID() As Integer 'возвращает порядковый номер пользователья
        UserID = ITS4App.GetUserID()
    End Function
    Function GetUserName() As String 'возвращает имя пользователья 'SYSDBA'
        GetUserName = ITS4App.GetUserName()
    End Function
    Function GetUserFullName_ByUserID_comp() As String 'возвращает имя пользователья 'Системный администратор' НА ТЕКУЩЕМ КОМПЕ!!!
        Dim GetUserID As Integer = ITS4App.GetUserID
        GetUserFullName_ByUserID_comp = ITS4App.GetUserFullName_ByUserID(GetUserID)
    End Function
    Function GetUserFullName_ByUserID_Po_ID(UserID As Long) As String 'возвращает имя пользователья 'Системный администратор' НА ТЕКУЩЕМ КОМПЕ!!!
        GetUserFullName_ByUserID_Po_ID = ITS4App.GetUserFullName_ByUserID(UserID)
    End Function
    '\\\\\\\\\\\\\\\\\ОТМЕЧЕННЫЕ ДОКУМЕНТЫ В СЕРЧЕ\\\\\\\\\\\\\\\\\\
    Function SelectedCount(GetSelectedItems As S4.TS4SelectedItems)
        SelectedCount = GetSelectedItems.SelectedCount
    End Function
    Function ActiveDocID() As Integer 'инвентарный номер отмеченного документа
        ActiveDocID = ITS4App.ActiveDocID
    End Function
    '\\\\\\\\\\\\\\\\\ПАРАМЕТРЫ ДОКУМЕНТА(ОСНОВНАЯ НАДПИСЬ ЧЕРТЕЖА)\\\\\\\\\\\\\\\\
    Function Get_param_Obozn(DocID As Integer) As String 'поле Обозначение отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_Obozn = ITS4App.GetFieldValue("Обозначение")
    End Function
    Function Get_param_Naim(DocID As Integer) As String 'поле Наименование отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_Naim = ITS4App.GetFieldValue("Наименование")
    End Function
    Function Get_param_Doc_Type(DocID As Integer) As String 'поле Тип документа отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_Doc_Type = ITS4App.GetFieldValue("Тип документа")
    End Function
    Function Get_param_Razrab(DocID As Integer) As String 'поле разработал отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_Razrab = ITS4App.GetFieldValue("Разработал")
    End Function
    Function Get_param_Prov(DocID As Integer) As String 'поле Пров отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_Prov = ITS4App.GetFieldValue("Проверил")
    End Function
    Function Get_param_NormControl(DocID As Integer) As String 'поле Н.контр отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_NormControl = ITS4App.GetFieldValue("Н.контр.")
    End Function
    Function Get_param_TechControl(DocID As Integer) As String 'поле Т.контр отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_TechControl = ITS4App.GetFieldValue("Т.контр.")
    End Function
    Function Get_param_Utverdil(DocID As Integer) As String 'поле Утвердил отмеченного документа
        ITS4App.OpenDocument(DocID)
        Get_param_Utverdil = ITS4App.GetFieldValue("Утвердил")
    End Function


    '\\\\\\\\\\\\\\\\\\SQL-запрос. функция возвращает значение конкретной ячейки\\\\\\\\\\\\\\\\\\\\\\\\\
    Function query_s4(table_name As String, column_name As String, find_column_name As String, find_text As String)
        ITS4App.OpenQuery("select * from " & table_name & " where " & column_name & " = '" & find_text & "'")
        query_s4 = ITS4App.QueryFieldByName(find_column_name)
        ITS4App.CloseQuery()
    End Function


    '\\\\\\\\\\\\\\\\\\ПАРАМЕТРЫ ОБЪЕКТА(ВКЛАДКА ОБЪЕКТА КАРТОЧКИ ДОКУМЕНТА)\\\\\\\\\\\\\\\\\\\\\\\\\
    Function GetArticlesCount(DocID As Integer) As Integer  'получаем количество вариантов исполнений изделия по данному документу
        ITS4App.OpenDocArticles(DocID)
        GetArticlesCount = ITS4App.GetArticlesCount
    End Function
    Function GetDocArticleID(i As Integer) As Integer  'получаем инвентарный номер текущего исполнения ОБЪЕКТА
        GetDocArticleID = ITS4App.GetDocArticleID(i)
    End Function


    '\\\\\\\\\\\\\\\\\\Список подписей(ВКЛАДКА ПОДПИСИ, КАРТОЧКИ ДОКУМЕНТА)\\\\\\\\\\\\\\\\\\\\\\\\\
    'RAZRAB
    Function GetSignedRazrab_ID(DocID As Integer) 'ВОЗВРАЩАЕТ id ГРАФЫ РАЗРАБОТАЛ
        GetSignedRazrab_ID = ITS4App.GetSignedRankUserIDs(DocID, 0, "7")
    End Function
    Function GetSignedRazrab_FIO(DocID As Integer) 'ВОЗВРАЩАЕТ ФИО ГРАФЫ РАЗРАБОТАЛ
        Dim ID() As Long = ITS4App.GetSignedRankUserIDs(DocID, 0, "7")
        GetSignedRazrab_FIO = GetUserFullName_ByUserID_Po_ID(ID(0))
    End Function
    'PROV
    Function GetSignedpROV_ID(DocID As Integer) 'ВОЗВРАЩАЕТ id ГРАФЫ ПРОВЕРИЛ
        GetSignedpROV_ID = ITS4App.GetSignedRankUserIDs(DocID, 0, ">")
    End Function
    Function GetSignedROV_FIO(DocID As Integer) 'ВОЗВРАЩАЕТ ФИО ГРАФЫ ПРОВЕРИЛ
        Dim ID() As Long = ITS4App.GetSignedRankUserIDs(DocID, 0, ">")
        GetSignedROV_FIO = GetUserFullName_ByUserID_Po_ID(ID(0))
    End Function
    'NORMCONTROL
    Function GetSignedpNorm_ID(DocID As Integer) 'ВОЗВРАЩАЕТ id ГРАФЫ НОРМОКОНТРОЛЬ
        GetSignedpNorm_ID = ITS4App.GetSignedRankUserIDs(DocID, 0, "4")
    End Function
    Function GetSignedNorm_FIO(DocID As Integer) 'ВОЗВРАЩАЕТ ФИО ГРАФЫ НОРМОКОНТРОЛЬ
        Dim ID() As Long = ITS4App.GetSignedRankUserIDs(DocID, 0, "4")
        GetSignedNorm_FIO = GetUserFullName_ByUserID_Po_ID(ID(0))
    End Function
    'TECHCONTROL
    Function GetSignedpTech_ID(DocID As Integer) 'ВОЗВРАЩАЕТ id ГРАФЫ ТЕХКОНТРОЛЬ
        GetSignedpTech_ID = ITS4App.GetSignedRankUserIDs(DocID, 0, "4")
    End Function
    Function GetSignedTech_FIO(DocID As Integer) 'ВОЗВРАЩАЕТ ФИО ГРАФЫ ТЕХКОНТРОЛЬ
        Dim ID() As Long = ITS4App.GetSignedRankUserIDs(DocID, 0, "4")
        GetSignedTech_FIO = GetUserFullName_ByUserID_Po_ID(ID(0))
    End Function
    'UTVERDIL
    Function GetSignedpUtv_ID(DocID As Integer) 'ВОЗВРАЩАЕТ id ГРАФЫ УТВЕРДИЛ
        GetSignedpUtv_ID = ITS4App.GetSignedRankUserIDs(DocID, 0, "4")
    End Function
    Function GetSignedUtv_FIO(DocID As Integer) 'ВОЗВРАЩАЕТ ФИО ГРАФЫ УТВЕРДИЛ
        Dim ID() As Long = ITS4App.GetSignedRankUserIDs(DocID, 0, "4")
        GetSignedUtv_FIO = GetUserFullName_ByUserID_Po_ID(ID(0))
    End Function


    Public Sub Report_txt(filename As String)
        path = "C:\IM\IMWork\" & filename & ".txt"
        If Not File.Exists(path) Then
            sw = File.CreateText(path)
        End If
    End Sub
    Public Sub Write_Report_txt(text As String)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
        file.WriteLine(text)
        file.Close()
    End Sub
End Class
