Imports System.IO
Module Pr_Options
    Public Opt_File_Puth As String = Application.StartupPath & "\CFG\Pr_Option.ini"

    Public Opt_Del_Char As String = "/"
    Public Opt_Comment_Char As String = "##"
    Public Opt_Color_Del_Char As String = ";"


    'параметр не выводящий типы объектов и компоненты с ручной связью
    Public NO_Art_Documentacia As Boolean = 0
    Public NO_Art_Documentacia_OpName As String = "NO_Art_Documentacia"
    Public NO_Ru4na9_Sv9z As Boolean = 1
    Public NO_Ru4na9_Sv9z_OpName As String = "NO_Ru4na9_Sv9z"
    Public NO_Sborka As Boolean = 1
    Public NO_Sborka_OpName As String = "NO_Sborka"
    Public NO_PartSecion As Boolean = 1
    Public NO_PartSecion_OpName As String = "NO_PartSecion"
    Public NO_Standart As Boolean = 1
    Public NO_Standart_OpName As String = "NO_Standart"
    Public NO_Pro4ee As Boolean = 1
    Public NO_Pro4ee_OpName As String = "NO_Pro4ee"
    Public NO_Material As Boolean = 1
    Public NO_Material_OpName As String = "NO_Material"
    Public Part_SB As Boolean = 0
    Public Part_SB_OpName As String = "Part_SB"
    'параметр заменяющий "~" на " "
    Public ReplaceTildaOnSpace As Boolean = 0
    'параметр заменяющий "?" на "/"
    Public ReplaceQuestionOnDrob As Boolean = 0

    'параметр экспортирующий данные о первом уровне вложенности
    Public OnlyFirstLewvel As Boolean = 0
    Public OnlyFirstLewvel_OpName As String = "OnlyFirstLewvel"
    'параметр экспортирующий данные о СОСТАВ ИЗДЕЛИЯ
    Public NO_Sostav As Boolean = 1
    Public NO_Sostav_OpName As String = "NO_Sostav"
    'параметр экспортирующий данные о применяемости деталей
    Public NO_Parts As Boolean = 1
    Public NO_Parts_OpName As String = "NO_Parts"
    'параметр экспортирующий данные о материалах и покупных
    Public NO_Purchated As Boolean = 1
    Public NO_Purchated_OpName As String = "NO_Purchated"
    'параметр экспортирующий ПОДРОБНЫЕ данные о материалах и покупных
    Public NO_Purchated_PDRBN As Boolean = 1
    Public NO_Purchated_PDRBN_OpName As String = "NO_Purchated_PDRBN"
    'брать параметры только из серча
    Public NO_TechCard As Boolean = 1
    Public NO_TechCard_OpName As String = "NO_TechCard"

    'передавать параметры для АВА
    Public NO_AVA As Boolean = 1
    Public NO_AVA_OpName As String = "NO_AVA"

    'параметр экспортирующий данные о материалах и покупных
    Public NO_ColorNote As Boolean = 1
    Public NO_ColorNote_OpName As String = "NO_ColorNote"

    'параметр экспортирующий данные о материалах и покупных
    Public NO_ExlProc_Visible As Boolean = 0
    Public NO_ExlProc_Visible_OpName As String = "NO_ExlProc_Visible"

    'параметр экспортирующий данные серчовских параметрах
    Public NO_S4_Columns As Boolean = 0
    Public NO_S4_Columns_OpName As String = "NO_S4_Columns"

    'параметр экспортирующий данные серчовских параметрах
    Public NO_Color_Info As Boolean = 0
    Public NO_Color_Info_OpName As String = "NO_Color_Info"


    Public Select_Foldef_OpName As String = "Select_Foldef"
    Public Select_Foldef_Val As Integer = 4

    Public Create_Unfold_OpName As String = "Create_Unfold"
    Public Create_Unfold_Val As Integer

    Public AddLog_OpName As String = "AddLog"
    Public AddLog_Val As Integer
    Public Add_SuppresCom_OpName As String = "Add_SuppresCom"
    Public Add_SuppresCom_Val As Integer = 1
    Public Add_BOM_Log_OpName = "Add_BOM_Log_OpName" ' "Записывать Log о структуре файла"
    Public Add_BOM_Log_Val As Integer = 1
    Sub Save_Parameters()
        Dim strFile As String = Opt_File_Puth
        File.Delete(strFile)
        Dim fileExists As Boolean = File.Exists(strFile)
        Using sw As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate), System.Text.Encoding.GetEncoding(1251))
            '0-Параметр;1-Значение
            sw.WriteLine(NO_Art_Documentacia_OpName & Opt_Del_Char & NO_Art_Documentacia)
            sw.WriteLine(NO_Ru4na9_Sv9z_OpName & Opt_Del_Char & NO_Ru4na9_Sv9z)
            sw.WriteLine(NO_Sborka_OpName & Opt_Del_Char & NO_Sborka)
            sw.WriteLine(NO_PartSecion_OpName & Opt_Del_Char & NO_PartSecion)
            sw.WriteLine(NO_Standart_OpName & Opt_Del_Char & NO_Standart)
            sw.WriteLine(NO_Pro4ee_OpName & Opt_Del_Char & NO_Pro4ee)
            sw.WriteLine(NO_Material_OpName & Opt_Del_Char & NO_Material)
            sw.WriteLine(Part_SB_OpName & Opt_Del_Char & part_SB)
            sw.WriteLine(OnlyFirstLewvel_OpName & Opt_Del_Char & OnlyFirstLewvel)
            sw.WriteLine(NO_Sostav_OpName & Opt_Del_Char & NO_Sostav)
            sw.WriteLine(NO_Parts_OpName & Opt_Del_Char & NO_Parts)
            sw.WriteLine(NO_Purchated_OpName & Opt_Del_Char & NO_Purchated)
            sw.WriteLine(NO_Purchated_PDRBN_OpName & Opt_Del_Char & NO_Purchated_PDRBN)
            sw.WriteLine(NO_ColorNote_OpName & Opt_Del_Char & NO_ColorNote)
            sw.WriteLine(NO_ExlProc_Visible_OpName & Opt_Del_Char & NO_ExlProc_Visible)
            sw.WriteLine(NO_S4_Columns_OpName & Opt_Del_Char & NO_S4_Columns)
            sw.WriteLine(NO_TechCard_OpName & Opt_Del_Char & NO_TechCard)
            sw.WriteLine(NO_AVA_OpName & Opt_Del_Char & NO_AVA)
        End Using
    End Sub
    Sub chekced_ParamChancge(Parname As String, par_val As Boolean)
        Select Case Parname
            Case NO_Art_Documentacia_OpName
                NO_Art_Documentacia = CInt(par_val)
            Case NO_Ru4na9_Sv9z_OpName
                NO_Ru4na9_Sv9z = CInt(par_val)
            Case NO_Sborka_OpName
                NO_Sborka = CInt(par_val)
            Case NO_PartSecion_OpName
                NO_PartSecion = CInt(par_val)
            Case NO_Standart_OpName
                NO_Standart = CInt(par_val)
            Case NO_Pro4ee_OpName
                NO_Pro4ee = CInt(par_val)
            Case NO_Material_OpName
                NO_Material = CInt(par_val)
            Case Part_SB_OpName
                Part_SB = CInt(par_val)
            Case OnlyFirstLewvel_OpName
                OnlyFirstLewvel = CInt(par_val)
            Case NO_Sostav_OpName
                NO_Sostav = CInt(par_val)
            Case NO_Parts_OpName
                NO_Parts = CInt(par_val)
            Case NO_Purchated_OpName
                NO_Purchated = CInt(par_val)
            Case NO_Purchated_PDRBN_OpName
                NO_Purchated_PDRBN = CInt(par_val)
            Case NO_ColorNote_OpName
                NO_ColorNote = CInt(par_val)
            Case NO_ExlProc_Visible_OpName
                NO_ExlProc_Visible = CInt(par_val)
            Case NO_S4_Columns_OpName
                NO_S4_Columns = CInt(par_val)
            Case NO_TechCard_OpName
                NO_TechCard = CInt(par_val)
            Case NO_AVA_OpName
                NO_AVA = CInt(par_val)
        End Select
        Save_Parameters()
    End Sub
    Sub ReadOptionsAll()
        Try
            Using r As StreamReader = New StreamReader(Opt_File_Puth, System.Text.Encoding.GetEncoding(1251))
                Dim line As String
                line = r.ReadLine

                Do While (Not line Is Nothing)
                    If line.IndexOf(Opt_Comment_Char) <> 0 Then
                        Dim line_arr() As String = line.Split(Opt_Del_Char)
                        Select Case line_arr(0)
                            Case NO_Art_Documentacia_OpName
                                NO_Art_Documentacia = line_arr(1)
                            Case NO_Ru4na9_Sv9z_OpName
                                NO_Ru4na9_Sv9z = line_arr(1)
                            Case NO_Sborka_OpName
                                NO_Sborka = line_arr(1)
                            Case NO_PartSecion_OpName
                                NO_PartSecion = line_arr(1)
                            Case NO_Standart_OpName
                                NO_Standart = line_arr(1)
                            Case NO_Pro4ee_OpName
                                NO_Pro4ee = line_arr(1)
                            Case NO_Material_OpName
                                NO_Material = line_arr(1)
                            Case Part_SB_OpName
                                Part_SB = line_arr(1)
                            Case OnlyFirstLewvel_OpName
                                OnlyFirstLewvel = line_arr(1)
                            Case NO_Sostav_OpName
                                NO_Sostav = line_arr(1)
                            Case NO_Parts_OpName
                                NO_Parts = line_arr(1)
                            Case NO_Purchated_OpName
                                NO_Purchated = line_arr(1)
                            Case NO_Purchated_PDRBN_OpName
                                NO_Purchated_PDRBN = line_arr(1)
                            Case NO_ColorNote_OpName
                                NO_ColorNote = line_arr(1)
                            Case NO_ExlProc_Visible_OpName
                                NO_ExlProc_Visible = line_arr(1)
                            Case NO_S4_Columns_OpName
                                NO_S4_Columns = line_arr(1)
                            Case NO_TechCard_OpName
                                NO_TechCard = line_arr(1)
                            Case NO_AVA_OpName
                                NO_AVA = line_arr(1)
                        End Select
                    End If
                    line = r.ReadLine
                Loop
            End Using
        Catch ex As Exception

        End Try
    End Sub

End Module
