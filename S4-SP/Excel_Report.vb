Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Module Excel_Report
    'Индексы столбцов отчета в ексел
    Public ArtID_ColNum As Integer = 1
    Public DocID_ColNum As Integer = 2
    Public ProjID_ColNum As Integer = 3
    Public Oboz_ColNum As Integer = 4
    Public Naim_ColNum As Integer = 5
    Public Kolvo_ColNum As Integer = 6
    Public Material_ColNum As Integer = 7
    Public tolshinaDiam_ColNum As Integer = 8
    Public Mass_ColNum As Integer = 9
    Public Dlina_ColNum As Integer = 10
    Public Prim_ColNum As Integer = 11
    Public ArtType_ColNum As Integer = 12
    Public DocType_ColNum As Integer = 13
    Public Format_ColNum As Integer = 14
    Public Izm_ColNum As Integer = 14

    'параметры нового документы ексель
    Public rprt_XlsApp As Excel.Application
    Public rprt_WB As Excel.Workbook
    Public rprt_WS As Excel.Worksheet

    Public Sub Create_EX_report(Visible As Double)
        rprt_XlsApp = New Excel.Application()
        rprt_XlsApp.Visible = Visible
        rprt_WB = xlApp.Workbooks.Add(1)
        rprt_WS = WB.Sheets(1)
        With rprt_WS
            .Cells(1, ArtID_ColNum) = "ArtID"
            .Cells(1, DocID_ColNum) = "DocID"
            .Cells(1, ProjID_ColNum) = "ProjID"
            .Cells(1, Oboz_ColNum) = "Обозначение"
            .Cells(1, Naim_ColNum) = "Наименование"
            .Cells(1, Kolvo_ColNum) = "Кол-во"
            .Cells(1, Material_ColNum) = "Материал"
            .Cells(1, tolshinaDiam_ColNum) = "Толщина/Диаметр"
            .Cells(1, Mass_ColNum) = "Масса"
            .Cells(1, Dlina_ColNum) = "Длина"
            .Cells(1, Prim_ColNum) = "Примечание"
            .Cells(1, ArtType_ColNum) = "Тип объекта"
            .Cells(1, DocType_ColNum) = "Тип документа"
            .Cells(1, Format_ColNum) = "Формат документа"
            .Cells(1, Izm_ColNum) = "Изменение"
        End With
    End Sub
    Public Sub exc_rprt_edit(ArtId As String, DocID As String, ProjID As String,
                             Oboz As String, Naim As String, Kolvo As String, Material As String,
                             tolshinaDiam As String, Mass As String, Dlina As String, Prim As String,
                             ArtType As String, DocType As String, Format As String, Izm As String)
        With rprt_WS
            Dim last_index As Integer = .Cells(.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1
            .Cells(last_index, ArtID_ColNum) = ArtId
            .Cells(last_index, DocID_ColNum) = DocID
            .Cells(last_index, ProjID_ColNum) = ProjID
            .Cells(last_index, Oboz_ColNum) = Oboz
            .Cells(last_index, Naim_ColNum) = Naim
            .Cells(last_index, Kolvo_ColNum) = Kolvo
            .Cells(last_index, Material_ColNum) = Material
            .Cells(last_index, tolshinaDiam_ColNum) = tolshinaDiam
            .Cells(last_index, Mass_ColNum) = Mass
            .Cells(last_index, Dlina_ColNum) = Dlina
            .Cells(last_index, Prim_ColNum) = Prim
            .Cells(last_index, ArtType_ColNum) = ArtType
            .Cells(last_index, DocType_ColNum) = DocType
            .Cells(last_index, Format_ColNum) = Format
            .Cells(last_index, Izm_ColNum) = Izm
        End With
    End Sub
    Public Sub exc_rprt_close()
        Try
            rprt_XlsApp.Quit()
            rprt_XlsApp = Nothing
        Catch ex As Exception

        End Try
    End Sub
    Public Sub exc_rprt_save(filename As String)
        Try
            rprt_WS.SaveAs(filename)
        Catch ex As Exception

        End Try
    End Sub
End Module
