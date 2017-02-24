Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlBorderWeight
Imports Microsoft.Office.Interop.Excel.XlDirection
Public Class Form1

    Structure Settings
        Dim FirstCol As Integer
        Dim FirstRow As Integer
        Dim NeedBorder As Boolean
        Dim NeedGrid As Boolean
        Dim NeedNamesOfCols As Boolean
        Dim NameOfTable As String
        Dim FontOfNames As String
        Dim FontOfCells As String
        Dim FormatOfCols As String
        Dim ColsNamesForSumOrAvgCalc() As String
        Dim ColsForAutoFilter As String
        Dim NameOfWorkSheet As String
        Dim NeedNumeration As Boolean
        Dim NeedFix As Integer
        Dim NeedAutoWidth As Boolean
        Dim ShowHiddenCols As Boolean
    End Structure

    Function StartExcel(Optional ByVal IsVisible As Boolean = True, Optional NumOfSheets As Integer = 1) As Object 'Excel.Application

        Dim objExcel As New Excel.Application
        objExcel.Visible = IsVisible
        objExcel.Application.SheetsInNewWorkbook = NumOfSheets
        Return objExcel

    End Function

    Sub ForceExcelToQuit(ByVal objExcel As Object) 'Excel.Application)

        Try
            objExcel.Quit()
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim stgs As New Settings
        ' ...

    End Sub

    Private Sub DataTableToExcelSheet(ByVal dt As DataTable, ByVal stgs As Settings)
        Dim xl As Excel.Application = StartExcel(True, 1)
        Dim xlws As Excel.Worksheet
        Dim xlwb As Excel.Workbook
        xlwb = xl.Workbooks.Add()
        xlws = xlwb.Sheets(1)

        ' Костыль для подсчета используемых строк
        Dim PlaceForSumAndAvg As Integer
        If stgs.ColsNamesForSumOrAvgCalc Is Nothing Then
            PlaceForSumAndAvg = 0
        Else
            PlaceForSumAndAvg = 2
        End If

        Dim nRow As Integer, nCol As Integer
        Dim fRow As Integer = stgs.FirstRow, fCol As Integer = stgs.FirstCol

        xlws.Name = stgs.NameOfWorkSheet

        ' Объединяем ячейки и вставляем туда название таблицы
        xlws.Cells(fRow, fCol).Font.Bold = True
        xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count) + fRow.ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        xlws.Cells(fRow, fCol) = stgs.NameOfTable
        stgs.FirstRow += 1

        ' Проставляем номера строк
        If stgs.NeedNumeration Then
            For nRow = 0 To dt.Rows.Count + PlaceForSumAndAvg
                xlws.Cells(stgs.FirstRow + nRow, stgs.FirstCol) = nRow + 1
            Next
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count + stgs.FirstCol) + fRow.ToString).MergeCells = True
            stgs.FirstCol += 1
        Else
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + stgs.FirstCol + dt.Columns.Count - 1) + fRow.ToString).MergeCells = True
        End If

        If stgs.NeedNamesOfCols Then
            For nCol = 0 To dt.Columns.Count - 1
                xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol) = dt.Columns(nCol).Caption
                xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol).Font.Name = stgs.FontOfNames
                xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next nCol
            stgs.FirstRow += 1
        End If

        ' Заполнение ячеек
        For nRow = 0 To dt.Rows.Count - 1
            For nCol = 0 To dt.Columns.Count - 1
                xlws.Cells(stgs.FirstRow + nRow, stgs.FirstCol + nCol) = dt.Rows(nRow).Item(nCol)
                xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol).Font.Name = stgs.FontOfCells
            Next nCol
        Next nRow

        ' Делаем автоширину, но не более 20 
        If stgs.NeedAutoWidth Then
            xlws.Cells.EntireColumn.AutoFit()
            For nCol = 0 To dt.Columns.Count - 1
                If xlws.Columns(stgs.FirstCol + nCol).ColumnWidth > 20 Then
                    xlws.Columns(stgs.FirstCol + nCol).ColumnWidth = 20
                End If
            Next
        End If

        If stgs.NeedBorder Then
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) + (fRow + nRow + 1 + PlaceForSumAndAvg).ToString + "").Select()
            With xl.Selection().Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xl.Selection().Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xl.Selection().Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xl.Selection().Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If

        If stgs.NeedGrid Then
            With xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) + (stgs.FirstRow + dt.Rows.Count - 1 + PlaceForSumAndAvg).ToString + "").Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End If

        If Not stgs.ColsForAutoFilter Is Nothing Then
            If stgs.ColsForAutoFilter = "ALL" Then
                xlws.Rows(fRow + 1).Select()
            Else
                xlws.Range(stgs.ColsForAutoFilter).Select()
            End If
            xl.Application.Selection.AutoFilter()
        End If

        If Not stgs.ColsNamesForSumOrAvgCalc Is Nothing Then
            For nCol = 0 To dt.Columns.Count - 1
                If stgs.ColsNamesForSumOrAvgCalc.Contains(xlws.Cells(fRow + 1, stgs.FirstCol + nCol).Value.ToString) Then
                    xlws.Cells(stgs.FirstRow + dt.Rows.Count, stgs.FirstCol + nCol).Value = "=SUM(" +
                        ChrW(64 + nCol + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + nCol + stgs.FirstCol) +
                        (stgs.FirstRow + dt.Rows.Count - 1).ToString + ")"
                    xlws.Cells(stgs.FirstRow + dt.Rows.Count + 1, stgs.FirstCol + nCol).Value = "=Average(" +
                        ChrW(64 + nCol + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + nCol + stgs.FirstCol) +
                        (stgs.FirstRow + dt.Rows.Count - 1).ToString + ")"
                End If
            Next
            xlws.Rows(stgs.FirstRow + dt.Rows.Count).Font.Bold = True
        End If

        If Not stgs.NeedFix = Nothing Then
            xlws.Application.ActiveWindow.SplitRow = stgs.NeedFix
            xlws.Application.ActiveWindow.FreezePanes = True
        End If

        'xlws.PageSetup.CenterFooter = "&P/&N"     ' Нумерация листов
        xlws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape    ' Вид печати
        xlws.PageSetup.PrintArea = "C3:M9"      ' Зона печати

    End Sub

    Private Sub ExcelClose(xl As Excel.Application, dir As String, pas As String)
        xl.Workbooks(0).SaveAs(dir, FileFormat:="xlxs", Password:=pas)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cn As New SqlClient.SqlConnection
        cn.ConnectionString = "Data Source=212.42.46.12;Initial Catalog=CBR_Emm_test;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY"
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandText = "select top 5 * from Stock_Dr"
        cmd.Connection = cn
        Dim adap As New SqlClient.SqlDataAdapter
        adap.SelectCommand = cmd
        Dim dt As New DataTable
        adap.Fill(dt)
        Dim s As New Settings
        s.FirstCol = 3
        s.FirstRow = 3
        s.NameOfTable = "Hello"
        s.NameOfWorkSheet = "1hihjmh"
        s.NeedBorder = True
        's.NeedFix = 3
        s.NeedNamesOfCols = True
        s.ColsForAutoFilter = "ALL"
        s.ColsNamesForSumOrAvgCalc = New String() {"ISSUESIZE"}
        's.NeedNumeration = True
        s.FontOfNames = "Verdana"
        s.NeedAutoWidth = True
        Call DataTableToExcelSheet(dt, s)
    End Sub

End Class
