Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlBorderWeight
Imports System.Net.Mail
Public Class Form1

    Private colNumber As Integer = 0

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
        Dim MaxWidth As Integer
        Dim ShowHiddenCols As Boolean
    End Structure

    Function StartExcel(Optional ByVal IsVisible As Boolean = False, Optional NumOfSheets As Integer = 1) As Object 'Excel.Application

        Dim objExcel As New Excel.Application
        objExcel.Visible = IsVisible
        objExcel.Application.SheetsInNewWorkbook = 1

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
        Dim r As New Random
        Dim xlws As Excel.Worksheet
        Dim cn As New SqlClient.SqlConnection
        cn.ConnectionString = "Data Source=212.42.46.12;Initial Catalog=CBR_Emm_test;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY"
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandText = "select top 5 *, 100000 from Stock_Dr"
        cmd.Connection = cn
        Dim adap As New SqlClient.SqlDataAdapter
        adap.SelectCommand = cmd
        Dim dt As New DataTable
        adap.Fill(dt)
        Dim s As New Settings
        s.FirstCol = 3
        s.FirstRow = 3
        s.NameOfWorkSheet = "List" + r.Next(1000).ToString
        s.NeedBorder = True
        's.NeedFix = 3
        s.ShowHiddenCols = True
        s.NeedNamesOfCols = True
        s.ColsForAutoFilter = "C4:F4"
        s.ColsNamesForSumOrAvgCalc = New String() {"ISSUESIZE"}
        s.NeedNumeration = True
        s.FontOfNames = "Verdana"
        s.NeedAutoWidth = True
        s.NameOfTable = "Hello"
        DataGridView1.DataSource = dt
        Call DGVToExcel(DataGridView1, xlws, s)
    End Sub

    Private Sub DataTableToExcelSheet(ByVal dt As DataTable, xlws As Excel.Worksheet, ByRef stgs As Settings)

        If stgs.FontOfCells Is Nothing Then
            stgs.FontOfCells = "Times New Roman"
        End If
        If stgs.FontOfNames Is Nothing Then
            stgs.FontOfNames = "Arial Narrow"
        End If

        ' Костыль для подсчета используемых строк
        Dim PlaceForSumAndAvg As Integer
        If stgs.ColsNamesForSumOrAvgCalc Is Nothing Then
            PlaceForSumAndAvg = 0
        Else
            PlaceForSumAndAvg = 2
        End If

        If stgs.FirstRow = 0 Then stgs.FirstRow = 1
        If stgs.FirstCol = 0 Then stgs.FirstCol = 1
        Dim nRow As Integer, nCol As Integer, fRow As Integer = stgs.FirstRow, fCol As Integer = stgs.FirstCol

        If stgs.NameOfWorkSheet Is Nothing Then
            xlws.Name = "List" + xlws.Application.Workbooks(1).Worksheets.Count.ToString
        Else
            xlws.Name = stgs.NameOfWorkSheet
        End If


        ' Объединяем ячейки и вставляем туда название таблицы
        xlws.Cells(fRow, fCol).Font.Bold = True
        xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count) + fRow.ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        If stgs.NameOfTable Is Nothing Then
            xlws.Cells(fRow, fCol) = "TableName"
        Else
            xlws.Cells(fRow, fCol) = stgs.NameOfTable
        End If
        stgs.FirstRow += 1

        ' Проставляем номера строк
        If stgs.NeedNumeration Then
            For nRow = 0 To dt.Rows.Count + PlaceForSumAndAvg - 1
                xlws.Cells(stgs.FirstRow + nRow + 1, stgs.FirstCol) = nRow + 1
            Next
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count + stgs.FirstCol) + fRow.ToString).MergeCells = True
            stgs.FirstCol += 1
        Else
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + stgs.FirstCol + dt.Columns.Count - 1) + fRow.ToString).MergeCells = True
        End If

        If stgs.NeedNamesOfCols Then
            If stgs.NeedNumeration Then
                xlws.Cells(stgs.FirstRow, stgs.FirstCol - 1) = "Number"
            End If
            For nCol = 0 To dt.Columns.Count - 1
                xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol) = dt.Columns(nCol).Caption
                xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol).Font.Name = stgs.FontOfNames
                xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            Next nCol
            stgs.FirstRow += 1
        End If

        ' Заполнение ячеек
        'For nRow = 0 To dt.Rows.Count - 1
        '    For nCol = 0 To dt.Columns.Count - 1
        '        xlws.Cells(stgs.FirstRow + nRow, stgs.FirstCol + nCol) = dt.Rows(nRow).Item(nCol)
        '        xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol).Font.Name = stgs.FontOfCells
        '    Next nCol
        'Next nRow
        nRow = dt.Rows.Count - 1

        For i = 0 To dt.Columns.Count - 1

            colNumber = i

            'Set the content from datatable (which Is converted as array And again converted as string) 
            Clipboard.SetText(AryToString(ToArray(dt)))

            'Identifiy And select the range of cells in Excel to paste the clipboard data. 
            xlws.Cells(stgs.FirstRow, i + stgs.FirstCol).Select()

            'Paste the clipboard data 
            xlws.Paste()
            Clipboard.Clear()
        Next
        xlws.Range("" + ChrW(64 + stgs.FirstCol) + (stgs.FirstRow).ToString + "", "" +
                           ChrW(64 + stgs.FirstCol) + (stgs.FirstRow + dt.Rows.Count - 1).ToString).Font.Name = stgs.FontOfCells

        ' Форматирование
        For idx As Integer = 0 To dt.Columns.Count - 1
            Dim intTest As Integer, dblTest As Double, dateTest As Date
            If Integer.TryParse(xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx).Value, intTest) AndAlso
                intTest = Double.Parse(xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx).Value) Then
                With xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                           ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dt.Rows.Count - 1).ToString)
                    .NumberFormat = "# ##0"
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                End With
            ElseIf Double.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, dblTest) Then
                With xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                           ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dt.Rows.Count - 1).ToString)
                    .NumberFormat = "# ##0.00"
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                End With
            ElseIf Date.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, dateTest) Then
                xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                            ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dt.Rows.Count - 1).ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                For i As Integer = 1 To dt.Rows.Count
                    If Date.TryParse(xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value, dateTest) Then
                        xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value = dateTest.ToShortDateString
                    End If
                Next
            Else
                xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                            ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dt.Rows.Count - 1).ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            End If
        Next

        ' Заменяем пустые ячейки
        xlws.Range("" + ChrW(64 + fCol) + (fRow + 1).ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) +
                 (stgs.FirstRow + dt.Rows.Count - 1).ToString + "").SpecialCells(Excel.XlCellType.xlCellTypeBlanks).Select()
        'xl.Selection.Value = "-"
        xlws.Application.Selection.Value = "-"

        ' Делаем автоширину, но не более 20 
        If stgs.NeedAutoWidth Then
            xlws.Cells.EntireColumn.AutoFit()
            If stgs.MaxWidth <> 0 Then
                For nCol = 0 To dt.Columns.Count - 1
                    If xlws.Columns(stgs.FirstCol + nCol).ColumnWidth > stgs.MaxWidth Then
                        xlws.Columns(stgs.FirstCol + nCol).ColumnWidth = stgs.MaxWidth
                    End If
                Next
            End If
        End If

        If stgs.NeedBorder Then
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) +
                 (stgs.FirstRow + dt.Rows.Count - 1 + PlaceForSumAndAvg).ToString + "").Select()
            With xlws.Application.Selection().Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xlws.Application.Selection().Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xlws.Application.Selection().Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xlws.Application.Selection().Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If

        If stgs.NeedGrid Then
            With xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) +
                            (stgs.FirstRow + dt.Rows.Count - 1 + PlaceForSumAndAvg).ToString + "").Borders
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
            xlws.Application.Application.Selection.AutoFilter()
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

        stgs.FirstRow += dt.Rows.Count + 2 + PlaceForSumAndAvg
    End Sub

    'Private Sub ExcelClose(xl As Excel.Application, dir As String, pas As String)
    '    xl.Workbooks(0).SaveAs(dir, FileFormat:="xlxs", Password:=pas)
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim xlws As New Excel.Worksheet
        Dim r As New Random
        Dim cn As New SqlClient.SqlConnection
        cn.ConnectionString = "Data Source=212.42.46.12;Initial Catalog=CBR_Emm_test;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY"
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandText = "select top 5 *, 100000 from Stock_Dr"
        cmd.Connection = cn
        Dim adap As New SqlClient.SqlDataAdapter
        adap.SelectCommand = cmd
        Dim dt As New DataTable
        adap.Fill(dt)
        Dim s As New Settings
        s.FirstCol = 3
        s.FirstRow = 3
        s.NameOfWorkSheet = "List" + r.Next(1000).ToString
        s.NeedBorder = True
        's.NeedFix = 3
        s.NeedNamesOfCols = True
        s.ColsForAutoFilter = "C4:F4"
        s.ColsNamesForSumOrAvgCalc = New String() {"ISSUESIZE"}
        s.NeedNumeration = True
        s.FontOfNames = "Verdana"
        s.NeedAutoWidth = True
        s.NameOfTable = "Hello"
        Call DataTableToExcelSheet(dt, xlws, s)
    End Sub

    Private Sub DGVToExcel(ByVal dgv As DataGridView, xlws As Excel.Worksheet, ByRef stgs As Settings, Optional CopyDGV As Boolean = False)

        If CopyDGV Then
            MarofetDataGreed(dgv, dgv.DataSource)
            'Data transfer from grid to Excel.  
            With xlws
                .Range("A1", Type.Missing).EntireRow.Font.Bold = True
                'Set Clipboard Copy Mode     
                DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
                DataGridView1.SelectAll()

                'Get the content from Grid for Clipboard     
                Dim str As String = TryCast(DataGridView1.GetClipboardContent().GetData(DataFormats.UnicodeText), String)

                'Set the content to Clipboard     
                Clipboard.SetText(str, TextDataFormat.UnicodeText)

                'Identify and select the range of cells in Excel to paste the clipboard data.     
                .Cells(1, 1).Select()

                'Paste the clipboard data     
                .Paste()
                Clipboard.Clear()
            End With
            Return
        End If

        Dim hiddenCols As Integer = 0
        For i As Integer = 0 To dgv.Columns.Count - 1
            If Not dgv.Columns(i).Visible Then
                hiddenCols += 1
            End If
        Next

        If stgs.FontOfCells Is Nothing Then
            stgs.FontOfCells = "Times New Roman"
        End If
        If stgs.FontOfNames Is Nothing Then
            stgs.FontOfNames = "Arial Narrow"
        End If

        ' Костыль для подсчета используемых строк
        Dim PlaceForSumAndAvg As Integer
        If stgs.ColsNamesForSumOrAvgCalc Is Nothing Then
            PlaceForSumAndAvg = 0
        Else
            PlaceForSumAndAvg = 2
        End If

        If stgs.FirstRow = 0 Then stgs.FirstRow = 1
        If stgs.FirstCol = 0 Then stgs.FirstCol = 1
        Dim nRow As Integer, nCol As Integer, fRow As Integer = stgs.FirstRow, fCol As Integer = stgs.FirstCol

        If stgs.NameOfWorkSheet Is Nothing Then
            xlws.Name = "List" + xlws.Application.Workbooks(1).Worksheets.Count.ToString
        Else
            xlws.Name = stgs.NameOfWorkSheet
        End If

        ' Объединяем ячейки и вставляем туда название таблицы
        xlws.Cells(fRow, fCol).Font.Bold = True
        xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count) + fRow.ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        If stgs.NameOfTable Is Nothing Then
            xlws.Cells(fRow, fCol) = "TableName"
        Else
            xlws.Cells(fRow, fCol) = stgs.NameOfTable
        End If
        stgs.FirstRow += 1

        ' Проставляем номера строк
        If stgs.NeedNumeration Then
            For nRow = 1 To dgv.Rows.Count + PlaceForSumAndAvg - 1
                xlws.Cells(stgs.FirstRow + nRow, stgs.FirstCol) = nRow
            Next

            If stgs.ShowHiddenCols Then
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count + stgs.FirstCol) + fRow.ToString).MergeCells = True
            Else
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count + stgs.FirstCol - hiddenCols) + fRow.ToString).MergeCells = True
            End If
            stgs.FirstCol += 1
        Else
            If stgs.ShowHiddenCols Then
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + stgs.FirstCol + dgv.Columns.Count - 1) + fRow.ToString).MergeCells = True
            Else
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + stgs.FirstCol + dgv.Columns.Count - 1 - hiddenCols) + fRow.ToString).MergeCells = True
            End If

        End If

        ' Проставляем имена столбцов
        If stgs.NeedNamesOfCols Then
            If stgs.NeedNumeration Then
                xlws.Cells(stgs.FirstRow, stgs.FirstCol - 1) = "Number"
            End If
            Dim idx As Integer = 0
            For nCol = 0 To dgv.Columns.Count - 1
                If dgv.Columns(nCol).Visible Or stgs.ShowHiddenCols Then
                    xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx) = dgv.Columns(nCol).HeaderText
                    xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx).Font.Name = stgs.FontOfNames
                    xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    idx += 1
                End If
            Next nCol
            stgs.FirstRow += 1
        End If

        nRow = dgv.Rows.Count - 1

        ' Заполнение ячеек
        If stgs.ShowHiddenCols Then
            'For nRow = 0 To dgv.Rows.Count - 1
            '    For nCol = 0 To dgv.Columns.Count - 1
            '        xlws.Cells(stgs.FirstRow + nRow, stgs.FirstCol + nCol) = dgv.Rows(nRow).Cells(nCol).Value
            '    Next nCol
            'Next nRow
            For i = 0 To dgv.Columns.Count - 1

                colNumber = i

                'Set the content from datatable (which Is converted as array And again converted as string) 
                Clipboard.SetText(AryToString(ToArray(dgv.DataSource)))

                'Identifiy And select the range of cells in Excel to paste the clipboard data. 
                xlws.Cells(stgs.FirstRow, i + stgs.FirstCol).Select()

                'Paste the clipboard data 
                xlws.Paste()
                Clipboard.Clear()
            Next
        Else
            'Data transfer from grid to Excel.  
            With xlws
                'Set Clipboard Copy Mode     
                DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
                DataGridView1.SelectAll()

                'Get the content from Grid for Clipboard     
                Dim str As String = TryCast(DataGridView1.GetClipboardContent().GetData(DataFormats.UnicodeText), String)

                'Set the content to Clipboard     
                Clipboard.SetText(str, TextDataFormat.UnicodeText)

                'Identify And select the range of cells in Excel to paste the clipboard data.     
                .Cells(stgs.FirstRow, stgs.FirstCol).Select()

                'Paste the clipboard data     
                .Paste()
                Clipboard.Clear()
            End With
        End If

        xlws.Range("" + ChrW(64 + stgs.FirstCol) + (stgs.FirstRow).ToString + "", "" +
                           ChrW(64 + stgs.FirstCol) + (stgs.FirstRow + dgv.Rows.Count - 1).ToString).Font.Name = stgs.FontOfCells

        ' Форматирование
        For idx As Integer = 0 To dgv.Columns.Count - 1
            Dim intTest As Integer, dblTest As Double, dateTest As Date, boolTest As Boolean
            If Integer.TryParse(xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx).Value, intTest) AndAlso
                intTest = Double.Parse(xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx).Value) Then
                With xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                           ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dgv.Rows.Count - 1).ToString)
                    .NumberFormat = "# ##0"
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                End With
            ElseIf Double.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, dblTest) Then
                With xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                           ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dgv.Rows.Count - 1).ToString)
                    .NumberFormat = "# ##0.00"
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                End With
            ElseIf Date.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, dateTest) Then
                xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                            ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dgv.Rows.Count - 1).ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                For i As Integer = 1 To dgv.Rows.Count
                    If Date.TryParse(xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value, dateTest) Then
                        xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value = dateTest.ToShortDateString
                    End If
                Next
                '    With xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                '            ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dgv.Rows.Count - 1).ToString)
                '    .NumberFormat = "DD/MM/YYYY"
                '    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                'End With
            ElseIf Boolean.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, boolTest) Then
                For i As Integer = 1 To dgv.Rows.Count
                    If Boolean.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, boolTest) Then
                        xlws.Cells(stgs.FirstRow + i, stgs.FirstCol + idx).Value = boolTest
                    End If
                Next
            Else
                xlws.Range("" + ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow).ToString + "", "" +
                            ChrW(64 + stgs.FirstCol + idx) + (stgs.FirstRow + dgv.Rows.Count - 1).ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            End If
        Next

        ' Заменяем пустые ячейки
        xlws.Range("" + ChrW(64 + fCol) + (fRow + 1).ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol - hiddenCols) +
                 (stgs.FirstRow + dgv.Rows.Count - 2).ToString + "").SpecialCells(Excel.XlCellType.xlCellTypeBlanks).Select()
        xlws.Application.Selection.Value = "-"

        ' Делаем автоширину, но не более 20 
        If stgs.NeedAutoWidth Then
            xlws.Cells.EntireColumn.AutoFit()
            If stgs.MaxWidth <> 0 Then
                For nCol = 0 To dgv.Columns.Count - 1
                    If xlws.Columns(stgs.FirstCol + nCol).ColumnWidth > stgs.MaxWidth Then
                        xlws.Columns(stgs.FirstCol + nCol).ColumnWidth = stgs.MaxWidth
                    End If
                Next
            End If
        End If

        If stgs.NeedBorder Then
            If stgs.ShowHiddenCols Then
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol) +
                 (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Select()
            Else
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol - hiddenCols) +
                 (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Select()
            End If
            With xlws.Application.Selection().Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xlws.Application.Selection().Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xlws.Application.Selection().Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With xlws.Application.Selection().Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If

        If stgs.NeedGrid Then
            If stgs.ShowHiddenCols Then
                With xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol) +
                            (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            Else
                With xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol - hiddenCols) +
                            (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            End If
        End If

        If Not stgs.ColsForAutoFilter Is Nothing Then
            If stgs.ColsForAutoFilter = "ALL" Then
                xlws.Rows(fRow + 1).Select()
            Else
                xlws.Range(stgs.ColsForAutoFilter).Select()
            End If
            xlws.Application.Application.Selection.AutoFilter()
        End If

        If Not stgs.ColsNamesForSumOrAvgCalc Is Nothing Then
            For nCol = 0 To dgv.Columns.Count - 1
                If stgs.ColsNamesForSumOrAvgCalc.Contains(xlws.Cells(fRow + 1, stgs.FirstCol + nCol).Value.ToString) Then
                    xlws.Cells(stgs.FirstRow + dgv.Rows.Count - 1, stgs.FirstCol + nCol).Value = "=SUM(" +
                        ChrW(64 + nCol + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + nCol + stgs.FirstCol) +
                        (stgs.FirstRow + dgv.Rows.Count - 1).ToString + ")"
                    xlws.Cells(stgs.FirstRow + dgv.Rows.Count, stgs.FirstCol + nCol).Value = "=Average(" +
                        ChrW(64 + nCol + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + nCol + stgs.FirstCol) +
                        (stgs.FirstRow + dgv.Rows.Count - 1).ToString + ")"
                End If
            Next
            xlws.Rows(stgs.FirstRow + dgv.Rows.Count).Font.Bold = True
        End If

        If Not stgs.NeedFix = Nothing Then
            xlws.Application.ActiveWindow.SplitRow = stgs.NeedFix
            xlws.Application.ActiveWindow.FreezePanes = True
        End If

        stgs.FirstRow += dgv.Rows.Count + 2 + PlaceForSumAndAvg
    End Sub

    Public Function ToArray(ByVal dr As DataTable) As String()
        Dim ary() As String = Array.ConvertAll(Of DataRow, String)(dr.Select(), AddressOf DataRowToString)
        Return ary
    End Function

    Public Function DataRowToString(ByVal dr As System.Data.DataRow) As String
        Return dr(colNumber).ToString
    End Function

    'Method convert Array to string 
    Public Function AryToString(ByVal ary As String()) As String
        Return String.Join(vbNewLine, ary)
    End Function

    Public Sub MarofetDataGreed(dgv As DataGridView, dttbl As DataTable)

        Dim s As String, TypeColum As String

        For i As Integer = 0 To dttbl.Columns.Count - 1
            TypeColum = dttbl.Columns(i).DataType.Name
            s = Replace(dttbl.Columns(i).ColumnName, ", ", " ")

            If TypeColum = "Double" Then

                With dgv.Columns(s).DefaultCellStyle

                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.BottomRight

                End With

            ElseIf TypeColum = "Int32" Then

                With dgv.Columns(s).DefaultCellStyle

                    .Format = "N0"
                    .Alignment = DataGridViewContentAlignment.BottomRight

                End With

            ElseIf TypeColum = "DateTime" Then

                dgv.Columns(s).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            ElseIf TypeColum = "String" Then

                dgv.Columns(s).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

            End If

        Next

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cn As New SqlClient.SqlConnection
        Dim xlws As Excel.Worksheet
        cn.ConnectionString = "Data Source=212.42.46.12;Initial Catalog=CBR_Emm_test;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY"
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandText = "select top 5 *, 100000 from Stock_Dr"
        cmd.Connection = cn
        Dim adap As New SqlClient.SqlDataAdapter
        adap.SelectCommand = cmd
        Dim dt As New DataTable
        adap.Fill(dt)
        Dim s As New Settings
        s.FirstCol = 3
        s.FirstRow = 3
        s.NameOfWorkSheet = "1hihjmh"
        s.NeedBorder = True
        's.NeedFix = 3
        s.NeedNamesOfCols = True
        s.ColsForAutoFilter = "C4:F4"
        s.ColsNamesForSumOrAvgCalc = New String() {"ISSUESIZE"}
        's.NeedNumeration = True
        s.FontOfNames = "Verdana"
        s.NeedAutoWidth = True
        s.NameOfTable = "Hello"
        DataGridView1.DataSource = dt
        Call DGVToExcel(DataGridView1, xlws, s, True)
    End Sub

    'Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
    '    xl = StartExcel()
    '    xlwb = xl.Workbooks.Add()
    'End Sub

    'Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
    '    ExcelClose()
    'End Sub

    Private Sub ExcelClose(xlwb As Excel.Workbook, Optional ByVal SavePath As Boolean = False,
                           Optional ByVal password As String = Nothing, Optional ByVal SendMail As Boolean = False,
                           Optional ByVal Email As String = "empty_box44@mail.ru")

        If Not password Is Nothing Then
            xlwb.Password = password
        End If

        Dim saveFileDialog1 As New SaveFileDialog

        If SavePath = True Then
            saveFileDialog1.Filter = "(*.xlsx) | *.xlsx"
            saveFileDialog1.Title = "Save an Excel File"
            saveFileDialog1.ShowDialog()
            If saveFileDialog1.FileName <> "" Then
                xlwb.SaveAs(saveFileDialog1.FileName)
            End If
        End If
        xlwb.Close(False, Type.Missing, Type.Missing)
        xlwb.Application.Quit()

        If SendMail Then
            SendEmail(saveFileDialog1.FileName, Email)
        End If

    End Sub

    ''Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    ''    FixesForPrint()
    ''End Sub

    Private Sub FixesForPrint(xlws As Excel.Worksheet, Optional ByVal NeedNumeration As Boolean = False, Optional ByVal LandscapeOrientation As Boolean = False,
                           Optional ByVal PrintArea As String = "A1:A2")
        If NeedNumeration Then
            'For Each xlws In xlwb.Sheets
            'xlwb.ActiveSheet.PageSetup.CenterFooter = "&P/&N"     ' Нумерация листов
            xlws.PageSetup.CenterFooter = "&P/&N"
            'Next
        End If

        'For Each xlws In xlwb.Sheets
        If LandscapeOrientation Then
            xlws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape ' Вид печати
        Else
            xlws.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait
        End If
        'Next

        'For Each xlws In xlwb.Sheets
        xlws.PageSetup.PrintArea = PrintArea     ' Зона печати
        'Next

    End Sub

    Private Sub SendEmail(ByVal FilePath As String, Optional ByVal mail As String = Nothing)

        Dim email As New MailMessage
        Dim smtp As SmtpClient
        smtp = New SmtpClient("smtp.mail.ru") ' Пример smtp.mail.ru
        'System.Net.ServicePointManager.CertificatePolicy = New MyPolicy

        smtp.Credentials = New Net.NetworkCredential("empty_box41@mail.ru", "123qwerTY")
        smtp.EnableSsl = True
        smtp.Port = 587

        email.From = New MailAddress("empty_box41@mail.ru")
        email.To.Add(New MailAddress(mail))

        email.Subject = "Excel file"
        email.Body = "mail with attachment"

        Dim attachment As System.Net.Mail.Attachment
        attachment = New System.Net.Mail.Attachment(FilePath)
        email.Attachments.Add(attachment)
        smtp.Send(email)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim xl As Excel.Application = StartExcel(True)
        Dim xlwb As Excel.Workbook = xl.Workbooks.Add()
        Dim xlws As Excel.Worksheet = xlwb.Worksheets.Add()
        xlwb.Sheets("Лист1").Delete()

        Dim cn As New SqlClient.SqlConnection("Data Source=212.42.46.12;Initial Catalog=CBR_Emm_test;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY")
        Dim cmd As New SqlClient.SqlCommand(Nothing, cn)
        cmd.CommandText = "select top 5 *, 100000 from Stock_Dr"
        Dim adap As New SqlClient.SqlDataAdapter(cmd)
        Dim dt As New DataTable
        adap.Fill(dt)



        Dim s As New Settings
        s.NameOfWorkSheet = "ListName1"
        s.NeedBorder = True
        's.NeedFix = 3
        s.NeedNamesOfCols = True
        s.ColsForAutoFilter = "C2:F2"
        s.ColsNamesForSumOrAvgCalc = New String() {"ISSUESIZE"}
        's.NeedNumeration = True
        s.NeedAutoWidth = True

        Call DataTableToExcelSheet(dt, xlws, s)

        ' Чтобы таблица отразилась на той же странице, нужно не создавать новые объекты Settings и Worksheet
        cmd.CommandText = "select top 10 * from Stock_Dr"
        dt = New DataTable
        adap.Fill(dt)

        'xlws = xlwb.Worksheets.Add()
        's = New Settings
        's.FirstCol = 5
        's.FirstRow = 4
        's.NameOfWorkSheet = "ListNameName"
        's.NeedGrid = True
        's.MaxWidth = False
        ''s.NeedFix = 3
        's.NeedNamesOfCols = True
        's.ColsForAutoFilter = "F4:G4"
        ''s.ColsNamesForSumOrAvgCalc = New String() {"ISSUESIZE"}
        's.NeedNumeration = True
        's.FontOfNames = "Times New Roman"
        's.NeedAutoWidth = True
        's.NameOfTable = "HelloToo"
        DataGridView1.DataSource = dt
        Call DGVToExcel(DataGridView1, xlws, s)


        cmd.CommandText = "select top 5 * from Stock_Dr"
        dt = New DataTable
        adap.Fill(dt)

        xlws = xlwb.Worksheets.Add()
        s = New Settings
        s.FirstCol = 1
        s.FirstRow = 1
        s.NameOfWorkSheet = "ListNameNameToo"
        's.NeedGrid = True
        's.NeedFix = 3
        s.NeedNamesOfCols = True
        's.ColsForAutoFilter = "F4:G4"
        's.ColsNamesForSumOrAvgCalc = New String() {"ISSUESIZE"}
        ''s.NeedNumeration = True
        s.FontOfNames = "Times New Roman"
        s.NeedAutoWidth = True
        s.NameOfTable = "HelloTooToo"
        DataGridView1.DataSource = dt
        Call DGVToExcel(DataGridView1, xlws, s, True)

        Call FixesForPrint(xlws)

        Call ExcelClose(xlwb, True, "123")
    End Sub
End Class
