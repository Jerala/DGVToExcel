Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlBorderWeight
Imports System.Net.Mail
Imports System.IO
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
        Dim ColsNamesForSumCalc() As String
        Dim ColsNamesForAvgCalc() As String
        Dim ColsForAutoFilter As String
        Dim NameOfWorkSheet As String
        Dim NeedNumeration As Boolean
        Dim CellToFix As String
        Dim NeedAutoWidth As Boolean
        Dim MaxWidth As Integer
        Dim ShowHiddenCols As Boolean
        Dim replaceBoolTrueVal As String
        Dim replaceBoolFalseVal As String
        Dim RangeOfCellsForNamesOfColsFromAnotherFile As String
    End Structure

    Function StartExcel(Optional ByVal IsVisible As Boolean = False, Optional NumOfSheets As Integer = 1) As Object 'Excel.Application

        Dim objExcel As New Excel.Application
        objExcel.Visible = IsVisible
        objExcel.Application.SheetsInNewWorkbook = 1

        Return objExcel

    End Function

    'Sub ForceExcelToQuit(ByVal objExcel As Object) 'Excel.Application)

    '    Try
    '        objExcel.Quit()
    '    Catch ex As Exception
    '        MsgBox(ex)
    '    End Try

    'End Sub

    Private Sub TakeNamesFromAnotherFile(xlws As Excel.Worksheet, ByVal stgs As Settings)
        Try
            Dim ofd As New OpenFileDialog
            ofd.Filter = "(*.xlsx) | *.xlsx"

            If ofd.ShowDialog() = DialogResult.OK Then

                Dim str As String() = stgs.RangeOfCellsForNamesOfColsFromAnotherFile.Split(New Char() {":", ";", ","})
                Dim exc As New Excel.Application
                Dim xlwb As Excel.Workbook = exc.Workbooks.Open(ofd.FileName)
                Dim xlws1 As Excel.Worksheet = xlwb.Sheets(1)
                exc.Visible = True

                For i As Integer = 0 To (Convert.ToInt32(str(1)(0)) - Convert.ToInt32(str(0)(0)))
                    Clipboard.SetText(xlws1.Cells(Val(str(0)(1)), Convert.ToInt32(str(0)(0)) - 64 + i).Value.ToString)

                    xlws.Cells(stgs.FirstRow, stgs.FirstCol + i).Select
                    xlws.Paste()
                    Clipboard.Clear()
                Next

                xlwb.Close(False)
                exc.Quit()
                exc = Nothing
                xlwb = Nothing
                xlws1 = Nothing
                GC.Collect()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SetDefaultValues(ByRef stgs As Settings)
        If stgs.replaceBoolTrueVal Is Nothing Then stgs.replaceBoolTrueVal = "Да"
        If stgs.replaceBoolFalseVal Is Nothing Then stgs.replaceBoolFalseVal = "Нет"
        If stgs.FontOfCells Is Nothing Then
            stgs.FontOfCells = "Times New Roman"
        End If
        If stgs.FontOfNames Is Nothing Then
            stgs.FontOfNames = "Arial Narrow"
        End If

        If stgs.FirstRow = 0 Then stgs.FirstRow = 1
        If stgs.FirstCol = 0 Then stgs.FirstCol = 1
    End Sub

    Private Sub DataTableToExcelSheet(ByVal dt As DataTable, xlws As Excel.Worksheet, ByRef stgs As Settings)

        Call SetDefaultValues(stgs)

        xlws.Cells.Font.Name = "Times New Roman"

        ' Костыль для подсчета используемых строк
        Dim PlaceForSumAndAvg As Integer = 0
        If Not stgs.ColsNamesForSumCalc Is Nothing Then PlaceForSumAndAvg += 1
        If Not stgs.ColsNamesForAvgCalc Is Nothing Then PlaceForSumAndAvg += 1

        Dim nRow As Integer, nCol As Integer, fRow As Integer = stgs.FirstRow, fCol As Integer = stgs.FirstCol

        If stgs.NameOfWorkSheet Is Nothing Then
            xlws.Name = "List" + xlws.Application.Workbooks(1).Worksheets.Count.ToString
        Else
            xlws.Name = stgs.NameOfWorkSheet
        End If


        ' Вставляем название таблицы
        If Not stgs.NameOfTable Is Nothing Then
            xlws.Cells(fRow, fCol).Font.Bold = True
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count) + fRow.ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            xlws.Cells(fRow, fCol) = stgs.NameOfTable
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dt.Columns.Count + stgs.FirstCol) + fRow.ToString).MergeCells = True
            stgs.FirstRow += 1
        End If

        xlws.Range(ChrW(64 + stgs.FirstCol) + (stgs.FirstRow + 1).ToString, ChrW(64 + stgs.FirstCol) + (stgs.FirstRow + dt.Rows.Count).ToString).Font.Name = stgs.FontOfCells
        ' Проставляем номера строк 
        If stgs.NeedNumeration Then
            For nRow = 0 To dt.Rows.Count - 1
                xlws.Cells(stgs.FirstRow + nRow + 1, stgs.FirstCol) = nRow + 1
            Next
            stgs.FirstCol += 1
        End If

        ' Именования суммы и среднего значения
        If Not stgs.ColsNamesForSumCalc Is Nothing Then
            xlws.Cells(stgs.FirstRow + dt.Rows.Count + 1, fCol) = "Сумма"
            xlws.Rows(stgs.FirstRow + dt.Rows.Count + 1).Font.Bold = True
        End If
        If Not stgs.ColsNamesForAvgCalc Is Nothing Then
            Dim NeedOneMoreCell As Integer = 0
            If Not stgs.ColsNamesForSumCalc Is Nothing Then NeedOneMoreCell = 1
            xlws.Cells(stgs.FirstRow + dt.Rows.Count + 1 + NeedOneMoreCell, fCol) = "Среднее"
            xlws.Rows(stgs.FirstRow + dt.Rows.Count + 1 + NeedOneMoreCell).Font.Bold = True
        End If

        ' Проставляем имена столбцов
        If stgs.NeedNamesOfCols Then

            xlws.Rows(stgs.FirstRow).Font.Name = stgs.FontOfNames
            xlws.Rows(stgs.FirstRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            If stgs.NeedNumeration Then
                xlws.Cells(stgs.FirstRow, stgs.FirstCol - 1) = "Number"
            End If

            If Not stgs.RangeOfCellsForNamesOfColsFromAnotherFile Is Nothing Then
                Call TakeNamesFromAnotherFile(xlws, stgs)
            Else
                For nCol = 0 To dt.Columns.Count - 1
                    xlws.Cells(stgs.FirstRow, stgs.FirstCol + nCol) = dt.Columns(nCol).Caption
                Next nCol
            End If
            stgs.FirstRow += 1
        End If

        ' Перенос ячеек
        For i = 0 To dt.Columns.Count - 1

            colNumber = i

            Try
                'Set the content from datatable (which Is converted as array And again converted as string) 
                Clipboard.SetText(AryToString(ToArray(dt)))

                'Identifiy And select the range of cells in Excel to paste the clipboard data. 
                xlws.Cells(stgs.FirstRow, i + stgs.FirstCol).Select()

                'Paste the clipboard data 
                xlws.Paste()
                Clipboard.Clear()
            Catch ex As Exception
                For j = 0 To dt.Rows.Count - 1
                    xlws.Cells(stgs.FirstRow + j, stgs.FirstCol + i).Value = dt.Rows(j)(i)
                Next
            End Try
        Next
        xlws.Range("" + ChrW(64 + stgs.FirstCol) + (stgs.FirstRow).ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) +
                 (stgs.FirstRow + dt.Rows.Count - 1).ToString + "").Font.Name = stgs.FontOfCells

        ' Форматирование
        For idx As Integer = 0 To dt.Columns.Count - 1
            Dim intTest As Integer, dblTest As Double, dateTest As Date, boolTest As Boolean

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
                    .NumberFormat = "# ##0.00" ' 2 знака после запятой с разделением разрядов
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
            ElseIf Boolean.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, boolTest) Then
                For i As Integer = 1 To dt.Rows.Count
                    If Boolean.TryParse(xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value, boolTest) Then
                        If boolTest Then
                            xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value = stgs.replaceBoolTrueVal
                        Else
                            xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value = stgs.replaceBoolFalseVal
                        End If
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
        xlws.Application.Selection.Value = "-"

        ' Делаем автоширину с возможным ограничением
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

        Dim CellForName As Integer = 0
        If Not stgs.NameOfTable Is Nothing Then CellForName = 1

        If stgs.NeedBorder Then
            xlws.Range("" + ChrW(64 + fCol) + (fRow + CellForName).ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) +
                 (stgs.FirstRow + dt.Rows.Count - 1 + PlaceForSumAndAvg).ToString + "").Select()
            With xlws.Application.Selection().Borders '(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If

        If stgs.NeedGrid Then
            With xlws.Range("" + ChrW(64 + fCol) + (fRow + CellForName).ToString + "", "" + ChrW(64 + dt.Columns.Count - 1 + stgs.FirstCol) +
                            (stgs.FirstRow + dt.Rows.Count - 1 + PlaceForSumAndAvg).ToString + "").Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With

            xlws.Range(ChrW(64 + fCol) + (fRow + CellForName).ToString, ChrW(64 + fCol + dt.Columns.Count) + (fRow + CellForName).ToString).Select()

            With xlws.Application.Selection.Borders '(xlEdgeLeft)
                .LineStyle = xlDouble
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
        End If

        If Not stgs.ColsForAutoFilter Is Nothing Then
            If stgs.ColsForAutoFilter = "ALL" Then
                xlws.Rows(fRow + CellForName).Select()
            Else
                Dim cells As String() = stgs.ColsForAutoFilter.Split(New Char() {";", ":", ","})
                xlws.Range(ChrW(65 + cells(0)) + (fRow + CellForName).ToString, ChrW(65 + cells(1)) + (fRow + CellForName).ToString).Select()
            End If
            xlws.Application.Application.Selection.AutoFilter()
        End If

        ' Для нужных столбцов высчитываем сумму и среднее значение
        If (Not stgs.ColsNamesForSumCalc Is Nothing) OrElse (Not stgs.ColsNamesForAvgCalc Is Nothing) Then
            Dim Captions As New List(Of String)
            For i As Integer = 0 To dt.Columns.Count - 1
                Captions.Add(dt.Columns(i).Caption)
            Next

            If Not stgs.ColsNamesForSumCalc Is Nothing Then
                For nCol = 0 To stgs.ColsNamesForSumCalc.Count - 1
                    Dim idh = Array.IndexOf(Captions.ToArray(), stgs.ColsNamesForSumCalc(nCol))
                    xlws.Cells(stgs.FirstRow + dt.Rows.Count, stgs.FirstCol + idh).Value = "=SUM(" +
                            ChrW(64 + idh + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + idh + stgs.FirstCol) +
                            (stgs.FirstRow + dt.Rows.Count - 1).ToString + ")"
                Next
            End If

            If Not stgs.ColsNamesForAvgCalc Is Nothing Then

                Dim NeedOneMoreCell As Integer = 0
                If Not stgs.ColsNamesForSumCalc Is Nothing Then NeedOneMoreCell = 1

                For nCol = 0 To stgs.ColsNamesForSumCalc.Count - 1
                    Dim idh = Array.IndexOf(Captions.ToArray(), stgs.ColsNamesForSumCalc(nCol))
                    xlws.Cells(stgs.FirstRow + dt.Rows.Count + NeedOneMoreCell, stgs.FirstCol + idh).Value = "=Average(" +
                                ChrW(64 + idh + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + idh + stgs.FirstCol) +
                                (stgs.FirstRow + dt.Rows.Count - 1).ToString + ")"
                Next
            End If
        End If

        ' Фиксация определенных строк
        Try
            If Not stgs.CellToFix = Nothing Then
                xlws.Range(stgs.CellToFix).Select()
                xlws.Application.ActiveWindow.FreezePanes = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        ' Структура возвращается после исполнения функции для возможности
        ' дальнейшего использования данного Worksheet
        If stgs.NeedNumeration Then stgs.FirstCol -= 1
        stgs.FirstRow += dt.Rows.Count + 2 + PlaceForSumAndAvg
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

        Call SetDefaultValues(stgs)

        xlws.Cells.Font.Name = "Times New Roman"

        ' Костыль для подсчета используемых строк
        Dim PlaceForSumAndAvg As Integer = 0
        If Not stgs.ColsNamesForSumCalc Is Nothing Then PlaceForSumAndAvg += 1
        If Not stgs.ColsNamesForAvgCalc Is Nothing Then PlaceForSumAndAvg += 1

        Dim nRow As Integer, nCol As Integer, fRow As Integer = stgs.FirstRow, fCol As Integer = stgs.FirstCol

        If stgs.NameOfWorkSheet Is Nothing Then
            xlws.Name = "List" + xlws.Application.Workbooks(1).Worksheets.Count.ToString
        Else
            xlws.Name = stgs.NameOfWorkSheet
        End If

        ' Объединяем ячейки и вставляем туда название таблицы
        If Not stgs.NameOfTable Is Nothing Then
            xlws.Cells(fRow, fCol) = stgs.NameOfTable
            stgs.FirstRow += 1
            xlws.Cells(fRow, fCol).Font.Bold = True
            xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count) + fRow.ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            If stgs.ShowHiddenCols Then
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count + stgs.FirstCol) + fRow.ToString).MergeCells = True
            Else
                xlws.Range("" + ChrW(64 + fCol) + fRow.ToString + "", "" + ChrW(64 + dgv.Columns.Count + stgs.FirstCol - hiddenCols) + fRow.ToString).MergeCells = True
            End If

        End If

        xlws.Range(ChrW(64 + stgs.FirstCol) + (stgs.FirstRow).ToString, ChrW(64 + stgs.FirstCol) + (stgs.FirstRow + dgv.Rows.Count - 1).ToString).Font.Name = stgs.FontOfCells
        ' Проставляем номера строк
        If stgs.NeedNumeration Then
            For nRow = 1 To dgv.Rows.Count - 1
                xlws.Cells(stgs.FirstRow + nRow, stgs.FirstCol) = nRow
            Next
            stgs.FirstCol += 1

            ' Именования суммы и среднего значения
            If Not stgs.ColsNamesForSumCalc Is Nothing Then
                xlws.Cells(stgs.FirstRow + dgv.Rows.Count, fCol) = "Сумма"
                xlws.Rows(stgs.FirstRow + dgv.Rows.Count).Font.Bold = True
            End If
            If Not stgs.ColsNamesForAvgCalc Is Nothing Then
                Dim NeedOneMoreCell As Integer = 0
                If Not stgs.ColsNamesForSumCalc Is Nothing Then NeedOneMoreCell = 1
                xlws.Cells(stgs.FirstRow + dgv.Rows.Count + NeedOneMoreCell, fCol) = "Среднее"
                xlws.Rows(stgs.FirstRow + dgv.Rows.Count + NeedOneMoreCell).Font.Bold = True
            End If

        End If

        ' Проставляем имена столбцов
        If stgs.NeedNamesOfCols Then
            If stgs.NeedNumeration Then
                xlws.Cells(stgs.FirstRow, stgs.FirstCol - 1) = "Number"
            End If
            If Not stgs.RangeOfCellsForNamesOfColsFromAnotherFile Is Nothing Then

                Call TakeNamesFromAnotherFile(xlws, stgs)

            Else

                xlws.Rows(stgs.FirstRow).Font.Name = stgs.FontOfNames
                xlws.Rows(stgs.FirstRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                Dim idx As Integer = 0
                For nCol = 0 To dgv.Columns.Count - 1
                    If dgv.Columns(nCol).Visible Or stgs.ShowHiddenCols Then
                        xlws.Cells(stgs.FirstRow, stgs.FirstCol + idx) = dgv.Columns(nCol).HeaderText
                        idx += 1
                    End If
                Next nCol
            End If
            stgs.FirstRow += 1
        End If

        ' Заполнение ячеек
        If stgs.ShowHiddenCols Then
            For i = 0 To dgv.Columns.Count - 1

                colNumber = i

                Try
                    'Set the content from datatable (which Is converted as array And again converted as string) 
                    Clipboard.SetText(AryToString(ToArray(dgv.DataSource)))

                    'Identifiy And select the range of cells in Excel to paste the clipboard data. 
                    xlws.Cells(stgs.FirstRow, i + stgs.FirstCol).Select()

                    'Paste the clipboard data 
                    xlws.Paste()
                    Clipboard.Clear()
                Catch ex As Exception
                    For j = 0 To dgv.Rows.Count - 1
                        xlws.Cells(stgs.FirstRow + j, stgs.FirstCol + i).Value = dgv.DataSource.Rows(j)(i)
                    Next
                End Try
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

        xlws.Range("" + ChrW(64 + stgs.FirstCol) + (stgs.FirstRow).ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol) +
                 (stgs.FirstRow + dgv.Rows.Count - 1).ToString + "").Font.Name = stgs.FontOfCells

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
            ElseIf Boolean.TryParse(xlws.Cells(stgs.FirstRow + 1, stgs.FirstCol + idx).Value, boolTest) Then
                For i As Integer = 1 To dgv.Rows.Count
                    If Boolean.TryParse(xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value, boolTest) Then
                        If boolTest Then
                            xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value = stgs.replaceBoolTrueVal
                        Else
                            xlws.Cells(stgs.FirstRow + i - 1, stgs.FirstCol + idx).Value = stgs.replaceBoolFalseVal
                        End If
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

        ' Делаем автоширину с возможной максимальной шириной
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

        Dim CellForName As Integer = 0
        If Not stgs.NameOfTable Is Nothing Then CellForName = 1

        If stgs.NeedBorder Then
            If stgs.ShowHiddenCols Then
                xlws.Range("" + ChrW(64 + fCol) + (fRow + CellForName).ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol) +
                 (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Select()
            Else
                xlws.Range("" + ChrW(64 + fCol) + (fRow + CellForName).ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol - hiddenCols) +
                 (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Select()
            End If
            With xlws.Application.Selection().Borders '(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With

        End If

        If stgs.NeedGrid Then

            If stgs.ShowHiddenCols Then
                With xlws.Range("" + ChrW(64 + fCol) + (fRow + CellForName).ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol) +
                            (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                xlws.Range(ChrW(64 + fCol) + (fRow + CellForName).ToString, ChrW(64 + fCol + dgv.Columns.Count) + (fRow + CellForName).ToString).Select()

                With xlws.Application.Selection.Borders '(xlEdgeLeft)
                    .LineStyle = xlDouble
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
            Else
                With xlws.Range("" + ChrW(64 + fCol) + (fRow + CellForName).ToString + "", "" + ChrW(64 + dgv.Columns.Count - 1 + stgs.FirstCol - hiddenCols) +
                            (stgs.FirstRow + dgv.Rows.Count - 2 + PlaceForSumAndAvg).ToString + "").Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                xlws.Range(ChrW(64 + fCol) + (fRow + CellForName).ToString, ChrW(64 + fCol + dgv.Columns.Count) + (fRow + CellForName).ToString).Select()

                With xlws.Application.Selection.Borders '(xlEdgeLeft)
                    .LineStyle = xlDouble
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
            End If
        End If

        If Not stgs.ColsForAutoFilter Is Nothing Then
            If stgs.ColsForAutoFilter = "ALL" Then
                xlws.Rows(fRow + CellForName).Select()
            Else
                Dim cells As String() = stgs.ColsForAutoFilter.Split(New Char() {";", ":", ","})
                xlws.Range(ChrW(65 + cells(0)) + (fRow + 1).ToString, ChrW(65 + cells(1)) + (fRow + CellForName).ToString).Select()
            End If
            xlws.Application.Selection.AutoFilter()
        End If

        ' Расчет суммы и среднего значения для требуемых колонок

        If (Not stgs.ColsNamesForSumCalc Is Nothing) OrElse (Not stgs.ColsNamesForAvgCalc Is Nothing) Then
            Dim Captions As New List(Of String)
            For i As Integer = 0 To dgv.Columns.Count - 1
                If dgv.Columns(i).Visible Or stgs.ShowHiddenCols Then
                    Captions.Add(dgv.Columns(i).HeaderText)
                End If
            Next

            If Not stgs.ColsNamesForSumCalc Is Nothing Then
                For nCol = 0 To stgs.ColsNamesForSumCalc.Count - 1
                    Dim idh = Array.IndexOf(Captions.ToArray(), stgs.ColsNamesForSumCalc(nCol))
                    xlws.Cells(stgs.FirstRow + dgv.Rows.Count - 1, stgs.FirstCol + idh).Value = "=SUM(" +
                            ChrW(64 + idh + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + idh + stgs.FirstCol) +
                            (stgs.FirstRow + dgv.Rows.Count - 2).ToString + ")"
                Next
            End If

            If Not stgs.ColsNamesForAvgCalc Is Nothing Then

                Dim NeedOneMoreCell As Integer = 0
                If Not stgs.ColsNamesForSumCalc Is Nothing Then NeedOneMoreCell = 1

                For nCol = 0 To stgs.ColsNamesForSumCalc.Count - 1
                    Dim idh = Array.IndexOf(Captions.ToArray(), stgs.ColsNamesForSumCalc(nCol))
                    xlws.Cells(stgs.FirstRow + dgv.Rows.Count - 1 + NeedOneMoreCell, stgs.FirstCol + idh).Value = "=Average(" +
                                ChrW(64 + idh + stgs.FirstCol) + (stgs.FirstRow).ToString + ":" + ChrW(64 + idh + stgs.FirstCol) +
                                (stgs.FirstRow + dgv.Rows.Count - 2).ToString + ")"
                Next
            End If
        End If

        Try
            If Not stgs.CellToFix = Nothing Then
                xlws.Range(stgs.CellToFix).Select()
                xlws.Application.ActiveWindow.FreezePanes = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        If stgs.NeedNumeration Then stgs.FirstCol -= 1
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

    Private Sub ExcelClose(ByRef xl As Object, ByRef xlwb As Excel.Workbook, Optional ByVal SaveByPath As Boolean = False,
                           Optional ByVal SaveByDefaultPath As Boolean = False,
                           Optional ByVal password As String = Nothing, Optional ByVal SendMail As Boolean = False,
                           Optional ByVal Email As String = "empty_box44@mail.ru")
        xl.Visible = True
        If Not password Is Nothing Then
            xlwb.Password = password
        End If

        Dim saveFileDialog1 As New SaveFileDialog
        Try
            If SaveByPath Then
                saveFileDialog1.Filter = "(*.xlsx) | *.xlsx"
                saveFileDialog1.Title = "Save an Excel File"
                saveFileDialog1.ShowDialog()
                If saveFileDialog1.FileName <> "" Then
                    xlwb.SaveAs(saveFileDialog1.FileName)
                End If
                xlwb.Close(False)
                xl.Quit()
            End If
            If SaveByDefaultPath Then
                Dim s As String = Directory.GetCurrentDirectory + "\Report" + DateTime.Now.Day.ToString + "." +
                DateTime.Now.Month.ToString + "." + DateTime.Now.Year.ToString + ".xlsx"
                xlwb.SaveAs(s)
                'xl.ActiveWorkBook.Close()
                xl.Quit()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Try
                'xl.Quit()
                xlwb = Nothing
                xl = Nothing
                GC.Collect()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Try

        If SendMail Then
            SendEmail(saveFileDialog1.FileName, Email)
        End If

    End Sub

    Private Sub FixesForPrint(xlws As Excel.Worksheet, Optional ByVal NeedNumeration As Boolean = False, Optional ByVal LandscapeOrientation As Boolean = False,
                           Optional ByVal PrintArea As String = "A1:A2")
        If NeedNumeration Then
            xlws.PageSetup.CenterFooter = "&P/&N"
        End If

        If LandscapeOrientation Then
            xlws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape ' Вид печати
        Else
            xlws.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait
        End If

        xlws.PageSetup.PrintArea = PrintArea     ' Зона печати

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

        Dim xl As Excel.Application = StartExcel()
        Dim xlwb As Excel.Workbook = xl.Workbooks.Add()
        Dim xlws As Excel.Worksheet = xlwb.Worksheets.Add()
        xlwb.Sheets(2).Delete()

        Dim cn As New SqlClient.SqlConnection("Data Source=212.42.46.12;Initial Catalog=CBR_Emm_test;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY")
        Dim cmd As New SqlClient.SqlCommand(Nothing, cn)
        cmd.CommandText = "select top 5 *from Stock_Dr"
        Dim adap As New SqlClient.SqlDataAdapter(cmd)
        Dim dt As New DataTable
        adap.Fill(dt)



        Dim s As New Settings
        s.NameOfWorkSheet = "ListName1"
        s.NeedGrid = True
        's.NameOfTable = "Hello"
        s.NeedNamesOfCols = True
        s.ColsForAutoFilter = "2:4"
        s.ColsNamesForAvgCalc = New String() {"ISSUESIZE"}
        s.ColsNamesForSumCalc = New String() {"ISSUESIZE"}
        s.NeedNumeration = True
        s.NeedAutoWidth = True
        's.RangeOfCellsForNamesOfColsFromAnotherFile = "B1:M1"
        's.CellToFix = "C5"


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
        's.ColsForAutoFilter = Nothing
        's.RangeOfCellsForNamesOfColsFromAnotherFile = Nothing
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

        Call ExcelClose(xl, xlwb)

    End Sub
End Class
