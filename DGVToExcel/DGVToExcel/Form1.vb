Imports Microsoft.Office.Interop
Public Class Form1

    Structure Settings
        Dim FirstCol As Integer
        Dim FirstRow As Integer
        Dim NeedBorder As Boolean
        Dim NeedNamesOfCols As Boolean
        Dim NameOfTable As String
        Dim FontOfNames As String
        Dim FontOfCells As String
        Dim FormatOfCols As String
        Dim ColsNamesForSumOrAvgCalc() As String
        Dim IndexesOfColsForAutoFilter As String
        Dim NameOfWorkSheet As String
        Dim NeedNumeration As Boolean
        ' -	закрепить ли области для удобного просмотра на экране (уточнить)
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

    Private Sub DataGridToExcelSheet(ByVal dgv As DataGridView, ByVal stgs As Settings)
        Dim xl As Excel.Application = StartExcel(False, 1)
        Dim xlws As Excel.Worksheet

        Dim nRow As Integer, nCol As Integer
        Dim strow As Integer, stcol As Integer

        strow = stgs.FirstRow
        stcol = stgs.FirstCol

        If stgs.NeedNamesOfCols Then
            For nCol = 0 To dgv.Columns.Count - 1
                If dgv.Columns(nCol).Visible Or stgs.ShowHiddenCols Then
                    xlws.Cells(strow, stcol + nCol) = dgv.Columns(nCol).HeaderText
                Else
                    stcol -= 1
                End If
            Next nCol
            strow += 1
        End If

    End Sub
End Class
