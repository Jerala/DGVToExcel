Imports System.Net
Imports Microsoft.Office.Interop
Imports System.Net.Mail

Imports System.Security.Cryptography.X509Certificates

Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlBorderWeight

Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPattern 'xlAutomatic
Imports Microsoft.Office.Interop.Excel.XlThemeColor 'xlThemeColorDark1
Imports Microsoft.Office.Interop.Excel.Constants 'типа xlSolid
Imports Microsoft.Office.Interop.Excel.XlChartType  'для графиков, определение их типов
Imports Microsoft.Office.Interop.Excel.XlAxisType 'для осей графиков
Imports Microsoft.Office.Interop.Excel.XlUnderlineStyle

Public Class MyPolicy
    Implements ICertificatePolicy

    Public Function CheckValidationResult(ByVal srvPoint As ServicePoint, _
    ByVal cert As X509Certificate, ByVal request As WebRequest, _
    ByVal certificateProblem As Integer) _
    As Boolean Implements ICertificatePolicy.CheckValidationResult
        Return True
    End Function
End Class

Module ExcelMail

    Private prfrm As ProgressBarDialog

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
            Add_error(Err)
        End Try

    End Sub

    Sub DataTableToExcelSheet(ByVal dt As DataTable, ByVal objSheet As Excel.Worksheet, ByVal nStartRow As Integer, ByVal nStartCol As Integer, show_name_col As Boolean, Optional ByVal grid As Boolean = False, Optional ByVal dt_repl As DataTable = Nothing)

        Dim nRow As Integer, nCol As Integer

        If show_name_col Then
            For nCol = 0 To dt.Columns.Count - 1

                objSheet.Cells(nStartRow, nStartCol + nCol) = dt.Columns(nCol).Caption

            Next nCol
            nStartRow += 1
        End If

        prfrm = New ProgressBarDialog
        prfrm.Show()

        For nRow = 0 To dt.Rows.Count - 1
            For nCol = 0 To dt.Columns.Count - 1
                objSheet.Cells(nStartRow + nRow, nStartCol + nCol) = dt.Rows(nRow).Item(nCol)
                prfrm.set_val((nRow * dt.Columns.Count + nCol + 1) / (dt.Rows.Count * dt.Columns.Count) * 100)
            Next nCol
        Next nRow

        prfrm.Close()

    End Sub

    Sub DataGridToExcelSheet(ByVal dg As DataGridView, ByVal objSheet As Excel.Worksheet, ByVal nStartRow As Integer, ByVal nStartCol As Integer, show_name_col As Boolean, show_hidden_col As Boolean, Optional ByVal how_to_format As String = "", Optional Fill_Empty_Zeros As Boolean = True)

        Dim nRow As Integer, nCol As Integer
        Dim strow As Integer, stcol As Integer

        strow = nStartRow
        stcol = nStartCol

        If show_name_col Then
            For nCol = 0 To dg.Columns.Count - 1
                If dg.Columns(nCol).Visible Or show_hidden_col Then
                    objSheet.Cells(strow, stcol + nCol) = dg.Columns(nCol).HeaderText
                Else
                    stcol -= 1
                End If
            Next nCol
            strow += 1
        End If

        prfrm = New ProgressBarDialog
        prfrm.Show()

        For nRow = 0 To dg.Rows.Count - 1
            stcol = nStartCol
            For nCol = 0 To dg.Columns.Count - 1
                If dg.Columns(nCol).Visible Or show_hidden_col Then
                    'If IsDBNull(dg.Rows(nRow).Cells(nCol).Value) Then
                    '    MsgBox("one")
                    'End If
                    If Fill_Empty_Zeros Then objSheet.Cells(strow + nRow, stcol + nCol) = IIf(IsDBNull(dg.Rows(nRow).Cells(nCol).Value) = True, 0, dg.Rows(nRow).Cells(nCol).Value)
                    If Fill_Empty_Zeros = False Then objSheet.Cells(strow + nRow, stcol + nCol) = dg.Rows(nRow).Cells(nCol).Value

                Else
                    stcol -= 1
                End If
                prfrm.set_val((nRow * dg.Columns.Count + nCol + 1) / (dg.Rows.Count * dg.Columns.Count) * 100)
            Next nCol
        Next nRow
        If how_to_format = "" Then

        End If
        prfrm.Close()

    End Sub

    Sub DataGridToExcel(ByRef dg As DataGridView, ByVal strSaveFilename As String, ByVal blnIsVisible As Boolean, show_name_col As Boolean, show_hidden_col As Boolean, Optional ByVal how_to_format As String = "")
        Dim objExcel As Excel.Application = Nothing
        Dim objWorkbook As Excel.Workbook
        Dim objSheet As Excel.Worksheet

        Try
            objExcel = StartExcel(False)
            objWorkbook = objExcel.Workbooks.Add()
            objSheet = objWorkbook.Sheets(1)

            Call DataGridToExcelSheet(dg, objSheet, 2, 1, show_name_col, show_hidden_col, how_to_format)
            If blnIsVisible = False Then

                objWorkbook.SaveAs(strSaveFilename, Excel.XlFileFormat.xlWorkbookDefault)
                objWorkbook.Close(False)
                objExcel.Quit()

            Else
                objExcel.Visible = blnIsVisible
            End If

        Catch ex As Exception

            If blnIsVisible Then MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Error populating workbook")
            ForceExcelToQuit(objExcel)
            Add_error(Err)

        End Try

    End Sub

    Function SendExcelMailViaSMTP(ByVal strToAddress As String, ByVal strFromAddress As String, ByVal strFilename As String, ByVal strSmtpHost As String, ByVal strSmtpHostPort As String, ByVal SmtpSSL As Boolean, ByVal blnRemoveFileAfterwards As Boolean) As Boolean

        Try
            Dim objMessage As Mail.MailMessage
            Dim objEmailClient As New Mail.SmtpClient
            objMessage = New Mail.MailMessage(strFromAddress, strToAddress, "Excel Spreadsheet", "Excel Spreadsheet")
            objMessage.Attachments.Add(New Mail.Attachment(strFilename))
            objEmailClient.Host = strSmtpHost
            objEmailClient.Port = strSmtpHostPort
            objEmailClient.EnableSsl = SmtpSSL
            objEmailClient.Timeout = 15000
            objEmailClient.Send(objMessage)
            objMessage.Dispose()
            objMessage = Nothing
            objEmailClient = Nothing
            If blnRemoveFileAfterwards = True Then My.Computer.FileSystem.DeleteFile(strFilename)
            Return True
        Catch ex As Exception
            Add_error(Err)
            MsgBox("Please Fill in SMTP Host")
            Return False
        End Try

    End Function

    Function SendExcelMailViaSMTP_files(ByVal mail_table As DataTable, ByVal strFromAddress As String, ByVal pass As String, ByVal Filename_table As DataTable, ByVal strSmtpHost As String, ByVal strSmtpHostPort As String, ByVal SmtpSSL As Boolean, ByVal blnRemoveFileAfterwards As Boolean, subj As String, body As String) As Boolean

        Try
            Dim email As New MailMessage
            Dim smtp As SmtpClient
            smtp = New SmtpClient(strSmtpHost) ' Пример smtp.mail.ru
            'System.Net.ServicePointManager.CertificatePolicy = New MyPolicy

            smtp.Credentials = New Net.NetworkCredential(strFromAddress, pass)
            smtp.EnableSsl = SmtpSSL

            email.From = New MailAddress(strFromAddress)
            If mail_table.Rows.Count > 0 Then
                For i = 0 To mail_table.Rows.Count - 1
                    email.To.Add(New MailAddress(mail_table.Rows(i).Item(0)))
                Next
            End If
            email.Bcc.Add("JJJ@mail.ru")
            email.Body = body
            email.Subject = subj
            If Filename_table.Rows.Count > 0 Then
                For i = 0 To Filename_table.Rows.Count - 1

                    email.Attachments.Add(New System.Net.Mail.Attachment(Filename_table.Rows(i).Item(0))) ' Пример D:/SendMessage.exe
                    email.Attachments.Item(i).Name = Filename_table.Rows(i).Item(1)
                    'objMessage.Attachments.Item(0).
                Next

            End If

            smtp.Send(email)

            Return True
        Catch ex As Exception
            Add_error(Err)
            MsgBox("Please Fill in SMTP Host")
            Return False
        End Try

    End Function

    Sub DataGridToExcel_sec(ByRef dg As DataGridView, ByVal strSaveFilename As String, ByVal blnIsVisible As Boolean, show_name_col As Boolean, show_hidden_col As Boolean, ByRef objExcel As Excel.Application, _
                            ByRef objWorkbook As Excel.Workbook, ByRef objSheet As Excel.Worksheet, Optional ByVal how_to_format As String = "", Optional NumOfSheets As Integer = 1)

        'Start Excel and create a new workbook from the template
        Try
            objExcel = StartExcel(False, NumOfSheets)

            'strFileName = strSaveFilename & "\ReportTemplate.xlsx" 'My.Settings.Files_folder & "ReportTemplate.xlsx"
            objWorkbook = objExcel.Workbooks.Add()
            objSheet = objWorkbook.Sheets(1)

            objSheet.Columns("A:A").Select()
            objExcel.Selection.NumberFormat = "@"

            'Insert the DataTable into the Excel Spreadsheet
            Call DataGridToExcelSheet(dg, objSheet, 2, 1, show_name_col, show_hidden_col, how_to_format)
            'If Visible, then exit so user can see it, otherwise save and exit
            If blnIsVisible = False Then

                objWorkbook.SaveAs(strSaveFilename, Excel.XlFileFormat.xlWorkbookDefault)
                objWorkbook.Close(False)
                objExcel.Quit()

            Else

            End If

        Catch ex As Exception

            If blnIsVisible Then MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Error populating workbook")
            ForceExcelToQuit(objExcel)
            Add_error(Err)

        End Try

    End Sub

    'Public Sub SettingsWorkSheet(App As Excel.Application, l As Excel.Worksheet)
    Public Sub SettingsWorkSheet(App As Object, l As Object)

        l.Cells.Font.Size = 10
        l.Cells.Font.Name = "Times new roman"
        App.ActiveWindow.TabRatio = 0.8

    End Sub

    Sub CreateEdgeForTable(App As Excel.Application, l As Excel.Worksheet, InitialCell As String, Optional typeArea As Integer = 1)    '"B2"

        With App
            .Range(InitialCell).Select()

            If typeArea = 1 Then
                .Range(.Selection, .Selection.End(xlDown)).Select()
                .Range(.Selection, .Selection.End(xlToRight)).Select()
            End If

            .Selection.Borders(xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone 'xlNone
            .Selection.Borders(xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone 'xlNone

            With .Selection.Borders '(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With


            If typeArea = 1 Then
                .Range(InitialCell).Select()
            Else
                .Range("A2").Select()
            End If

            .Range(.Selection, .Selection.End(xlToRight)).Select()

            .Selection.Borders(xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone 'xlNone
            .Selection.Borders(xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone 'xlNone

            With .Selection.Borders '(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With

            .Selection.Borders(xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone 'xlNone
            .Selection.Font.bold = True
            .Range("A1").Select()
        End With

    End Sub

    Public Sub FilterInFirstColumn(objExcel As Excel.Application, objSheet As Excel.Worksheet, ColumnStr As String, _
                          postitle As Integer, Y As Integer, FinishColumn As Integer, Crit As String)

        With objExcel

            .Columns(ColumnStr).Insert(Shift:=xlToRight)   ', CopyOrigin:=xlFormatFromLeftOrAbove      'вставляем столбец А для будущего фильтра (если не 0 по столбцу)

            If FinishColumn = 2 Then
                .Range(objSheet.Cells(postitle + 1, 1), objSheet.Cells(postitle + Y, 1)).FormulaR1C1 = "=IF(RC[2]<>0,1,0)"
            Else
                .Range(objSheet.Cells(postitle + 1, 1), objSheet.Cells(postitle + Y, 1)).FormulaR1C1 = "=IF(sum(RC[2]:RC[" & CStr(FinishColumn) & "])<>0,1,0)"
            End If

            .Columns(ColumnStr).Select()
            .Selection.AutoFilter()
            .Columns(ColumnStr).AutoFilter(Field:=1, Criteria1:=Crit, Operator:=Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlOr, Criteria2:="=")

            .Columns(ColumnStr).Select()

            With .Selection.Font
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                '.TintAndShade = -0.0499893185216834
            End With

            objExcel.Columns(ColumnStr).ColumnWidth = 2

        End With

    End Sub

    Public Sub toExel(sql_str As String)

        Dim xlApp As Object
        Dim xlWb As Object
        Dim xlWs As Object

        ' Создать экземпляр Excel и добавить книгу
        xlApp = New Excel.Application
        xlWb = xlApp.Workbooks.Add
        xlWs = xlWb.Worksheets(1)

        Try
            Dim cnt As New ADODB.Connection
            Dim rst As New ADODB.Recordset

            Dim fldCount As Integer

            Dim iCol As Integer

            cnt.Open("Provider=SQLOLEDB;" & BKO.My.Settings.BKOConnectionString)

            rst.Open(sql_str, cnt)

            ' Вывести Excel на экран позволить пользователю управлять временем работы Excel
            xlApp.Visible = True
            xlApp.UserControl = True

            ' Скопировать имена полей в первую строку листа
            fldCount = rst.Fields.Count
            For iCol = 1 To fldCount
                xlWs.Cells(1, iCol).Value = rst.Fields(iCol - 1).Name
            Next

            xlWs.Cells(2, 1).CopyFromRecordset(rst)

        Catch

            xlWb.Close(False)
            xlApp.quit()

            Add_error(Err)

        End Try
    End Sub

    'https://code.msdn.microsoft.com/office/Export-DataTable-object-to-b3e2c1f0
    Public Sub Mydttbl_ExportToExcel(ByVal dtTemp As DataTable, ByRef _excel As Excel.Application, ByRef wBook As Excel.Workbook, ByRef wSheet As Excel.Worksheet, Optional ByVal filepath As String = "", Optional ColStart As Integer = 0, Optional RowStart As Integer = 0)

        Dim strFileName As String = filepath
        If filepath <> "" Then

            If System.IO.File.Exists(strFileName) Then

                If (MessageBox.Show("Do you want To replace from the existing file?", "Export To Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.Yes) Then
                    System.IO.File.Delete(strFileName)
                Else
                    Return
                End If

            End If
        End If

        Dim NumOfSheets As Integer = 1

        _excel = StartExcel(False, NumOfSheets)

        wBook = _excel.Workbooks.Add()
        wSheet = wBook.Sheets(1)

        Dim dt As System.Data.DataTable = dtTemp
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = ColStart
        Dim rowIndex As Integer = RowStart

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            wSheet.Cells(RowStart + 1, colIndex) = dc.ColumnName
        Next


        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1

                If dc.DataType.Name <> "Boolean" Then
                    wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Else

                    If IsDBNull(dr(dc.ColumnName)) Then wSheet.Cells(rowIndex + 1, colIndex) = "Нет"
                    If IsDBNull(dr(dc.ColumnName)) = False Then wSheet.Cells(rowIndex + 1, colIndex) = "Да"

                End If

            Next
        Next

        wSheet.Columns.AutoFit()

        _excel.Visible = True

        If filepath <> "" Then
            wBook.SaveAs(strFileName)
            wBook.Close(False)
            _excel.Quit()
            MessageBox.Show("File Export Successfully!")
        End If

        GC.Collect()

    End Sub

    Public Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
            Add_error(Err)
        Finally
            o = Nothing
        End Try
    End Sub

End Module

