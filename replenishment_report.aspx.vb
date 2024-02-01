Imports System.Data.Sql
Imports Microsoft.Office.Interop
Public Class replenishment_report
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim cmd As String
        Dim ds As New DataSet

        btn_download.Enabled = False

        Try
            cmd = "SELECT distinct [Company Ownership Code] FROM [TOMR].[dbo].[TO_MR_Warehouse]"
            ds = executeQuery(cmd)

            If ds.Tables(0).Rows.Count > 0 Then
                If ddcompany.Text = "" Then
                    ddcompany.Items.Clear()
                    ddcompany.Items.Add("")
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        ddcompany.Items.Add(ds.Tables(0).Rows(i)("Company Ownership Code"))
                    Next
                    ddcompany.Items.Add("ALL COMPANY")
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_download_Click(sender As Object, e As EventArgs) Handles btn_download.Click
        Dim xls As Microsoft.Office.Interop.Excel.Application
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        xls = New Microsoft.Office.Interop.Excel.Application
        xlsWorkBook = xls.Workbooks.Add
        xlsWorkSheet = xlsWorkBook.Sheets(1)

        xlsWorkSheet.Range("a1").ColumnWidth = 15
        xlsWorkSheet.Range("b1").ColumnWidth = 15
        xlsWorkSheet.Range("c1").ColumnWidth = 15
        xlsWorkSheet.Range("d1").ColumnWidth = 25
        xlsWorkSheet.Range("e1").ColumnWidth = 15

        xlsWorkSheet.Range("a1").Font.Bold = True
        xlsWorkSheet.Range("b1").Font.Bold = True
        xlsWorkSheet.Range("c1").Font.Bold = True
        xlsWorkSheet.Range("d1").Font.Bold = True
        xlsWorkSheet.Range("e1").Font.Bold = True

        xlsWorkSheet.Cells(1, 1) = "DOCUMENT NO"
        xlsWorkSheet.Cells(1, 2) = "DATE"
        xlsWorkSheet.Cells(1, 3) = "BRAND"
        xlsWorkSheet.Cells(1, 4) = "LOCATION"
        xlsWorkSheet.Cells(1, 5) = "TOTAL QTY SERVED"

        Dim x As Integer = 2
        xlsWorkSheet.Range("a1:e1").Borders.LineStyle = Excel.XlLineStyle.xlContinuous

        For i = 0 To dg.Rows.Count - 1
            xlsWorkSheet.Range("a" & x & ":" & "e" & x).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            xlsWorkSheet.Cells(x, 1) = dg.Rows(i).Cells(0).Text
            xlsWorkSheet.Cells(x, 2) = dg.Rows(i).Cells(1).Text
            xlsWorkSheet.Cells(x, 3) = dg.Rows(i).Cells(2).Text
            xlsWorkSheet.Cells(x, 4) = dg.Rows(i).Cells(3).Text
            xlsWorkSheet.Cells(x, 5) = dg.Rows(i).Cells(4).Text
            x = x + 1
        Next

        Dim xfile As String
        xfile = Server.MapPath("~\reportsDownload") & "\" & "Reports.xlsx"
        xls.DisplayAlerts = False
        xlsWorkBook.SaveAs(xfile)
        xlsWorkBook.Close()
        xls.Quit()

        Response.ContentType = "file/xlsx"
        Response.AppendHeader("Content-Disposition", "attachment; filename=Reports.xlsx")
        Response.TransmitFile(Server.MapPath("~/reportsDownload/Reports.xlsx"))
        Response.End()

    End Sub

    Private Sub btn_search_Click(sender As Object, e As EventArgs) Handles btn_search.Click

        Dim cmd As String
        Dim ds As New DataSet

        Try
            cmd = "exec get_brands '" & txt_date_from.Text & "','" & txt_date_to.Text & "','" & Replace(GetSelectedValues(dd_brands), "'", "''") & "'"

            ds = executeQuery(cmd)
            If ds.Tables(0).Rows.Count > 0 Then
                dg.DataSource = ds.Tables(0)
                dg.DataBind()
                btn_download.Enabled = True
            Else
                btn_download.Enabled = False
                dg.DataSource = Nothing
                dg.DataBind()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ddcompany_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddcompany.SelectedIndexChanged
        Dim cmd As String
        Dim ds As New DataSet

        Try
            If ddcompany.Text = "ALL COMPANY" Then
                cmd = "select [Brand Description] from t_Brands where [Company Ownership Code] <> '' order by [Brand Description]  asc"
            Else
                cmd = "select [Brand Description] from t_Brands where [Company Ownership Code] = '" & ddcompany.Text & "'"
            End If

            ds = executeQueryRS(cmd)
            If ds.Tables(0).Rows.Count > 0 Then
                dd_brands.Items.Clear()

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dd_brands.Items.Add(ds.Tables(0).Rows(i)("Brand Description"))
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class