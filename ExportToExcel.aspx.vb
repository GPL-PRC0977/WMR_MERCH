Imports System.Data.SqlClient
Imports System.Data
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Drawing


Public Class ExportToExcel
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim cmd_openUpload As String
        Dim ds_openUpload As New DataSet

        cmd_openUpload = "select * from [t_Uploaded_Document] where [DocNo] = '" & Session("selectedDocwithoutdash") & "' and [dateuploaded] = '" & Date.Now & "'"
        ds_openUpload = executeQuery(cmd_openUpload)
        If ds_openUpload.Tables(0).Rows.Count > 0 Then
            executeQuery("delete from t_Uploaded_Document where [DocNo] = '" & Session("selectedDocwithoutdash") & "' and [dateuploaded] = '" & Date.Now & "'")
        End If

        exporttoExcel()
        Response.Redirect("MR.aspx")
    End Sub

    Protected Sub exporttoExcel()
        Dim cmd_updater As String
        Dim ds_updater As New DataSet
        Dim cmd_whsw_qty As String
        Dim ds_whse_qty As New DataSet
        Dim cmd_whsw_qty2 As String
        Dim ds_whse_qty2 As New DataSet
        Try
            cmd_updater = "exec sp_selectUpdate_dynamicTable '" & Replace(Session("selectedDoc"), "-", "") & "'"
            executeQuery(cmd_updater)

            Dim cmd_series As String
            Dim ds_series As DataSet
            Dim series_ As Integer
            cmd_series = "select top 1 [download_series] from Downloading_Series order by [ID] desc"
            ds_series = executeQuery(cmd_series)
            series_ = ds_series.Tables(0).Rows(0)("download_series")

            Dim increment_series As Integer
            increment_series = Val(series_) + 1
            cmd_series = "update Downloading_Series set [download_series] = '" & increment_series & "'"
            'cmd_series = "insert into Downloading_Series ([download_series],[docno],[downloaded_by],[date_downloaded]) values ('" & increment_series & "','" & Session("selectedDoc") & "','" & Session("tmpuser") & "','" & Date.Now & "')"
            executeQuery(cmd_series)

            Dim cmd_download As String
            cmd_download = "insert into TO_MR_Log2 ([user name],[date],[doc no],[action]) values ('" & Session("tmpuser") & "','" & DateAndTime.Now & "','" & Session("selectedDocwithoutdash") & "','download')"
            executeQuery(cmd_download)


            Dim xls As Microsoft.Office.Interop.Excel.Application
            Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
            Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value

            xls = New Microsoft.Office.Interop.Excel.Application
            xlsWorkBook = xls.Workbooks.Add
            xlsWorkSheet = xlsWorkBook.Sheets(1)

            '// select columns from table
            Dim cmd_columns As String
            Dim ds_columns As New DataSet
            cmd_columns = "exec sp_collect_columns '" & Session("selectedDoc") & "','',''"
            ds_columns = executeQuery(cmd_columns)

            '// populate header columns to excel
            Dim x As Integer = 1
            For i = 0 To ds_columns.Tables(0).Rows.Count - 1
                xlsWorkSheet.Cells(1, x) = ds_columns.Tables(0).Rows(i)("COLUMN_NAME")
                x = x + 1
            Next

            Dim cmd_data As String
            Dim ds_data As New DataSet
            cmd_data = "exec sp_select_dynamicTable '" & Session("selectedDoc") & "','',''"
            ds_data = executeQuery(cmd_data)

            Dim rowstart As Integer

            If My.Settings.ViewDetails = 1 Then
                rowstart = 1
            Else
                rowstart = 2
            End If

            Dim ii As Integer
            Dim alphabet_increment As String
            Dim zplus As Integer = 0

            '// bold header
            xlsWorkSheet.Range("a1:cz1").Font.Bold = True

            Dim style As Excel.Style = xlsWorkSheet.Application.ActiveWorkbook.Styles.Add("NewStyle")
            style.Font.Bold = True
            style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#e6f2ff"))
            style.Font.Color = Color.Red


            For r = 0 To (ds_data.Tables(0).Rows.Count) - 1 '// for rows
                For c = 0 To ds_columns.Tables(0).Rows.Count - 1 '// for columns


                    If My.Settings.ViewDetails <> 1 Then
                        If c > 25 And c < 52 Then
                            alphabet_increment = "a" & Chr(Asc("a") + Val(zplus))
                            zplus = zplus + 1
                            xlsWorkSheet.Range("a" & r + 2 & ":" & alphabet_increment & r + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        ElseIf c > 51 And c < 78 Then
                            alphabet_increment = "b" & Chr(Asc("a") + Val(zplus) - 26)
                            zplus = zplus + 1
                            xlsWorkSheet.Range("b" & r + 2 & ":" & alphabet_increment & r + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        ElseIf c > 77 And c < 104 Then
                            alphabet_increment = "c" & Chr(Asc("a") + Val(zplus) - 52)
                            zplus = zplus + 1
                            xlsWorkSheet.Range("c" & r + 2 & ":" & alphabet_increment & r + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        ElseIf c > 103 And c < 130 Then
                            alphabet_increment = "d" & Chr(Asc("a") + Val(zplus) - 104)
                            zplus = zplus + 1
                            xlsWorkSheet.Range("d" & r + 2 & ":" & alphabet_increment & r + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        Else
                            alphabet_increment = Chr(Asc("a") + Val(c))
                            xlsWorkSheet.Range("a" & r + 2 & ":" & alphabet_increment & r + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        End If

                        xlsWorkSheet.Range(alphabet_increment & c + 1).ColumnWidth = Len(ds_columns.Tables(0).Rows(c)("COLUMN_NAME")) + 5 '// adjust columns width
                        xlsWorkSheet.Cells(rowstart, c + 1) = ds_data.Tables(0).Rows(r)(ds_columns.Tables(0).Rows(c)("COLUMN_NAME")) '// populate all data 
                    Else
                        If c > 25 And c < 52 Then
                            alphabet_increment = "a" & Chr(Asc("a") + Val(zplus))
                            zplus = zplus + 1
                            xlsWorkSheet.Range("a" & r * 2 + 2 + 1 & ":" & alphabet_increment & r * 2 + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        ElseIf c > 51 And c < 78 Then
                            alphabet_increment = "b" & Chr(Asc("a") + Val(zplus) - 26)
                            zplus = zplus + 1
                            xlsWorkSheet.Range("b" & r * 2 + 2 + 1 & ":" & alphabet_increment & r * 2 + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        ElseIf c > 77 And c < 104 Then
                            alphabet_increment = "c" & Chr(Asc("a") + Val(zplus) - 52)
                            zplus = zplus + 1
                            xlsWorkSheet.Range("c" & r * 2 + 2 + 1 & ":" & alphabet_increment & r * 2 + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        ElseIf c > 103 And c < 130 Then
                            alphabet_increment = "d" & Chr(Asc("a") + Val(zplus) - 104)
                            zplus = zplus + 1
                            xlsWorkSheet.Range("d" & r * 2 + 2 + 1 & ":" & alphabet_increment & r * 2 + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        Else
                            alphabet_increment = Chr(Asc("a") + Val(c))
                            xlsWorkSheet.Range("a" & r * 2 + 2 + 1 & ":" & alphabet_increment & r * 2 + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                        End If

                        xlsWorkSheet.Range(alphabet_increment & c + 1).ColumnWidth = Len(ds_columns.Tables(0).Rows(c)("COLUMN_NAME")) + 5 '// adjust columns width
                        xlsWorkSheet.Cells(rowstart * 2, c + 1) = ds_data.Tables(0).Rows(r)(ds_columns.Tables(0).Rows(c)("COLUMN_NAME")) '// populate all data 
                        xlsWorkSheet.Cells(rowstart * 2 + 1, c + 1) = "-"
                        xlsWorkSheet.Range("a" & rowstart * 2 + 1 & ":" & alphabet_increment & rowstart * 2 + 1).Style = "NewStyle"
                        xlsWorkSheet.Range("a" & rowstart * 2 + 1 & ":" & alphabet_increment & rowstart * 2 + 1).Font.Size = 9
                        xlsWorkSheet.Range("a" & rowstart * 2 + 1 & ":" & alphabet_increment & rowstart * 2 + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        xlsWorkSheet.Range("a" & rowstart * 2 + 1 & ":" & alphabet_increment & rowstart * 2 + 1).HorizontalAlignment = HorizontalAlign.Left
                        xlsWorkSheet.Range("a" & rowstart * 2 + 1 & ":" & alphabet_increment & rowstart * 2 + 1).NumberFormat = "@"

                        If c = 12 Then
                            xlsWorkSheet.Cells(rowstart * 2 + 1, c + 1) = "On Hand"
                        ElseIf c > 12 Then
                            cmd_whsw_qty = "SELECT [Store Qty]
                                  FROM [TOMR].[dbo].[TO_MR_Invt_Sales] where [Location Name] = '" & ds_columns.Tables(0).Rows(c)("COLUMN_NAME") & "'
                                  and [Item No] = '" & ds_data.Tables(0).Rows(r)(ds_columns.Tables(0).Rows(2)("COLUMN_NAME")) & "'"
                            ds_whse_qty = executeQuery(cmd_whsw_qty)
                            If ds_whse_qty.Tables(0).Rows.Count = 0 Then
                                xlsWorkSheet.Cells(rowstart * 2 + 1, c + 1) = "0"
                            Else
                                If Val(ds_whse_qty.Tables(0).Rows(0)("Store Qty")) < 0 Then
                                    xlsWorkSheet.Cells(rowstart * 2 + 1, c + 1) = 0
                                Else
                                    xlsWorkSheet.Cells(rowstart * 2 + 1, c + 1) = ds_whse_qty.Tables(0).Rows(0)("Store Qty")
                                End If
                            End If

                        End If
                    End If

                Next
                rowstart = rowstart + 1
                zplus = 0
            Next

            '// save file
            Dim xfile As String
            xfile = Server.MapPath("~\ExcelExport") & "\" & Session("selectedDocwithoutdash") & "-" & Format(Date.Now, "MMdd") & "-" & series_ & "-" & "(" & Replace(Replace(Session("tmpuser").ToString.ToLower, "primergrp\", ""), ".", " ") & ")" & ".xlsx"
            xls.DisplayAlerts = False
            xlsWorkBook.SaveAs(xfile)

            xlsWorkBook.Close()
            xls.Quit()
            xlsWorkBook = Nothing
            xlsWorkSheet = Nothing
            xls = Nothing
            System.GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

            '// download file
            Response.ContentType = "file/xlsx"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & Session("selectedDocwithoutdash") & ".xlsx")
            Response.TransmitFile(Server.MapPath("~/ExcelExport/" & Session("selectedDocwithoutdash") & "-" & Format(Date.Now, "MMdd") & "-" & series_ & "-" & "(" & Replace(Replace(Session("tmpuser").ToString.ToLower, "primergrp\", ""), ".", " ") & ")" & ".xlsx"))
            Response.End()
        Catch ex As Exception
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
            Exit Try
        End Try

    End Sub
End Class