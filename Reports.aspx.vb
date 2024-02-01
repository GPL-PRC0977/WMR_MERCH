Imports System.Data.SqlClient
Imports System.Data
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Drawing
Public Class Reports
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mainnav As Control = Master.FindControl("main_navigation")
        mainnav.Visible = False
        backgrounddiv.Visible = False
        deleted_items_div.Visible = False

        If CType(Session.Item("tmpuser"), String) = "" Then
            Page.Response.Redirect("Default.aspx")
        End If

        Dim cmd As String
        Dim ds As New DataSet


        If Not IsPostBack Then

            cmd = "select distinct [SiteConcept] from [Reports].[dbo].[t_locations4] where [Company Ownership Code] In (Select distinct [company] from [TOMR].[dbo].[TO_MR_User_Security] where [user name] = '" & Session("tmpuser") & "') and [SiteConcept] not like '%NONE%' order by [SiteConcept]"
            ds = executeQuery(cmd)
            For i = 0 To ds.Tables(0).Rows.Count - 1
                chk_siteConcept.Items.Add(ds.Tables(0).Rows(i)("SiteConcept"))
            Next
        End If

    End Sub

    Private Sub chk_selectAll_CheckedChanged(sender As Object, e As EventArgs) Handles chk_selectAll.CheckedChanged
        Dim storenames As String
        Dim brandnames As String

        If chk_selectAll.Checked = True Then
            For i = 0 To chk_list.Items.Count - 1
                chk_list.Items(i).Selected = True
            Next
            For ii = 0 To chk_brand.Items.Count - 1
                chk_brand.Items(ii).Selected = True
            Next
            selectAll.Checked = True
            SelectAllBrand.Checked = True
        Else
            For i = 0 To chk_list.Items.Count - 1
                chk_list.Items(i).Selected = False
            Next
            For ii = 0 To chk_brand.Items.Count - 1
                chk_brand.Items(ii).Selected = False
            Next
            selectAll.Checked = False
            SelectAllBrand.Checked = False
        End If


        Try
            Dim countStores As Integer = 0
            Dim countBrands As Integer = 0

            For i = 0 To chk_list.Items.Count - 1
                If chk_list.Items(i).Selected = True Then
                    storenames &= ",'" & Replace(chk_list.Items(i).Text, "'", "''") & "'"
                End If
            Next
            Dim selectedStores As String

            For ii = 0 To chk_brand.Items.Count - 1
                If chk_brand.Items(ii).Selected = True Then
                    brandnames &= ",'" & Replace(chk_brand.Items(ii).Text, "'", "''") & "'"
                End If
            Next

            Dim selectedbrands As String

            selectedStores = storenames.Remove(0, 1)
            selectedbrands = brandnames.Remove(0, 1)

            selected_holder.Text = Replace(selectedStores, "'", "")
            selected_holder_brands.Text = Replace(selectedbrands, "'", "")
        Catch ex As Exception
            selected_holder.Text = ""
            selected_holder_brands.Text = ""
        End Try


    End Sub

    Private Sub homebtn_Click(sender As Object, e As EventArgs) Handles homebtn.Click
        Response.Redirect("MR.aspx")
    End Sub

    Private Function CreateLog() As Boolean
        Dim connstr As String = [TO].My.Settings.SQLConnection.ToString
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "insert into [TOMR].[dbo].[TO_MR_Log2] ([user name], [date], [doc no], [action]) values ('" & Session("tmpuser") & "', GetDate(), '', 'logout') "

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()
        Return True
    End Function

    Private Sub logout_btn_Click(sender As Object, e As EventArgs) Handles logout_btn.Click
        Dim connstr As String = [TO].My.Settings.SQLConnection.ToString
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "update TO_MR_OpenDoc SET [Active] = 0 WHERE [Current User] = '" & Session.Item("tmpuser") & "' "

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        CreateLog()
        Page.Response.Redirect("Default.aspx")
    End Sub


    Private Sub search_btn_Click(sender As Object, e As EventArgs) Handles search_btn.Click
        Dim cmd As String
        Dim ds As New DataSet
        Dim storenames As String
        Dim brandnames As String
        Dim countStores As Integer = 0
        Dim countBrands As Integer = 0

        Try

            For i = 0 To chk_list.Items.Count - 1
                If chk_list.Items(i).Selected = True Then
                    storenames &= ",'" & Replace(chk_list.Items(i).Text, "'", "''") & "'"
                    countStores = Val(countStores) + 1
                End If
            Next
            Dim selectedStores As String

            For ii = 0 To chk_brand.Items.Count - 1
                If chk_brand.Items(ii).Selected = True Then
                    brandnames &= ",'" & Replace(chk_brand.Items(ii).Text, "'", "''") & "'"
                    countBrands = Val(countBrands) + 1
                End If
            Next
            Dim selectedbrands As String

            'storeCollections.Text = storenames.Remove(0, 1)


            If countStores = 0 Or countBrands = 0 Then
                ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('Stores and brands is required.');", True)
            Else

                selectedStores = storenames.Remove(0, 1)
                selectedbrands = brandnames.Remove(0, 1)



                If chk_withZero.Checked = False Then
                    cmd = "select distinct
                            [DOCUMENT NO]
                            ,[Location Name] as [LOCATION NAME]
                            ,format([Date],'MM/dd/yyyy') as [DATE]
                            ,[STORE NAME]
                            ,[Item No] as [ITEM NO]
                            ,[Description] as [DESCRIPTION]
                            ,[Brand] as [BRAND]
                            --,[REQUESTED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY]
                                   ELSE [REQUESTED QTY]
                                   END as [REQUESTED QTY2]
                            ,[ALLOCATED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY] - [ALLOCATED QTY]
                                   ELSE [UNSERVED QTY]
                                   END as [UNSERVED QTY]
                            --,[UNSERVED QTY]
                            from [TOMR].[dbo].[AllocatedQty_Reports] where [STORE NAME] in (" & selectedStores & ") and [BRAND] in (" & selectedbrands & ") and [Date] >= '" & date_from.Text & "' and [Date] <= '" & date_to.Text & "' --order by [STORE NAME]   
                            "

                Else
                    cmd = "select distinct
                            [DOCUMENT NO]
                            ,[Location Name] as [LOCATION NAME]
                            ,format([Date],'MM/dd/yyyy') as [DATE]
                            ,[STORE NAME]
                            ,[Item No] as [ITEM NO]
                            ,[Description] as [DESCRIPTION]
                            ,[Brand] as [BRAND]
                            --,[REQUESTED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY]
                                   ELSE [REQUESTED QTY]
                                   END as [REQUESTED QTY]
                            ,[ALLOCATED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY] - [ALLOCATED QTY]
                                   ELSE [UNSERVED QTY]
                                   END as [UNSERVED QTY]
                            --,[UNSERVED QTY]
                            from [TOMR].[dbo].[AllocatedQty_Reports] where [UNSERVED QTy] <> 0 and [UNSERVED QTY] > 0 and [REQUESTED QTY] <> 0 and [STORE NAME] in (" & selectedStores & ") and [BRAND] in (" & selectedbrands & ") and [Date] >= '" & date_from.Text & "' and [Date] <= '" & date_to.Text & "' --order by [STORE NAME]  
                            "
                End If

                ds = executeQuery(cmd)

                dg_result.DataSource = ds.Tables(0)
                dg_result.DataBind()

                countStores = 0
                countBrands = 0

            End If


        Catch ex As Exception
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('" & ex.Message & "');", True)
        End Try
    End Sub


    Private Sub Exportbtn_Click(sender As Object, e As EventArgs) Handles Exportbtn.Click
        Dim cmd_updater As String
        Dim ds_updater As New DataSet

        Dim storenames As String
        Dim brandnames As String

        Try

            Dim cmd_series As String
            Dim ds_series As DataSet
            Dim series_ As Integer
            cmd_series = "select top 1 [download_series] from Downloading_Series order by [ID] desc"
            ds_series = executeQuery(cmd_series)
            series_ = ds_series.Tables(0).Rows(0)("download_series")

            Dim increment_series As Integer
            increment_series = Val(series_) + 1
            cmd_series = "update Downloading_Series set [download_series] = '" & increment_series & "'"
            executeQuery(cmd_series)


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
            cmd_columns = <code>
                              <![CDATA[
                                select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS  where TABLE_NAME = 'AllocatedQty_Reports'
                              ]]>
                          </code>.Value
            ds_columns = executeQuery(cmd_columns)

            '// populate header columns to excel
            Dim x As Integer = 1
            For i = 0 To ds_columns.Tables(0).Rows.Count - 1
                xlsWorkSheet.Cells(1, x) = ds_columns.Tables(0).Rows(i)("COLUMN_NAME").ToString.ToUpper
                x = x + 1
            Next

            For i = 0 To chk_list.Items.Count - 1
                If chk_list.Items(i).Selected = True Then
                    storenames &= ",'" & Replace(chk_list.Items(i).Text, "'", "''") & "'"
                End If
            Next
            Dim selectedStores As String

            For ii = 0 To chk_brand.Items.Count - 1
                If chk_brand.Items(ii).Selected = True Then
                    brandnames &= ",'" & Replace(chk_brand.Items(ii).Text, "'", "''") & "'"
                End If
            Next
            Dim selectedbrands As String

            selectedStores = storenames.Remove(0, 1)
            selectedbrands = brandnames.Remove(0, 1)

            Dim cmd_data As String
            Dim ds_data As New DataSet
            If chk_withZero.Checked = False Then
                cmd_data = "select distinct
                            [DOCUMENT NO]
                            ,[Location Name] as [LOCATION NAME]
                            ,format([Date],'MM/dd/yyyy') as [DATE]
                            ,[STORE NAME]
                            ,[Item No] as [ITEM NO]
                            ,[Description] as [DESCRIPTION]
                            ,[Brand] as [BRAND]
                            --,[REQUESTED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY]
                                   ELSE [REQUESTED QTY]
                                   END as [REQUESTED QTY]
                            ,[ALLOCATED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY] - [ALLOCATED QTY]
                                   ELSE [UNSERVED QTY]
                                   END as [UNSERVED QTY]
                            --,[UNSERVED QTY]
                            from [TOMR].[dbo].[AllocatedQty_Reports] where [STORE NAME] in (" & selectedStores & ") and [BRAND] in (" & selectedbrands & ") and [Date] >= '" & date_from.Text & "' and [Date] <= '" & date_to.Text & "' order by [STORE NAME]"
            Else
                cmd_data = "select distinct
                            [DOCUMENT NO]
                            ,[Location Name] as [LOCATION NAME]
                            ,format([Date],'MM/dd/yyyy') as [DATE]
                            ,[STORE NAME]
                            ,[Item No] as [ITEM NO]
                            ,[Description] as [DESCRIPTION]
                            ,[Brand] as [BRAND]
                            --,[REQUESTED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY]
                                   ELSE [REQUESTED QTY]
                                   END as [REQUESTED QTY]
                            ,[ALLOCATED QTY]
                            , case when [REQUESTED QTY] < [ALLOCATED QTY] and [REQUESTED QTY] = 0
                                   then [ALLOCATED QTY] - [ALLOCATED QTY]
                                   ELSE [UNSERVED QTY]
                                   END as [UNSERVED QTY]
                            --,[UNSERVED QTY]
                            from [TOMR].[dbo].[AllocatedQty_Reports] where [UNSERVED QTy] <> 0 and [UNSERVED QTY] > 0 and [REQUESTED QTY] <> 0 and [STORE NAME] in (" & selectedStores & ") and [BRAND] in (" & selectedbrands & ") and [Date] >= '" & date_from.Text & "' and [Date] <= '" & date_to.Text & "' order by [STORE NAME]"
            End If
            ds_data = executeQuery(cmd_data)

            Dim rowstart As Integer = 2
            Dim alphabet_increment As String
            Dim zplus As Integer = 0

            '// bold header
            xlsWorkSheet.Range("a1:az1").Font.Bold = True

            For r = 0 To ds_data.Tables(0).Rows.Count - 1 '// for rows
                For c = 0 To ds_columns.Tables(0).Rows.Count - 1 '// for columns
                    If c > 25 And c < 52 Then
                        alphabet_increment = "a" & Chr(Asc("a") + Val(zplus))
                        zplus = zplus + 1
                    ElseIf c > 51 And c < 78 Then
                        alphabet_increment = "b" & Chr(Asc("a") + Val(zplus))
                        zplus = zplus + 1
                    ElseIf c > 77 And c < 104 Then
                        alphabet_increment = "c" & Chr(Asc("a") + Val(zplus))
                        zplus = zplus + 1
                    ElseIf c > 103 And c < 130 Then
                        alphabet_increment = "d" & Chr(Asc("a") + Val(zplus))
                        zplus = zplus + 1
                    Else
                        alphabet_increment = Chr(Asc("a") + Val(c))
                    End If
                    xlsWorkSheet.Range("a" & r + 2 & ":" & alphabet_increment & r + 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous '// create borders
                    xlsWorkSheet.Range(alphabet_increment & c + 1).ColumnWidth = Len(ds_columns.Tables(0).Rows(c)("COLUMN_NAME")) + 5 '// adjust the width of columns
                    xlsWorkSheet.Cells(rowstart, c + 1) = ds_data.Tables(0).Rows(r)(ds_columns.Tables(0).Rows(c)("COLUMN_NAME")) '// populate all data

                Next
                rowstart = rowstart + 1
                zplus = 0
            Next

            '// save file
            Dim xfile As String
            xfile = Server.MapPath("~\ExcelExport") & "\" & "REPORTS" & "-" & Format(Date.Now, "MMdd") & "-" & series_ & "-" & "(" & Replace(Replace(Session("tmpuser").ToString.ToLower, "primergrp\", ""), ".", " ") & ")" & ".xlsx"
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
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & "REPORTS" & ".xlsx")
            Response.TransmitFile(Server.MapPath("~/ExcelExport/" & "REPORTS" & "-" & Format(Date.Now, "MMdd") & "-" & series_ & "-" & "(" & Replace(Replace(Session("tmpuser").ToString.ToLower, "primergrp\", ""), ".", " ") & ")" & ".xlsx"))
            Response.End()
        Catch ex As Exception
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
            Exit Try
        End Try
    End Sub

    Private Sub selectAll_CheckedChanged(sender As Object, e As EventArgs) Handles selectAll.CheckedChanged

        If selectAll.Checked = True Then
            For i = 0 To chk_list.Items.Count - 1
                chk_list.Items(i).Selected = True
            Next
            allStore()
        Else
            For i = 0 To chk_list.Items.Count - 1
                chk_list.Items(i).Selected = False
            Next
            selected_holder.Text = ""
        End If


    End Sub
    Protected Sub allStore()
        Dim storenames As String
        Dim countStores As Integer = 0
        Dim countBrands As Integer = 0

        selected_holder.Text = ""

        For i = 0 To chk_list.Items.Count - 1
            If chk_list.Items(i).Selected = True Then
                storenames &= ",'" & Replace(chk_list.Items(i).Text, "'", "''") & "'"
            End If
        Next
        Dim selectedStores As String
        selectedStores = storenames.Remove(0, 1)

        selected_holder.Text = Replace(Replace(selectedStores, "'", ""), ",", ", ")
    End Sub

    Private Sub SelectAllBrand_CheckedChanged(sender As Object, e As EventArgs) Handles SelectAllBrand.CheckedChanged

        If SelectAllBrand.Checked = True Then
            For ii = 0 To chk_brand.Items.Count - 1
                chk_brand.Items(ii).Selected = True
            Next
            allbrands()
        Else
            For ii = 0 To chk_brand.Items.Count - 1
                chk_brand.Items(ii).Selected = False
            Next
            selected_holder_brands.Text = ""
        End If
    End Sub
    Protected Sub allbrands()
        Dim brandnames As String
        For ii = 0 To chk_brand.Items.Count - 1
            If chk_brand.Items(ii).Selected = True Then
                brandnames &= ",'" & Replace(chk_brand.Items(ii).Text, "'", "''") & "'"
            End If
        Next
        Dim selectedbrands As String
        selectedbrands = brandnames.Remove(0, 1)

        selected_holder_brands.Text = Replace(Replace(selectedbrands, "'", ""), ",", ", ")
    End Sub

    Private Sub chk_list_SelectedIndexChanged(sender As Object, e As EventArgs) Handles chk_list.SelectedIndexChanged
        Try
            Dim storenames As String
            Dim countStores As Integer = 0
            Dim countBrands As Integer = 0

            selected_holder.Text = ""

            For i = 0 To chk_list.Items.Count - 1
                If chk_list.Items(i).Selected = True Then
                    storenames &= ",'" & Replace(chk_list.Items(i).Text, "'", "''") & "'"
                End If
            Next
            Dim selectedStores As String
            selectedStores = storenames.Remove(0, 1)

            selected_holder.Text = Replace(Replace(selectedStores, "'", ""), ",", ", ")
        Catch ex As Exception
            selected_holder.Text = ""
        End Try
    End Sub

    Private Sub chk_brand_SelectedIndexChanged(sender As Object, e As EventArgs) Handles chk_brand.SelectedIndexChanged
        Try
            Dim brandnames As String
            For ii = 0 To chk_brand.Items.Count - 1
                If chk_brand.Items(ii).Selected = True Then
                    brandnames &= ",'" & Replace(chk_brand.Items(ii).Text, "'", "''") & "'"
                End If
            Next
            Dim selectedbrands As String
            selectedbrands = brandnames.Remove(0, 1)

            selected_holder_brands.Text = Replace(Replace(selectedbrands, "'", ""), ",", ", ")
        Catch ex As Exception
            selected_holder_brands.Text = ""
        End Try
    End Sub

    Private Sub selectAllSiteConcept_CheckedChanged(sender As Object, e As EventArgs) Handles selectAllSiteConcept.CheckedChanged
        If selectAllSiteConcept.Checked = True Then
            For ii = 0 To chk_siteConcept.Items.Count - 1
                chk_siteConcept.Items(ii).Selected = True
            Next
        Else
            For ii = 0 To chk_siteConcept.Items.Count - 1
                chk_siteConcept.Items(ii).Selected = False
            Next
        End If

        Try
            Dim siteconcept As String
            For i = 0 To chk_siteConcept.Items.Count - 1
                If chk_siteConcept.Items(i).Selected = True Then
                    siteconcept &= ",'" & Replace(chk_siteConcept.Items(i).Text, "'", "''") & "'"
                End If
            Next
            Dim selected_siteconcept As String
            selected_siteconcept = siteconcept.Remove(0, 1)
            selectAll.Checked = False
            SelectAllBrand.Checked = False
            selectedSiteConcept_holder.Text = Replace(Replace(selected_siteconcept, "'", ""), ",", ", ")
            ListOfSites(selected_siteconcept)
            ListOfBrands(selected_siteconcept)
        Catch ex As Exception
            selectedSiteConcept_holder.Text = ""
            chk_list.Items.Clear()
            chk_brand.Items.Clear()
        End Try
    End Sub

    Private Sub chk_siteConcept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles chk_siteConcept.SelectedIndexChanged
        Try

            selectedSiteConcept_holder.Text = ""

            Dim siteconcept As String
            For i = 0 To chk_siteConcept.Items.Count - 1
                If chk_siteConcept.Items(i).Selected = True Then
                    siteconcept &= ",'" & Replace(chk_siteConcept.Items(i).Text, "'", "''") & "'"
                End If
            Next
            Dim selected_siteconcept As String = ""
            selected_siteconcept = siteconcept.Remove(0, 1)

            selectedSiteConcept_holder.Text = Replace(Replace(selected_siteconcept, "'", ""), ",", ", ")
            selectAll.Checked = False
            SelectAllBrand.Checked = False
            ListOfSites(selected_siteconcept)
            ListOfBrands(selected_siteconcept)
        Catch ex As Exception
            selectedSiteConcept_holder.Text = ""
            chk_list.Items.Clear()
            chk_brand.Items.Clear()
        End Try
    End Sub

    Protected Sub ListOfSites(selected_site_concepts As String)
        Dim cmd As String
        Dim ds As New DataSet

        cmd = "select distinct [Location Name] from [reports].[dbo].[t_locations4] 
                    where [Company Ownership Code] In (Select distinct [company] from [TOMR].[dbo].[TO_MR_User_Security] where [user name] = '" & Session("tmpuser") & "')
                    And [Type] in ('KIOSK','FREE STANDING STORE') 
                    And [Location Name] Not Like '%CLOSE%' and [Location Name] not like '%DO NOT USE%'
                    and [SiteConcept] in (" & selected_site_concepts & ")"
        ds = executeQuery(cmd)

        chk_list.Items.Clear()

        For i = 0 To ds.Tables(0).Rows.Count - 1
            chk_list.Items.Add(ds.Tables(0).Rows(i)("Location Name"))
        Next
    End Sub

    Protected Sub ListOfBrands(selected_brands As String)
        Dim cmd As String
        Dim ds As New DataSet

        'cmd = "SELECT distinct [Brand Description]
        '        FROM [Reports].[dbo].[t_item_master]
        '        where [Company Ownership Code] in (select [Company Ownership Code] from [Reports].[dbo].[t_locations4] where [SiteConcept] In (" & selected_brands & ") and [Type] = 'FREE STANDING STORE')
        '        order by [Brand Description]"

        cmd = "select distinct [brand description] from [Reports].[dbo].[t_report_004] where [site concept] in (" & selected_brands & ") and [location name] not like '%CLOSED%' order by [brand description]"

        ds = executeQuery(cmd)

        chk_brand.Items.Clear()

        For i = 0 To ds.Tables(0).Rows.Count - 1
            chk_brand.Items.Add(ds.Tables(0).Rows(i)("Brand Description"))
        Next
    End Sub

    Protected Sub OnRowDataBound(sender As Object, e As GridViewRowEventArgs)

        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Attributes("ondblclick") = Page.ClientScript.GetPostBackClientHyperlink(dg_result, "Delete$" & e.Row.RowIndex)

            e.Row.ToolTip = "Double click to view the deleted items for the selected document."
            e.Row.Attributes("style") = "cursor:pointer"
        End If


    End Sub


    Protected Sub dg_result_SelectedIndexChanged(sender As Object, e As GridViewEditEventArgs) Handles dg_result.SelectedIndexChanged
        'For Each row As GridViewRow In dg_result.Rows
        '    If row.RowIndex = dg_result.SelectedIndex Then
        '        row.BackColor = ColorTranslator.FromHtml("#A1DCF2")
        '        row.ToolTip = String.Empty
        '    Else
        '        row.BackColor = ColorTranslator.FromHtml("#FFFFFF")
        '        row.ToolTip = "Click to select this row."
        '    End If
        'Next
    End Sub

    Protected Sub dg_result_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles dg_result.RowEditing
        '
    End Sub

    Protected Sub dg_result_RowDeleting(sender As Object, e As GridViewDeleteEventArgs) Handles dg_result.RowDeleting
        Dim cmd As String
        Dim ds As New DataSet
        backgrounddiv.Visible = True
        deleted_items_div.Visible = True

        cmd = "select distinct b.[Location Name]
                    ,a.[ITEMNO]
                    ,a.[ITEMDESC]
                    ,a.[REQQTY]
                    ,a.[BRAND]
                    --,a.[STATUS]
                    , CASE WHEN a.[STATUS] = 'C' THEN 'CANCELLED' END as [STATUS]               
                    from [dbo].[DocDetail] as a
					inner join [reports].[dbo].[t_locations4] as b
					on a.[DOCNO] like '%' + b.[Location code] + '%'
                    where [DOCNO] in (select [Doc No] from [dbo].[TO_MR_Collated_Header_2] 
                    where [Document No] = '" & dg_result.Rows(e.RowIndex).Cells(0).Text & "')
                    and [brand] = '" & dg_result.Rows(e.RowIndex).Cells(6).Text & "'
                    and [status] = 'c'"
        ds = executeQuery(cmd)
        deletedItems.DataSource = ds.Tables(0)
        deletedItems.DataBind()

        datacounter.Text = "Deleted item/s count: " & ds.Tables(0).Rows.Count

        selectedDocument.Text = "<strong>" & dg_result.Rows(e.RowIndex).Cells(0).Text & "</strong>"
    End Sub
End Class
