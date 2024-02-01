Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class Upload
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mainnav As Control = Master.FindControl("main_navigation")
        mainnav.Visible = False
        If CType(Session.Item("tmpuser"), String) = "" Then
            Page.Response.Redirect("Default.aspx")
        End If

        activeUser.Text = "User: " & "<a style='color:green;'>" & Session("tmpuser") & "</a>"

        con = New SqlConnection(My.Settings.SQLConnection)
        con.Open()

        selectedDocNo.Text = Session("selectedDoc")

        Dim cmd As String
        Dim ds As New DataSet

        cmd = "insert into [dbo].[TO_MR_OpenDoc] ([Document No], [Current User],[Date Time],[active]) values ('" & Session("selectedDoc") & "','" & Session.Item("tmpuser") & "','" & Date.Now & "','1') "
        executeQuery(cmd)

        cmd = "select * from " & Replace(Session("selectedDoc"), "-", "") & ""
        ds = executeQuery(cmd)
        DA = New SqlDataAdapter(cmd, con)
        Dim DT As New DataTable
        DA.Fill(DT)
        uploadgrid.DataSource = DT
        uploadgrid.DataBind()
        DT.Reset()
        con.Close()

    End Sub

    Private Sub btnUpload_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        Dim cmd_series As String
        Dim ds_series As DataSet
        Dim series_ As Integer
        Dim cmd_checker As String
        Dim ds_checker As New DataSet

        cmd_checker = "select * from t_Uploaded where [DocNo] = '" & Session("selectedDoc") & "'"
        ds_checker = executeQuery(cmd_checker)
        If ds_checker.Tables(0).Rows.Count > 0 Then
            cmd_checker = "delete from t_Uploaded where [DocNo] = '" & Session("selectedDoc") & "'"
            executeQuery(cmd_checker)
        End If

        cmd_checker = "select top 1 * from t_Uploaded_Document where [DocNo] = '" & Session("selectedDoc") & "' and [dateuploaded] = '" & Date.Now & "' and [uploadedby] <> '" & Session("tmpuser") & "' order by [ID] desc"
        ds_checker = executeQuery(cmd_checker)
        If ds_checker.Tables(0).Rows.Count > 0 Then
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('This document is already uploaded by " & Replace(ds_checker.Tables(0).Rows(0)("uploadedby").ToString.ToLower, "primergrp\", "") & ". If you have any changes made download the document again for the updated records.');", True)
        Else

            cmd_series = "select [download_series] from Downloading_Series"
            ds_series = executeQuery(cmd_series)
            series_ = ds_series.Tables(0).Rows(0)("download_series")

            Dim increment_series As Integer
            increment_series = Val(series_) + 1
            cmd_series = "update Downloading_Series set [download_series] = '" & increment_series & "'"
            executeQuery(cmd_series)

            Dim cmd As String
            Dim ds As New DataSet

            Dim xlapp As Excel.Application
            Dim xlworkbook As Excel.Workbook
            Dim wsheet As Excel.Worksheet
            Dim x, y, z, rowno As Int64
            Dim xfile As String

            Dim docNo_length As Integer
            docNo_length = Len(Replace(Session("selectedDoc"), "-", ""))

            Dim uniquename, fExt As String
            uniquename = Session("selectedDoc") & "-" & Format(Date.Now, "MMdd") & "-" & series_ & "-" & "(" & Replace(Replace(Session("tmpuser"), "PRIMERGRP\", ""), ".", " ") & ")"
            fExt = System.IO.Path.GetExtension(fileupload1.PostedFile.FileName)
            xfile = Server.MapPath("~\Files")
            fileupload1.SaveAs(xfile & "\" & uniquename & fExt)

            Try
                xlapp = New Excel.Application
                xlapp.DisplayAlerts = False
                xlapp.AskToUpdateLinks = False
                xlworkbook = xlapp.Workbooks.Open(xfile & "\" & uniquename & fExt)
                Dim mysheetCount As Integer = xlworkbook.Worksheets.Count

                'breaklinks
                Dim links As Array = DirectCast(xlworkbook.LinkSources(Excel.XlLink.xlExcelLinks), Array)
                If links IsNot Nothing Then
                    For i As Integer = 1 To links.Length
                        xlworkbook.BreakLink(DirectCast(links.GetValue(i), String), Excel.XlLinkType.xlLinkTypeExcelLinks)
                    Next
                End If
                wsheet = xlworkbook.Worksheets(1)


                If wsheet.Cells(2, 2).value.ToString <> Session("selectedDoc") Then
                    ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('Failed to upload. Invalid File! Please check the document number of the file that you are trying to upload.');", True)
                    Exit Try
                Else
                    Try
                        x = 2
                        y = 14
                        z = 1
                        Dim qty_value As Integer
                        Dim sku As String
                        Do While wsheet.Cells(x, 1).value.ToString <> Nothing
                            Do While wsheet.Cells(1, y).value <> Nothing
                                If wsheet.Cells(x, y).value = Nothing Or wsheet.Cells(x, y).value = 0 Then
                                    qty_value = 0
                                Else
                                    qty_value = wsheet.Cells(x, y).value
                                End If

                                If Len(Trim(wsheet.Cells(x, 3).value)) < 7 Then
                                    sku = "0" & Trim(wsheet.Cells(x, 3).value)
                                Else
                                    sku = Trim(wsheet.Cells(x, 3).value)
                                End If
                                cmd = "INSERT INTO t_Uploaded ([DocNo],[ItemNo],[StoreName],[Qty]) values ('" & Session("selectedDoc") & "','" & sku & "','" & wsheet.Cells(z, y).value & "','" & qty_value & "')"
                                ds = executeQuery(cmd)
                                y = y + 1
                            Loop
                            x = x + 1
                            y = 14
                        Loop
                        x = 2
                    Catch ex As Exception
                        'do nothing
                    End Try

                    Dim cmd_ As String
                    Dim ds_ As New DataSet

                    cmd_ = "select [WHSE QTY] as [qty_],[ITEM NO] as [itemnumber] from " & Replace(Session("selectedDoc"), "-", "") & ""
                    ds_ = executeQuery(cmd_)
                    For a = 0 To ds_.Tables(0).Rows.Count - 1
                        cmd = "select * from t_Uploaded where [DocNo] = '" & Session("selectedDoc") & "' and [ItemNo] = '" & ds_.Tables(0).Rows(a)("itemnumber") & "'"
                        ds = executeQuery(cmd)
                        Dim totalreq As Integer = 0
                        Dim qtyavail As Integer = 0
                        Dim itemcode As String
                        For i = 0 To ds.Tables(0).Rows.Count - 1
                            totalreq += ds.Tables(0).Rows(i)("Qty")
                            qtyavail = ds_.Tables(0).Rows(a)("qty_") - totalreq
                            itemcode = ds.Tables(0).Rows(i)("ItemNo")

                            cmd = "update [dbo].[" & Replace(Session("selectedDoc"), "-", "") & "] set [" & ds.Tables(0).Rows(i)("StoreName").ToString & "] = '" & ds.Tables(0).Rows(i)("Qty") & "'
                               ,[TOTAL REQ] = '" & totalreq & "'
                               ,[QTY AVAIL] = '" & qtyavail & "'
                               where [DOCUMENT NO] = '" & ds.Tables(0).Rows(i)("DocNo") & "'
                               and [ITEM NO] like '%" & itemcode & "'"
                            executeQuery(cmd)
                        Next
                        totalreq = 0
                        qtyavail = 0
                    Next
                    ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('Data successfully uploaded!');", True)
                    cmd = "delete from t_Uploaded where [DocNo] = '" & Session("selectedDoc") & "'"
                    executeQuery(cmd)
                    DocDetailUpdateQty()
                    UpdateDocDetail()
                    xlapp.Quit()

                    cmd = "insert into t_Uploaded_Document ([DocNo],[uploadedby],[dateuploaded]) values ('" & Session("selectedDoc") & "','" & Session("tmpuser") & "','" & Date.Now & "')"
                    executeQuery(cmd)

                    Dim cmd_download As String
                    cmd_download = "insert into TO_MR_Log2 ([user name],[date],[doc no],[action]) values ('" & Session("tmpuser") & "','" & DateAndTime.Now & "','" & Session("selectedDoc") & "','upload')"
                    executeQuery(cmd_download)

                End If
            Catch ex As Exception
                ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('Failed to upload. Invalid File!. Error source: ' + '" & ex.Message.ToString & "' + '.');", True)
                'My.Computer.FileSystem.DeleteFile(xfile & "\" & uniquename & fExt, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin)
            End Try

        End If


    End Sub
    Protected Sub DocDetailUpdateQty()
        'Dim cmd_docdetail_updateQty_select As String
        'Dim ds_docdetail_updateQty_select As New DataSet

        'cmd_docdetail_updateQty_select = <code>
        '                                     <![CDATA[
        '                    DECLARE @Count AS INT
        '                    DECLARE @WMRNo AS VARCHAR(50) = ']]><%= Session("selectedDoc") %><![CDATA['
        '                    DECLARE @Str AS VARCHAR(MAX)

        '                    IF OBJECT_ID('tempdb..#TEMP_TABLENAMES') IS NOT NULL
        '                    DROP TABLE #TEMP_TABLENAMES

        '                    SELECT [COLUMN_NAME] , ROW_NUMBER() OVER(ORDER BY [COLUMN_NAME]) AS Row
        '                    INTO #TEMP_TABLENAMES
        '                    FROM INFORMATION_SCHEMA.COLUMNS
        '                    WHERE TABLE_NAME = replace(@WMRNo,'-','') AND [ORDINAL_POSITION] > 11

        '                    SELECT @Count = COUNT([COLUMN_NAME])
        '                    FROM INFORMATION_SCHEMA.COLUMNS
        '                    WHERE TABLE_NAME = replace(@WMRNo,'-','') AND [ORDINAL_POSITION] > 11

        '                    IF OBJECT_ID('tempdb..#TEMP_GETDATA') IS NOT NULL
        '                    DROP TABLE #TEMP_GETDATA

        '                    IF OBJECT_ID('tempdb..#TEMP_FINAL') IS NOT NULL
        '                    DROP TABLE #TEMP_FINAL

        '                    CREATE TABLE #TEMP_GETDATA (
        '                     ID INT NOT NULL IDENTITY(1,1),
        '                     [WMRNo] VARCHAR(20),
        '                    [ItemNo] VARCHAR(50),
        '                     [Location] VARCHAR(100),
        '                    [Quantity] NUMERIC(18,0)
        '                    )

        '                    WHILE @Count <> 0
        '                    BEGIN

        '                    SET @Str = 'INSERT INTO #TEMP_GETDATA
        '                    SELECT ' + CHAR(39) + replace(@WMRNo,'-','') + CHAR(39) + ', [ITEM NO],' + CHAR(39) + (SELECT DISTINCT [Location Code] 
        '                    FROM [Reports].[dbo].[t_locations2] WHERE [Location Name] IN (SELECT [COLUMN_NAME] FROM #TEMP_TABLENAMES 
        '                    WHERE [Row] = @Count)) + CHAR(39) + ',[' + (SELECT [COLUMN_NAME] FROM #TEMP_TABLENAMES WHERE [Row] = @Count) + '] FROM ' + replace(@WMRNo,'-','')
        '                    EXEC (@Str)
        '                    SET @Count = @Count - 1
        '                    END

        '                    SELECT [WMRNo], [ItemNo], [Location], ISNULL([Quantity], 0) AS [Quantity] , B.[Doc No]
        '                    INTO #TEMP_FINAL
        '                    FROM #TEMP_GETDATA AS A
        '                    INNER JOIN [TO_MR_Collated_Header_2] AS B
        '                    ON A.[Location] = LEFT(B.[Doc No],9)
        '                    AND A.[WMRNo] = REPLACE(B.[Document No],'-','')

        '                    select distinct bb.[Doc No], aa.[ITEMNO], aa.[ITEMDESC], aa.[HOQTY], aa.[REQQTY], aa.[BRAND], aa.[STATUS], bb.[Quantity] as [ALLOCATEDQTY]
        '                    into #TEMP_FINAL_RESULT
        '                    from [TOMR].[dbo].[DocDetail] as aa
        '                    left join #TEMP_FINAL as bb
        '                    on aa.[ITEMNO] = bb.[ItemNo]
        '                    and aa.[DOCNO] = bb.[Doc No]
        '                    where [DOCNO] in (select [Doc No] from [dbo].[TO_MR_Collated_Header_2] where [Document No] = @wmrNo) and [Doc No] <> ''

        '                    select * from #TEMP_FINAL_RESULT

        '                    drop table #TEMP_FINAL_RESULT
        '                    ]]></code>.Value
        'ds_docdetail_updateQty_select = executeQuery(cmd_docdetail_updateQty_select)
        'If ds_docdetail_updateQty_select.Tables(0).Rows.Count > 0 Then
        '    Dim cmd_docdetail_updateQty As String
        '    Dim ds_docdetail_updateQty As New DataSet
        '    For i = 0 To ds_docdetail_updateQty_select.Tables(0).Rows.Count - 1
        '        cmd_docdetail_updateQty = "update [dbo].[DocDetail] set [ALLOCATEDQTY] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ALLOCATEDQTY") & "'
        '                                   where [DOCNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("Doc No") & "'
        '                                   and [ITEMNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ITEMNO") & "'"
        '        ds_docdetail_updateQty = executeQuery(cmd_docdetail_updateQty)
        '    Next
        'End If

        Dim cmd_docdetail_updateQty_select As String
        Dim ds_docdetail_updateQty_select As New DataSet

        cmd_docdetail_updateQty_select = "EXEC sp_selectDocDetail 0,'" & Session("selectedDoc") & "',''"
        ds_docdetail_updateQty_select = executeQuery(cmd_docdetail_updateQty_select)

        If ds_docdetail_updateQty_select.Tables(0).Rows.Count > 0 Then
            Dim cmd_docdetail_updateQty As String
            Dim ds_docdetail_updateQty As New DataSet
            For i = 0 To ds_docdetail_updateQty_select.Tables(0).Rows.Count - 1
                cmd_docdetail_updateQty = "update [dbo].[DocDetail] set [ALLOCATEDQTY] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ALLOCATEDQTY") & "'
                                            where [DOCNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("Doc No") & "'
                                           and [ITEMNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ITEMNO") & "'"
                ds_docdetail_updateQty = executeQuery(cmd_docdetail_updateQty)
            Next
        End If

    End Sub
    Protected Sub UpdateDocDetail()
        Dim cmddocdetail_updater As String
        Dim cmddocheader_updater As String
        Dim ds As New DataSet

        cmddocdetail_updater = <code>
                                   <![CDATA[
                       update DocDetail set [STATUS] = 'C',[ALLOCATEDQTY] = '0'
                       where [docno] in (select [Doc No] from [dbo].[TO_MR_Collated_Header_2] where [Document No] = ']]><%= Session("selectedDoc") %><![CDATA[')
                       and [ITEMNO] in (select [ITEM NO] from [dbo].[]]><%= Replace(Session("selectedDoc"), "-", "") %><![CDATA[] where [TOTAL REQ] = 0) and [Brand] = ']]><%= Session("tmpbrand") %><![CDATA['
                       ]]>
                               </code>.Value
        executeQuery(cmddocdetail_updater)

        cmddocheader_updater = "SELECT sum([TOTAL REQ]) as [totalsum] FROM [dbo].[" & Replace(Session("selectedDoc"), "-", "") & "]"
        ds = executeQuery(cmddocheader_updater)
        If ds.Tables(0).Rows(0)("totalsum").ToString = 0 Then
            'cmddocheader_updater = "UPDATE DocHeader SET STATUS = 'C' WHERE [DOCNO] = '" & Session("selectedDoc") & "'"
            cmddocheader_updater = "UPDATE DocHeader SET STATUS = 'C' WHERE [DOCNO] in (select [Doc NO] from [TO_MR_Collated_Header_2] where [Document No] = '" & Session("selectedDoc") & "')"
            executeQuery(cmddocheader_updater)

            cmddocheader_updater = "UPDATE TO_MR_Collated_Header_2 SET [Status] = 'CANCELLED' where [Document No] ='" & Session("selectedDoc") & "'"
            executeQuery(cmddocheader_updater)
        End If

        'Dim cmd_INC As String
        'Dim ds_INC As New DataSet

        'cmd_INC = <code>
        '              <![CDATA[
        '              	select distinct [DOc No]
        '             from [dbo].[TO_MR_Collated_Header_2]
        '             where [Document No] = ']]><%= Session("selectedDoc") %><![CDATA['
        '             and [status] = 'IN-PROCESS'
        '                ]]>
        '          </code>.Value
        'ds_INC = executeQuery(cmd_INC)
        'Dim cmd_update_INC_status As String
        'Dim ds_update_INC_status As New DataSet
        'If ds_INC.Tables(0).Rows.Count > 0 Then
        '    For i = 0 To ds_INC.Tables(0).Rows.Count - 1
        '        cmd_update_INC_status = "EXEC sp_update_INC_status '','','" & Session("selectedDoc") & "','" & ds_INC.Tables(0).Rows(i)("DOc No") & "',''"
        '        executeQuery(cmd_update_INC_status)
        '    Next
        'End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Response.Redirect("MR.aspx")
    End Sub

    Private Sub homebtn_Click(sender As Object, e As EventArgs) Handles homebtn.Click
        Response.Redirect("MR.aspx")
    End Sub

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
End Class