Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Drawing

Public Class MR
    Inherits System.Web.UI.Page
    Private bfield As New BoundField()
    Private btnfield As New ButtonField()
    Dim connstr As String = [TO].My.Settings.SQLConnection.ToString  '"Server='10.63.1.161';Initial Catalog='TOMR';user id='sa';password='Pa$$w0rd';"

    Dim dsmr As DataSet
    Dim dsloc As DataSet

    Dim tmptext As String
    Dim tmpint As Int16 = 0

    Dim docnomail As String

    Public Shared tmpdocno As String
    Public Shared tmpco As String
    Public Shared tmpbrand As String

    Private Function EndDocumentSession() As Boolean
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "update TO_MR_OpenDoc SET [Active] = 0 WHERE [Document No] = '" & Session("tmpdoc") & "' and [Current User] = '" & Session.Item("tmpuser") & "' "

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()
        Return True
    End Function

    Private Function CreateLog(docno As String) As Boolean
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "insert into [dbo].[TO_MR_Log2] ([user name], [date], [doc no], [action]) values ('" & Session("tmpuser") & "', GetDate(), '" & docno & "', 'textfile') "

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()
        Return True
    End Function

    Private Function GetData(tmpco As String, tmpbrand As String) As DataSet
        Dim tmpco1 As String
        tmpco1 = tmpco
        If tmpco = "" Then
            tmpco1 = "PKCI"
        End If
        Dim sqlscript As String =
        <code>
            <![CDATA[

                declare @CompanyCode as nvarchar(10)
                declare @Brand as nvarchar(25)
                declare @qtyscript nvarchar(MAX)

                set @CompanyCode = ']]><%= tmpco1 %><![CDATA['
                set @Brand = ']]><%= tmpbrand %><![CDATA['


                create table #whseqty
                (
                    [Item No] [nvarchar](20) NULL,
	                [Quantity] [int] NULL,
                )

                set @qtyscript = 
				'
				insert into #whseqty
				select 
					loc1.[Item No_] as [Item No]
					, SUM(loc1.[Quantity]) as [Quantity]
				from [Test51217].[dbo].[' + @CompanyCode + '$Item Ledger Entry] as loc1 with (nolock)
				left outer join [Reports].[dbo].[t_locations4] as loc2 with (nolock)
					on loc1.[Location Code] = loc2.[Location Code] COLLATE DATABASE_DEFAULT
				where loc1.[Location Code] like ''W%'' and loc1.[Item Category Code] = ''OWN'' and loc2.[Location Name] like ''%GOOD%''
				group by loc1.[Item No_]
				'

                exec(@qtyscript)

                create table #tmp1
                (
                    [ID] [int] identity(1,1) NOT NULL, 
                    [Item No] [nvarchar](20) NULL,
	                [Item Description] [nvarchar](100) NULL,
	                [Brand] [nvarchar](50) NULL,
	                [Company Owner] [nvarchar](10) NULL,
	                [Location Code] [nvarchar](25) NULL,
	                [Req Qty] [int] NULL,
					[Whse Qty] [int] NULL,
                )

                insert into #tmp1
                select 
	                a.[Item No]
	                , a.[Item Description]
	                , a.[Brand]
	                , a.[Item Owner]
	                , a.[Location Code]
	                , a.[Request Qty]
					, wqty.[Quantity]
                from [dbo].[TO_MR_Collate] as a with (nolock)

				left outer join #whseqty as wqty
					on a.[Item No] = wqty.[Item No]

                where 
	                a.[Item Owner] in (@CompanyCode)
	                and a.[Brand] in (@Brand)


                DECLARE @cols AS NVARCHAR(MAX),
                    @query  AS NVARCHAR(MAX);

                SET @cols = STUFF((select distinct ',' + QUOTENAME(tmp1.[Location Code]) 
                            FROM #tmp1 as tmp1
                            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '')

                set @query = 
	                '
	                SELECT
						[ID]
		                , [Company Owner]
		                , [Brand] 
		                , [Item No]
		                , [Item Description]
						, [Whse Qty]
						, [Total Req]
		                , ' + @cols + ' 
	                from 
		                (
		                select
							[ID]
			                , [Company Owner]
			                , [Brand]
			                , [Item No]
			                , [Item Description]
			                , [Location Code]
			                , [Req Qty]
							, [Whse Qty]
							, sum([Req Qty]) as [Total Req]
		                from #tmp1
						group by 
							[ID]
			                , [Company Owner]
			                , [Brand]
			                , [Item No]
			                , [Item Description]
			                , [Location Code]
			                , [Req Qty]
							, [Whse Qty]
		                ) x
	                pivot 
		                (
			                sum([Req Qty])
			                for [Location Code] in (' + @cols + ')
		                ) p 
					
	                '

                execute(@query)

				drop table #whseqty
                drop table #tmp1

            ]]>
        </code>.Value

        Dim cmd As New SqlCommand(sqlscript)
        Using con As New SqlConnection(connstr)
            Using sda As New SqlDataAdapter()
                cmd.Connection = con
                sda.SelectCommand = cmd
                Using ds As New DataSet()
                    sda.Fill(ds)
                    Return ds
                End Using
            End Using
            con.Dispose()
        End Using
    End Function

    Private Function GetLocations(tmpco As String, tmpbrand As String) As DataSet
        Dim sqlscript As String =
        <code>
            <![CDATA[

                declare @CompanyCode as nvarchar(10)
                declare @Brand as nvarchar(25)

                set @CompanyCode = ']]><%= tmpco %><![CDATA['
                set @Brand = ']]><%= tmpbrand %><![CDATA['
                
                select 
	                distinct [Location Code]
                from [dbo].[TO_MR_Collate] as a with (nolock)
                where 
	                a.[Item Owner] in (@CompanyCode)
	                and a.[Brand] in (@Brand)

            ]]>
        </code>.Value

        Dim cmd As New SqlCommand(sqlscript)
        Using con As New SqlConnection(connstr)
            Using sda As New SqlDataAdapter()
                cmd.Connection = con
                sda.SelectCommand = cmd
                Using ds As New DataSet()
                    sda.Fill(ds)
                    Return ds
                End Using
            End Using
            con.Dispose()
        End Using
    End Function

    Private Function CreateText1(tmpco As String, tmpdocno As String) As DataSet
        Dim sqlscript As String =
        <code>
            <![CDATA[

                declare @tablename as nvarchar(100) = REPLACE(']]><%= tmpdocno %><![CDATA[', '-', '')


                create table #tmp1
                    (
		                [Document No] [nvarchar](50) NULL,
		                [Item No] [nvarchar](20) NULL,
		                [Item Owner] [nvarchar](10) NULL,
		                [Brand] [nvarchar](50) NULL,
		                [Location Name] [nvarchar](200) NULL,
		                [Request Qty] [int] NULL,
                    )


                declare @cols as nvarchar(max) =
	                STUFF((select ',' + QUOTENAME(a.[COLUMN_NAME])
	                from INFORMATION_SCHEMA.COLUMNS as a
	                left outer join
		                ( select distinct
			                [Location Name]
		                from [Reports].[dbo].[t_locations2] with (nolock)
		                ) as loc
		                on a.[COLUMN_NAME] = loc.[Location Name]
	                where TABLE_NAME = @tablename
		                and loc.[Location Name] is not null
		                for xml path(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '')


                declare @unpscript as nvarchar(max) =
                '
	                insert into #tmp1
	                select
		                u.[DOCUMENT NO]
		                , u.[ITEM NO]
		                , u.[ITEM OWNER]
		                , u.[BRAND]
		                , u.[STORE]
		                , u.[REQ QTY]
	                from [dbo].[' + @tablename + '] as tmp with (nolock)
	                unpivot
	                (
		                [REQ QTY]
		                for [STORE] in (' + @cols + ')
	                ) as u
                '
                exec(@unpscript)


                select
	                whse.[Warehouse Code] as [From Code]
	                , loc.[Location Code] as [To Code]
	                , tmp.[Location Name]
	                , whseout.[Location Code] as [Transit Code]
	                , CAST(GETDATE() as date) as [Posting Date]
	                , tmp.[Document No]
	                , 'REPLENISHMENT' as [Transfer Type]
	                , 'TO-MR-01' as [TO Reason]
	                , tmp.[Item No]
	                , tmp.[Request Qty]
	                , loc.[Type] as [Location Type]
	                ,
	                CASE
	                WHEN right(substring(ltrim(cast((itm.[Unit Price]) as varchar(25))), 1, charindex('.',ltrim(cast((itm.[Unit Price]) as varchar(25)))) - 1), 2) = '20'
		                THEN 'MARKED DOWN'
	                WHEN right(substring(ltrim(cast((itm.[Unit Price]) as varchar(25))), 1, charindex('.',ltrim(cast((itm.[Unit Price]) as varchar(25)))) - 1), 2) = '70'
		                THEN 'MARKED DOWN'
	                WHEN right(substring(ltrim(cast((itm.[Unit Price]) as varchar(25))), 1, charindex('.',ltrim(cast((itm.[Unit Price]) as varchar(25)))) - 1), 2) = '00'
		                THEN 'MARKED DOWN'
	                ELSE 'REGULAR'
	                END as [Price Type]
	                , ISNULL(rems.[Remarks], '') as [Remarks]
	                , itm.[MainCat Description]
	                , ISNULL(bt.[Barcode Size], '') as [Barcode Size]
                into #tmp2
                from #tmp1 as tmp
                left outer join
	                ( select distinct
		                [Company Code]
		                , [Location Code]
		                , [Location Name]
		                , [Type]
	                from [Reports].[dbo].[t_locations2] with (nolock)
	                ) as loc
	                on tmp.[Item Owner] = loc.[Company Code]
	                and tmp.[Location Name] = loc.[Location Name]
                left outer join [dbo].[TO_MR_Warehouse] as whse with (nolock)
	                on tmp.[Item Owner] = whse.[Company Ownership Code]
                    and tmp.[Brand] = whse.[Brand]
                left outer join 
	                ( select distinct
		                [Location Code]
		                , [Company Code]
	                from [t_locations2] with (nolock)
	                where [Location Code] like 'LOGOUT%'
	                ) as whseout
	                on tmp.[Item Owner] = whseout.[Company Code]
                left outer join [Reports].[dbo].[t_item_master] as itm with (nolock)
	                on tmp.[Item No] = itm.[Item No]
                left outer join
	                ( select distinct
		                [Doc No]
		                , [Remarks]
	                from [dbo].[TO_MR_Collated_Header_2]
	                where REPLACE([Document No], '-', '') = @tablename
	                ) as rems
	                on loc.[Location Code] = LEFT(rems.[Doc No], 9)
                left outer join [dbo].[TO_MR_Barcode_Types] as bt with (nolock)
	                on tmp.[Item Owner] = bt.[Item Owner]
	                and tmp.[Brand] = bt.[Brand]
	                and itm.[MainCat Description] = bt.[Main Category]

                where ([Request Qty] <> '0' and [Request Qty] is not null)

                drop table #tmp1


                select
	                [From Code]
	                , [To Code]
	                , [Transit Code]
	                , [Posting Date]
	                , [Document No]
	                , [Transfer Type]
	                , [TO Reason]
	                , [Item No]
	                , [Request Qty]
	                , 
	                CASE
	                WHEN ([Location Type] = 'FREE STANDING STORE' or [Location Type] = 'KIOSK') THEN [Remarks]
	                ELSE 
		                CASE 
		                WHEN [Location Name] like 'SM%' THEN [Remarks] + ' - ' + [Price Type] + ' - ' + [Barcode Size]
		                ELSE [Remarks] + ' - ' + [Price Type]
		                END
	                END as [Remarks]
                from #tmp2
                where
	
	                ([Location Type] in ('FREE STANDING STORE', 'KIOSK')
	                or ([Location Type] = 'CONCESSION' and [Price Type] = 'REGULAR'))

                    order by [To Code] asc
               -- order by [Price Type] DESC, [To Code], [Item No]

                drop table #tmp2


            ]]>
        </code>.Value

        Dim cmd As New SqlCommand(sqlscript)
        Using con As New SqlConnection(connstr)
            Using sda As New SqlDataAdapter()
                cmd.Connection = con
                sda.SelectCommand = cmd
                Using ds As New DataSet()
                    sda.Fill(ds)
                    Return ds
                End Using
            End Using
            con.Dispose()
        End Using
    End Function

    Private Function CreateText2(tmpco As String, tmpdocno As String) As DataSet
        Dim sqlscript As String =
        <code>
            <![CDATA[

                declare @tablename as nvarchar(100) = REPLACE(']]><%= tmpdocno %><![CDATA[', '-', '')


                create table #tmp1
                    (
		                [Document No] [nvarchar](50) NULL,
		                [Item No] [nvarchar](20) NULL,
		                [Item Owner] [nvarchar](10) NULL,
		                [Brand] [nvarchar](50) NULL,
		                [Location Name] [nvarchar](200) NULL,
		                [Request Qty] [int] NULL,
                    )


                declare @cols as nvarchar(max) =
	                STUFF((select ',' + QUOTENAME(a.[COLUMN_NAME])
	                from INFORMATION_SCHEMA.COLUMNS as a
	                left outer join
		                ( select distinct
			                [Location Name]
		                from [Reports].[dbo].[t_locations2] with (nolock)
		                ) as loc
		                on a.[COLUMN_NAME] = loc.[Location Name]
	                where TABLE_NAME = @tablename
		                and loc.[Location Name] is not null
		                for xml path(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '')


                declare @unpscript as nvarchar(max) =
                '
	                insert into #tmp1
	                select
		                u.[DOCUMENT NO]
		                , u.[ITEM NO]
		                , u.[ITEM OWNER]
		                , u.[BRAND]
		                , u.[STORE]
		                , u.[REQ QTY]
	                from [dbo].[' + @tablename + '] as tmp with (nolock)
	                unpivot
	                (
		                [REQ QTY]
		                for [STORE] in (' + @cols + ')
	                ) as u
                '
                exec(@unpscript)


                select
	                whse.[Warehouse Code] as [From Code]
	                , loc.[Location Code] as [To Code]
	                , tmp.[Location Name]
	                , whseout.[Location Code] as [Transit Code]
	                , CAST(GETDATE() as date) as [Posting Date]
	                , tmp.[Document No]
	                , 'REPLENISHMENT' as [Transfer Type]
	                , 'TO-MR-01' as [TO Reason]
	                , tmp.[Item No]
	                , tmp.[Request Qty]
	                , loc.[Type] as [Location Type]
	                ,
	                CASE
	                WHEN right(substring(ltrim(cast((itm.[Unit Price]) as varchar(25))), 1, charindex('.',ltrim(cast((itm.[Unit Price]) as varchar(25)))) - 1), 2) = '20'
		                THEN 'MARKED DOWN'
	                WHEN right(substring(ltrim(cast((itm.[Unit Price]) as varchar(25))), 1, charindex('.',ltrim(cast((itm.[Unit Price]) as varchar(25)))) - 1), 2) = '70'
		                THEN 'MARKED DOWN'
	                WHEN right(substring(ltrim(cast((itm.[Unit Price]) as varchar(25))), 1, charindex('.',ltrim(cast((itm.[Unit Price]) as varchar(25)))) - 1), 2) = '00'
		                THEN 'MARKED DOWN'
	                ELSE 'REGULAR'
	                END as [Price Type]
	                , ISNULL(rems.[Remarks], '') as [Remarks]
	                , itm.[MainCat Description]
	                , ISNULL(bt.[Barcode Size], '') as [Barcode Size]
                into #tmp2
                from #tmp1 as tmp
                left outer join
	                ( select distinct
		                [Company Code]
		                , [Location Code]
		                , [Location Name]
		                , [Type]
	                from [Reports].[dbo].[t_locations2] with (nolock)
	                ) as loc
	                on tmp.[Item Owner] = loc.[Company Code]
	                and tmp.[Location Name] = loc.[Location Name]
                left outer join [dbo].[TO_MR_Warehouse] as whse with (nolock)
	                on tmp.[Item Owner] = whse.[Company Ownership Code]
                    and tmp.[Brand] = whse.[Brand]
                left outer join 
	                ( select distinct
		                [Location Code]
		                , [Company Code]
	                from [t_locations2] with (nolock)
	                where [Location Code] like 'LOGOUT%'
	                ) as whseout
	                on tmp.[Item Owner] = whseout.[Company Code]
                left outer join [Reports].[dbo].[t_item_master] as itm with (nolock)
	                on tmp.[Item No] = itm.[Item No]
                left outer join
	                ( select distinct
		                [Doc No]
		                , [Remarks]
	                from [dbo].[TO_MR_Collated_Header_2]
	                where REPLACE([Document No], '-', '') = @tablename
	                ) as rems
	                on loc.[Location Code] = LEFT(rems.[Doc No], 9)
                left outer join [dbo].[TO_MR_Barcode_Types] as bt with (nolock)
	                on tmp.[Item Owner] = bt.[Item Owner]
	                and tmp.[Brand] = bt.[Brand]
	                and itm.[MainCat Description] = bt.[Main Category]

                where ([Request Qty] <> '0' and [Request Qty] is not null)

                drop table #tmp1


                select
	                [From Code]
	                , [To Code]
	                , [Transit Code]
	                , [Posting Date]
	                , [Document No]
	                , [Transfer Type]
	                , [TO Reason]
	                , [Item No]
	                , [Request Qty]
	                , 
	                CASE
	                WHEN ([Location Type] = 'FREE STANDING STORE' or [Location Type] = 'KIOSK') THEN [Remarks]
	                ELSE 
		                CASE 
		                WHEN [Location Name] like 'SM%' THEN [Remarks] + ' - ' + [Price Type] + ' - ' + [Barcode Size]
		                ELSE [Remarks] + ' - ' + [Price Type]
		                END
	                END as [Remarks]
                from #tmp2
                where
	
	                ([Location Type] = 'CONCESSION' and [Price Type] = 'MARKED DOWN')

                    order by [To Code] asc
                --order by [Price Type] DESC, [To Code], [Item No]

                drop table #tmp2


            ]]>
        </code>.Value

        Dim cmd As New SqlCommand(sqlscript)
        Using con As New SqlConnection(connstr)
            Using sda As New SqlDataAdapter()
                cmd.Connection = con
                sda.SelectCommand = cmd
                Using ds As New DataSet()
                    sda.Fill(ds)
                    Return ds
                End Using
            End Using
            con.Dispose()
        End Using
    End Function

    Private Function GetHeader(Field As String, Filter As String) As DataSet
        Dim sqlscript As String =
        <code>
            <![CDATA[

                select 
	                [Document No]
                    , [Company]
                    , [Brand]
                    , Format([Created Date],'MM/dd/yyyy') as [Created Date]
                    , [Status]    
                from [dbo].[TO_MR_Collated_Header] as a with (nolock)
                where [Company] in (
					select distinct [Company] 
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)
					and [Brand] in (
					select distinct [Brand]
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)             
                                
            ]]>
        </code>.Value
        Select Case ddlFields.SelectedValue
            Case "Brand"
                sqlscript &= " and [Brand] like '" & dd_status.Text & "' "
            Case "Company"
                sqlscript &= " and [Company] like '" & dd_status.Text & "' "
            Case "Status"
                sqlscript &= " and [Status] like '" & dd_status.Text & "' "
        End Select

        sqlscript &= " order by Status, [Created Date], Company, Brand"

        Dim cmd As New SqlCommand(sqlscript)
        Using con As New SqlConnection(connstr)
            Using sda As New SqlDataAdapter()
                cmd.Connection = con
                sda.SelectCommand = cmd
                Using ds As New DataSet()
                    sda.Fill(ds)
                    Return ds
                End Using
            End Using
            con.Dispose()
        End Using
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Dim mainnav As Control = Master.FindControl("main_navigation")
        mainnav.Visible = False

        Dim cmd As String
        Dim ds As New DataSet

        cmd = "select * from [TOMR].[dbo].[TO_MR_User_Security] where [user name] = '" & Session("tmpuser") & "' and [level] = 'ADMIN'"
        ds = executeQuery(cmd)
        If ds.Tables(0).Rows.Count > 0 Then
            reportsBTN2.Visible = True
            'reportsBTN2.Enabled = True
        Else
            reportsBTN2.Visible = False
            'reportsBTN2.Enabled = False
        End If

        activeUser.Text = "User: <a style='color:green'>" & Session("tmpuser") & "</a>"

        If Session("deletesuccess") = "deleted" Then
            ClientScript.RegisterStartupScript(Me.GetType(), "script", " alert('Document is successfully cancelled.'); ", True)
        End If
        Session("deletesuccess") = ""

        Try
            If CType(Session.Item("tmpuser"), String) = "" Then
                Page.Response.Redirect("Default.aspx")
            End If
            lbl_msgbox.Visible = False

            ds = executeQuery("select * from [TO_MR_OpenDoc] where [Current User] = '" & Session("tmpuser") & "'")
            If ds.Tables(0).Rows.Count > 0 Then
                executeQuery("update [TO_MR_OpenDoc] set [Active] = '0' where [Current User] = '" & Session("tmpuser") & "'")
            End If

            If Not Page.IsPostBack Then
                cmd = "select distinct [status] from [TO_MR_Collated_Header_2] where [status] <> 'DEL' and [status] <> '' order by [status] asc"
                ds = executeQuery(cmd)
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dd_status.Items.Add(ds.Tables(0).Rows(i)("status"))
                Next

                ddlFields.SelectedIndex = 3
                dd_status.SelectedValue = "IN-PROCESS"
                EndDocumentSession()
                CreateGrid()
                BindGrid(ddlFields.SelectedValue, dd_status.Text)
            End If
        Catch ex As Exception
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('" & ex.Message & "');", True)
        End Try
    End Sub

    Private Sub CreateGrid()
        GridView1.Columns.Clear()

        Dim nameColumn As BoundField

        nameColumn = New BoundField()
        nameColumn.DataField = "Document No"
        nameColumn.HeaderText = "DOCUMENT NO."
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Company"
        nameColumn.HeaderText = "COMPANY CODE"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Brand"
        nameColumn.HeaderText = "BRAND"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        'nameColumn = New BoundField()
        'nameColumn.DataField = "TotalRequest"
        'nameColumn.HeaderText = "TOTAL REQ"
        'nameColumn.InsertVisible = False
        'nameColumn.ReadOnly = True
        'GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Created Date"
        nameColumn.HeaderText = "CREATED DATE"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Status"
        nameColumn.HeaderText = "STATUS"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Text File Date"
        nameColumn.HeaderText = "TEXT FILE DATE"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        btnfield = New ButtonField()
        btnfield.ButtonType = ButtonType.Button
        btnfield.HeaderText = ""
        btnfield.CommandName = "viewbtn"
        btnfield.HeaderStyle.Width = "30"
        btnfield.Text = "View / Edit"
        'btnfield.ControlStyle.BackColor = ColorTranslator.FromHtml("#428bca")
        'btnfield.ControlStyle.BorderStyle = BorderStyle.None
        'btnfield.ControlStyle.ForeColor = Color.White
        GridView1.Columns.Add(btnfield)

        Dim ef As CommandField = New CommandField()
        ef.ButtonType = ButtonType.Button
        ef.ShowDeleteButton = True
        ef.DeleteText = "Create NAV Upload File"
        'ef.ControlStyle.BackColor = ColorTranslator.FromHtml("#428bca")
        'ef.ControlStyle.BorderStyle = BorderStyle.None
        'ef.ControlStyle.ForeColor = Color.White
        GridView1.Columns.Add(ef)

        btnfield = New ButtonField()
        btnfield.ButtonType = ButtonType.Button
        btnfield.HeaderText = ""
        btnfield.CommandName = "Downloadbtn"
        btnfield.HeaderStyle.Width = "40"
        btnfield.Text = "Download"
        'btnfield.ControlStyle.BackColor = ColorTranslator.FromHtml("#428bca")
        'btnfield.ControlStyle.BorderStyle = BorderStyle.None
        'btnfield.ControlStyle.ForeColor = Color.White
        GridView1.Columns.Add(btnfield)

        btnfield = New ButtonField()
        btnfield.ButtonType = ButtonType.Button
        btnfield.HeaderText = ""
        btnfield.CommandName = "Uploadbtn"
        btnfield.HeaderStyle.Width = "40"
        btnfield.Text = "Upload"
        'btnfield.ControlStyle.BackColor = ColorTranslator.FromHtml("#428bca")
        'btnfield.ControlStyle.BorderStyle = BorderStyle.None
        'btnfield.ControlStyle.ForeColor = Color.White
        GridView1.Columns.Add(btnfield)


    End Sub

    Private Sub BindGrid(_Field As String, _Filter As String)
        Try
            Dim cmd_filtered_date As String
            Dim ds_filtered_date As DataSet
            Dim filtered_date As String

            cmd_filtered_date = "select top 1 format(filtered_date,'MM/dd/yyyy') as [filtered_date] from FilteredDate_Settings order by [ID] desc"
            ds_filtered_date = executeQuery(cmd_filtered_date)
            filtered_date = ds_filtered_date.Tables(0).Rows(0)("filtered_date").ToString

            Dim sConnectionString As String
            Dim objConn As SqlConnection
            sConnectionString = [TO].My.Settings.SQLConnection.ToString
            objConn = New SqlConnection(sConnectionString)
            objConn.Open()

            Dim sqlscript As String =
            <code>
                <![CDATA[

                select distinct
	                a.[Document No]
                    , a.[Company]
                    , a.[Brand]
                    , Format(a.[Created Date],'MM/dd/yyyy') as [Created Date]
                    , a.[Status]
	                , ISNULL(Format(log2.[date],'MM/dd/yyyy'), '') as [Text File Date]
                from [dbo].[TO_MR_Collated_Header_2] as a with (nolock)

                left outer join
	                ( select
		                [doc no], MIN([date]) as [date]
	                from [dbo].[TO_MR_Log2]
	                where [action] = 'textfile'
					group by [doc no]
	                ) as log2
	                on a.[Document No] = log2.[doc no]

                where [Company] in (
					select distinct [Company] 
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)
					and [Brand] in (
					select distinct [Brand]
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)      
                                
            ]]>
            </code>.Value

            Dim sqlscript2 As String =
            <code>
                <![CDATA[


                DECLARE @CountTotal AS INT = (SELECT COUNT(DISTINCT [Document No]) AS [Count] FROM [TO_MR_Collated_Header_2] where [created date] >= ']]><%= filtered_date %><![CDATA['), @CountCurrent AS INT = 1
                DECLARE @TableName AS VARCHAR(50), @StrQuery AS VARCHAR(MAX) ,@Var1 AS VARCHAR(50)

                IF OBJECT_ID('tempdb..#TEMP1') IS NOT NULL
                DROP TABLE #TEMP1

                CREATE TABLE #TEMP1 (
                ID INT IDENTITY(1,1) NOT NULL,
                TotalRequest INT,
                DocumentNo VARCHAR(50))

                WHILE @CountTotal > @CountCurrent
                BEGIN
                IF OBJECT_ID('tempdb..#TEMP_MASTERDATA') IS NOT NULL
                DROP TABLE #TEMP_MASTERDATA

                SELECT REPLACE([Document No],'-','') AS [DocumentNo] ,ROW_NUMBER() OVER(ORDER BY REPLACE([Document No],'-','') DESC) AS [Row] INTO #TEMP_MASTERDATA FROM [dbo].[TO_MR_Collated_Header_2] WITH(NOLOCK) WHERE [created date] >= ']]><%= filtered_date %><![CDATA[' GROUP BY [Document No]

                SET @TableName = (SELECT [DocumentNo] FROM #TEMP_MASTERDATA WHERE [Row] = @CountCurrent)

                SET @StrQuery = 'SELECT SUM([TOTAL REQ]) AS [TotalRequest], ' + CHAR(39) + @TableName + CHAR(39) + ' AS [DocumentNo] FROM ' + @TableName
                PRINT @StrQuery

                INSERT INTO #TEMP1
                EXECUTE (@StrQuery)
                SET @CountCurrent = @CountCurrent + 1
                END

                IF OBJECT_ID('tempdb..#tmpfinal') IS NOT NULL
                DROP TABLE #tmpfinal

                SELECT DISTINCT
                A.[Document No]
                , A.[Company]
                , A.[Brand]
                , B.[TotalRequest]
                , Format(a.[Created Date],'MM/dd/yyyy') AS [Created Date]
                , A.[Status]
                , ISNULL(Format(log2.[date],'MM/dd/yyyy'), '') AS [Text File Date]
                into #tmpfinal
                FROM [dbo].[TO_MR_Collated_Header_2] AS A WITH(NOLOCK)
                LEFT OUTER JOIN
                (SELECT
                [doc no]
                , MIN([Date]) AS [Date]
                FROM [dbo].[TO_MR_Log2]
                WHERE [action] = 'textfile'
                GROUP BY [doc no]) AS log2
                ON a.[Document No] = log2.[doc no]

                LEFT JOIN #TEMP1 AS B WITH(NOLOCK)
                ON REPLACE(A.[Document No],'-','') = B.[DocumentNo]
                WHERE A.[Created Date] >= ']]><%= filtered_date %><![CDATA['

                select * from #tmpfinal where [created date] >= ']]><%= filtered_date %><![CDATA[' and [Company]
                in (select distinct [company] from [dbo].[TO_MR_User_Security]
                where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
                )
                and
                [Brand] in (
                select distinct [Brand]
                from [dbo].[TO_MR_User_Security]
                where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
                )

   
                                
            ]]>
            </code>.Value


            Select Case _Field
                    Case "Brand"
                        sqlscript &= " and [Brand] like '" & _Filter & "' and [created date] >= '" & filtered_date & "'"
                    Case "Company"
                        sqlscript &= " and [Company] like '" & _Filter & "' and [created date] >= '" & filtered_date & "'"
                    Case "Status"
                        sqlscript &= " and [Status] like '" & _Filter & "' and [created date] >= '" & filtered_date & "'"
                End Select


            'If _Field = "Status" And _Filter = "NAV ON QUEUE" Then
            'sqlscript &= " order by Format(a.[Created Date],'MM/dd/yyyy') desc"
            'Else
            sqlscript &= " order by [Status], [Created Date], [Company], [Brand]"
            'End If

            Dim sqlcmd As New SqlCommand(sqlscript, objConn)
            GridView1.DataSource = sqlcmd.ExecuteReader()
            GridView1.DataBind()

            objConn.Close()
            objConn.Dispose()


            'dsmr = Me.GetHeader(_Field, _Filter)
            'If dsmr.Tables.Count <> 0 Then
            '    GridView1.DataSource = dsmr
            '    GridView1.DataBind()
            'End If
            'div_filter.Attributes.Add("width", GridView1.Width.ToString)
            div_filter.Style.Add("width", GridView1.Width.ToString & "px")
        Catch ex As Exception
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
            Exit Try
        End Try


    End Sub

    Function CheckNeg(tname As String) As Int16
        Dim reccount As Int32
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()
        Dim cmd As String =
        <code>
            <![CDATA[

                    declare @tablename as nvarchar(100) = REPLACE(']]><%= tname %><![CDATA[', '-', '')

                    declare @runsql as nvarchar(max) =
                    '
                    select count(*) as [rcnt]
                    from [dbo].[' + @tablename + ']
                    where [QTY AVAIL] < 0
                    '
                    exec(@runsql)

                    ]]>
        </code>.Value
        sqlcmd = New SqlCommand(cmd, objConn)
        'sqlcmd.ExecuteNonQuery()
        reccount = Convert.ToInt32(sqlcmd.ExecuteScalar())
        objConn.Close()
        objConn.Dispose()


        Return reccount

        'If reccount <> 0 Then
        '    Response.Write("<script language=""javascript"">alert('There are lines with insufficient quantity.');</script>")
        '    'MessageLabel.Text = "There are " + reccount.ToString + " lines with insufficient quantity."
        'Else

        'End If

    End Function

    Protected Sub OnRowDataBound(sender As Object, e As GridViewRowEventArgs)

        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.Cells(4).Text = "NAV ON QUEUE" Then
                e.Row.Cells(7).Enabled = False
                e.Row.Cells(6).Enabled = True
                e.Row.Cells(8).Enabled = True
                e.Row.Cells(9).Enabled = False
                'e.Row.Cells(5).Visible = False
            ElseIf e.Row.Cells(4).Text = "CANCELLED" Then
                e.Row.Cells(7).Enabled = False
                e.Row.Cells(6).Enabled = True
                e.Row.Cells(8).Enabled = True
                e.Row.Cells(9).Enabled = False
                'e.Row.Cells(5).Visible = False
            End If

            CType(e.Row.Cells(7).Controls(0), Button).OnClientClick = "if ( ! CreateFile()) return false;"

        End If

        'If e.Row.RowType = DataControlRowType.DataRow Then
        'e.Row.Attributes("ondblclick") = Page.ClientScript.GetPostBackClientHyperlink(GridView1, "Edit$" & e.Row.RowIndex)
        'e.Row.Attributes("style") = "cursor:pointer"
        'End If

    End Sub

    Protected Sub gv_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)


        'Dim cmd As String
        'Dim ds As New DataSet
        'cmd = "SELECT top 1 [Document No],[Current User] FROM TO_MR_OpenDoc WHERE [Active] = 1 and [Document No] = '" & GridView1.Rows(e.NewEditIndex).Cells(0).Text & "' and [Current User] <> '" & Session.Item("tmpuser") & "' order by [Date Time] desc"
        'ds = executeQuery(cmd)
        'If ds.Tables(0).Rows.Count = 0 Then
        '    tmpdocno = GridView1.Rows(e.NewEditIndex).Cells(0).Text
        '    tmpco = GridView1.Rows(e.NewEditIndex).Cells(1).Text
        '    tmpbrand = GridView1.Rows(e.NewEditIndex).Cells(2).Text
        '    Session("tmpdoc") = GridView1.Rows(e.NewEditIndex).Cells(0).Text
        '    Session("tmpco") = GridView1.Rows(e.NewEditIndex).Cells(1).Text
        '    Session("tmpbrand") = GridView1.Rows(e.NewEditIndex).Cells(2).Text
        '    If GridView1.Rows(e.NewEditIndex).Cells(4).Text = "NAV ON QUEUE" Then
        '        Session("tmpviewedit") = "VIEW"
        '    Else
        '        Session("tmpviewedit") = "EDIT"
        '    End If
        '    Page.Response.Redirect("MRDetails.aspx")
        'Else
        '    lbl_msgbox.InnerText = "Warning! Document is currently being used by"
        '    lbl_msgbox.Visible = True
        '    lbl_msgbox.InnerHtml = lbl_msgbox.InnerText & " <strong>" & Replace(ds.Tables(0).Rows(0)("Current User").ToString.ToLower, "primergrp\", "") & "</strong>"
        'End If


    End Sub

    Protected Sub gv_CancelEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        '
    End Sub

    Protected Sub gv_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        '
    End Sub

    Private Function CreateText(tmpdocno As String, spName As String) As DataSet

        Dim sqlcmd As New SqlCommand

            sqlcmd.CommandText = spName

            sqlcmd.CommandType = CommandType.StoredProcedure
            sqlcmd.CommandTimeout = 0
            'sqlcmd.Connection = objConn
            sqlcmd.Parameters.Add("@DocNo", SqlDbType.NVarChar).Value = tmpdocno

            Using con As New SqlConnection(connstr)
                Using sda As New SqlDataAdapter()
                    sqlcmd.Connection = con
                    sda.SelectCommand = sqlcmd
                    Using ds As New DataSet()
                        sda.Fill(ds)
                        Return ds
                    End Using
                End Using
                con.Dispose()
            End Using



    End Function

    Protected Function TextFile(tmpsp As String, tmp0 As String, tmp1 As String, tmp2 As String, tmpname As String) As String

        TextFile = ""
            Dim sPath As String = Server.MapPath("\Uploads\email.txt")

            Dim objStreamReader As StreamReader
            objStreamReader = File.OpenText(sPath)
            Dim contents As String = objStreamReader.ReadToEnd()
            objStreamReader.Close()

            'Dim sScript As String = ""
            Dim tmpds As DataSet
            tmpds = Me.CreateText(tmp0, tmpsp)
            Dim tmpfile As String = tmp0 + " - " + tmpname + " - " + tmp2
        If tmpds.Tables(0).Rows.Count <> 0 Then
            Dim strFile As String = Server.MapPath("\Uploads\" + tmpfile + ".txt")
            'Dim strServerFile As String = "\\fileserver\ICT\Web Merchandise Replenishment (WMR)\" + tmp1 + "\" + tmpfile + ".txt"
            Dim strServerFile As String = Server.MapPath("\LocalFolder\" + tmp1 + "\" + tmpfile + ".txt")
            'Dim strServerFile As String = "\\10.63.1.231\Navision_Shared_Files\Merchandising\Web Merchandise Replenishment (WMR)\" + tmp1 + "\" + tmpfile + ".txt"

            Dim os As New StreamWriter(strFile)
            For Each dr As DataRow In tmpds.Tables(0).Rows
                os.WriteLine(dr("From Code").ToString + vbTab + dr("To Code").ToString +
                        vbTab + dr("Transit Code").ToString + vbTab + Format(dr("Posting Date"), "MM/dd/yyyy").ToString + vbTab + dr("Document No").ToString +
                        vbTab + dr("Transfer Type").ToString + vbTab + dr("TO Reason").ToString + vbTab + dr("Item No").ToString +
                        vbTab + dr("Request Qty").ToString + vbTab + dr("Remarks").ToString)
            Next
            os.Close()
            Try
                File.Copy(strFile, strServerFile)
                contents = contents.Replace("[@DOCNO@]", tmp0 + " - " + tmpname + " - " + tmp2)
                SendMail("support.mss@primergrp.com", contents, tmp0, strServerFile)
                'SendMail("gilbert.laman@primergrp.com", contents, tmp0, strServerFile)
            Catch ex As Exception
                Dim cmd As String
                Dim ds As New DataSet

                cmd = "INSERT INTO Email_Logs ([ErrorMessage]) values ('" & ex.Message & "')"
                executeQuery(cmd)
            End Try

            TextFile = " window.open('Download.aspx?file=" & tmpfile & "', '_blank');"
        End If




    End Function

    Protected Function SMTextFile(tmpsp As String, tmp0 As String, tmp1 As String, tmp2 As String, tmpname As String) As String
        SMTextFile = ""
        Dim sPath As String = Server.MapPath("\Uploads\email.txt")

        Dim objStreamReader As StreamReader
        objStreamReader = File.OpenText(sPath)
        Dim contents As String = objStreamReader.ReadToEnd()
        objStreamReader.Close()


        Dim tmpds0 As DataSet
        tmpds0 = Me.CreateText(tmp0, tmpsp)
        Dim tmpds As DataSet
        tmpds = Me.CreateText(tmp0, Replace(tmpsp, "0", ""))

        If tmpds0.Tables(0).Rows.Count <> 0 Then
            For Each dr0 As DataRow In tmpds0.Tables(0).Rows
                'SMTextFile = SMTextFile + "-" + dr0("Sub Dept").ToString + " " + dr0("MainCat").ToString

                Dim tmpfile As String = tmp0 + " - " + tmpname + " - " + dr0("Sub Dept").ToString + " - " + dr0("MainCat").ToString + " - " + tmp2
                If tmpds.Tables(0).Rows.Count <> 0 Then
                    Dim strFile As String = Server.MapPath("\Uploads\" + tmpfile + ".txt")
                    'Dim strServerFile As String = "\\fileserver\ICT\Web Merchandise Replenishment (WMR)\" + tmp1 + "\" + tmpfile + ".txt"
                    Dim strServerFile As String = Server.MapPath("\LocalFolder\" + tmp1 + "\" + tmpfile + ".txt")
                    'Dim strServerFile As String = "\\10.63.1.231\Navision_Shared_Files\Merchandising\Web Merchandise Replenishment (WMR)\" + tmp1 + "\" + tmpfile + ".txt"

                    Dim os As New StreamWriter(strFile)
                    'tmpds.Tables(0).DefaultView.RowFilter = "[Sub Dept] = '" + dr0("Sub Dept").ToString + "' and [MainCat] = '" + dr0("MainCat").ToString + "'"
                    For Each dr As DataRow In tmpds.Tables(0).Rows
                        If dr("Sub Dept").ToString = dr0("Sub Dept").ToString And dr("MainCat").ToString = dr0("MainCat").ToString Then
                            os.WriteLine(dr("From Code").ToString + vbTab + dr("To Code").ToString +
                                vbTab + dr("Transit Code").ToString + vbTab + Format(dr("Posting Date"), "MM/dd/yyyy").ToString + vbTab + dr("Document No").ToString +
                                vbTab + dr("Transfer Type").ToString + vbTab + dr("TO Reason").ToString + vbTab + dr("Item No").ToString +
                                vbTab + dr("Request Qty").ToString + vbTab + dr("Remarks").ToString)
                        End If
                    Next
                    os.Close()
                    Try
                        File.Copy(strFile, strServerFile)
                        contents = contents.Replace("[@DOCNO@]", tmp0 + " - " + tmpname + " - " + tmp2)
                        SendMail("support.mss@primergrp.com", contents, tmp0, strServerFile)
                    Catch ex As Exception
                    End Try

                    SMTextFile = SMTextFile + " window.open('Download.aspx?file=" & tmpfile & "', '_blank');"
                End If

            Next
        Else
            SMTextFile = ""
        End If

    End Function

    Protected Sub gv_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)

        Dim docnum As String = Replace(GridView1.Rows(e.RowIndex).Cells(0).Text, "-", "")
        Dim cmdcount As String
        Dim dscount As New DataSet

        cmdcount = <code>
                       <![CDATA[
                       select sum([TOTAL REQ]) as [TOTALREQUEST] from [dbo].[]]><%= docnum %><![CDATA[]
                        ]]>
                   </code>.Value
        dscount = executeQuery(cmdcount)
        Dim totalcount As Integer = dscount.Tables(0).Rows(0)("TOTALREQUEST")
        If totalcount <> 0 Then
            If CheckNeg(GridView1.Rows(e.RowIndex).Cells(0).Text) = 0 Then

                Dim sMain As String = ""

                sMain = sMain & TextFile("spFSSTrade", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "FSS")
                sMain = sMain & TextFile("spFSSNonTrade", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "FSS-NONTRADE")
                sMain = sMain & TextFile("sp3PTrade", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "3P")
                sMain = sMain & TextFile("sp3PNonTrade", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "3P-NONTRADE")
                sMain = sMain & TextFile("spConTradeReg", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "CONREG")
                sMain = sMain & TextFile("spConTradeMD", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "CONMDN")
                sMain = sMain & TextFile("spConNonTrade", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "CON-NONTRADE")
                sMain = sMain & SMTextFile("spSMTradeReg0", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "SMCONREG")
                sMain = sMain & SMTextFile("spSMTradeMD0", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "SMCONMDN")
                sMain = sMain & TextFile("spSMNonTrade", GridView1.Rows(e.RowIndex).Cells(0).Text, GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(2).Text, "SMCON-NONTRADE")


                Dim cmd_docdetail_updateQty_select As String
                Dim ds_docdetail_updateQty_select As New DataSet

                cmd_docdetail_updateQty_select = "EXEC sp_selectDocDetail 0,'" & GridView1.Rows(e.RowIndex).Cells(0).Text & "',''"
                ds_docdetail_updateQty_select = executeQuery(cmd_docdetail_updateQty_select)

                If ds_docdetail_updateQty_select.Tables(0).Rows.Count > 0 Then
                    Dim cmd_docdetail_updateQty As String
                    Dim ds_docdetail_updateQty As New DataSet
                    For i = 0 To ds_docdetail_updateQty_select.Tables(0).Rows.Count - 1
                        cmd_docdetail_updateQty = "update [dbo].[DocDetail] set [ALLOCATEDQTY] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ALLOCATEDQTY") & "'
                                            where [DOCNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("Doc No") & "'
                                           and [ITEMNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ITEMNO") & "'
                                            and [BRAND] = '" & GridView1.Rows(e.RowIndex).Cells(2).Text & "'"
                        ds_docdetail_updateQty = executeQuery(cmd_docdetail_updateQty)
                    Next
                End If

                Dim objConn As SqlConnection = New SqlConnection(connstr)
                Dim sqlcmd As SqlCommand
                objConn.Open()
                Dim cmd As String =
                <code>
                    <![CDATA[

                        BEGIN TRANSACTION;
                            UPDATE [dbo].[DocDetail]
                            SET [STATUS] = 'C'
                            WHERE [DOCNO] + [BRAND]
                            IN
                            (
                            SELECT DISTINCT [Doc No] + [Brand]
                             FROM [dbo].[TO_MR_Collated_Header_2]
                             WHERE [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
                                and [Status] = 'IN-PROCESS'
                            ) and [ALLOCATEDQTY] = 0
                        COMMIT TRANSACTION;

                        BEGIN TRANSACTION;
                            UPDATE [dbo].[DocDetail]
                            SET 
                                [Status] = 'N'
                            WHERE [DOCNO] + [BRAND] IN
                             (
                             SELECT DISTINCT [Doc No] + [Brand]
                             FROM [dbo].[TO_MR_Collated_Header_2]
                             WHERE [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
                                and [Status] = 'IN-PROCESS'
                             ) and [STATUS] <> 'C'
                        COMMIT TRANSACTION;

                        BEGIN TRANSACTION;
                             UPDATE [dbo].[DocHeader]
                             SET
                             [Status] = 'N'
                             WHERE [DOCNO] in 
                             (
                             SELECT DISTINCT [Doc No]
                             FROM [dbo].[TO_MR_Collated_Header_2]
                             WHERE [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
                             and [Status] = 'IN-PROCESS'
                             )
                        COMMIT TRANSACTION;

                        BEGIN TRANSACTION;
                            UPDATE [dbo].[TO_MR_Collated_Header_2]
                            SET 
                                [Status] = 'NAV ON QUEUE'
                            WHERE
                                [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
                        COMMIT TRANSACTION;

                        ]]>
                </code>.Value
                sqlcmd = New SqlCommand(cmd, objConn)
                sqlcmd.ExecuteNonQuery()
                objConn.Close()
                objConn.Dispose()

                '

                ClientScript.RegisterStartupScript(Me.GetType(), "script", sMain, True)

                CreateLog(GridView1.Rows(e.RowIndex).Cells(0).Text)

                CreateGrid()
                BindGrid(ddlFields.SelectedValue, dd_status.Text)


            Else
                'Response.Write("<script language=""javascript"">alert('There are lines with insufficient quantity.');</script>")
                If (Not ClientScript.IsStartupScriptRegistered("alert")) Then
                    Page.ClientScript.RegisterStartupScript _
                (Me.GetType(), "alert", "ShowMsg();", True)
                End If
            End If


            'MessageLabel.Text = CheckNeg(GridView1.Rows(e.RowIndex).Cells(0).Text)

        Else
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('Cant create NAV text file. This document number have 0 request total!');", True)
        End If



    End Sub

    Protected Sub RowDeleting1(sender As Object, e As GridViewDeleteEventArgs)
        Dim sScript As String = ""
        Dim sScript2 As String = ""

        Dim sPath As String = Server.MapPath("\Uploads\email.txt")

        Dim objStreamReader As StreamReader
        objStreamReader = File.OpenText(sPath)
        Dim contents As String = objStreamReader.ReadToEnd()
        objStreamReader.Close()

        Dim tmpds As DataSet
        tmpds = Me.CreateText1(GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(0).Text)
        Dim tmpfile As String = GridView1.Rows(e.RowIndex).Cells(0).Text + " FSS-CONREG " + GridView1.Rows(e.RowIndex).Cells(2).Text
        If tmpds.Tables(0).Rows.Count <> 0 Then
            Dim strFile As String = Server.MapPath("\Uploads\" + tmpfile + ".txt")
            'Dim strServerFile As String = "\\fileserver\ICT\Web Merchandise Replenishment (WMR)\" + GridView1.Rows(e.RowIndex).Cells(1).Text + "\" + tmpfile + ".txt"
            Dim strServerFile As String = Server.MapPath("\LocalFolder\" + GridView1.Rows(e.RowIndex).Cells(1).Text + "\" + tmpfile + ".txt")
            'Dim strServerFile As String = "\\10.63.1.231\Navision_Shared_Files\Merchandising\Web Merchandise Replenishment (WMR)\" + GridView1.Rows(e.RowIndex).Cells(1).Text + "\" + tmpfile + ".txt"

            Dim os As New StreamWriter(strFile)
            'Dim os As New StreamWriter("C:\Uploads\" + tmpfile + ".txt")
            For Each dr As DataRow In tmpds.Tables(0).Rows
                os.WriteLine(dr("From Code").ToString + vbTab + dr("To Code").ToString +
                    vbTab + dr("Transit Code").ToString + vbTab + Format(dr("Posting Date"), "MM/dd/yyyy").ToString + vbTab + dr("Document No").ToString +
                    vbTab + dr("Transfer Type").ToString + vbTab + dr("TO Reason").ToString + vbTab + dr("Item No").ToString +
                    vbTab + dr("Request Qty").ToString + vbTab + dr("Remarks").ToString)
            Next
            os.Close()
            Try
                File.Copy(strFile, strServerFile)
                contents = contents.Replace("[@DOCNO@]", GridView1.Rows(e.RowIndex).Cells(0).Text + " with " + GridView1.Rows(e.RowIndex).Cells(2).Text + " Brand")
                SendMail("support.mss@primergrp.com", contents, GridView1.Rows(e.RowIndex).Cells(0).Text, strServerFile)
            Catch ex As Exception
            End Try

            sScript = " window.open('Download.aspx?file=" & tmpfile & "', '_blank');"
        End If

        Dim tmpds2 As DataSet
        tmpds2 = Me.CreateText2(GridView1.Rows(e.RowIndex).Cells(1).Text, GridView1.Rows(e.RowIndex).Cells(0).Text)
        Dim tmpfile2 As String = GridView1.Rows(e.RowIndex).Cells(0).Text + " CONMARK " + GridView1.Rows(e.RowIndex).Cells(2).Text
        If tmpds2.Tables(0).Rows.Count <> 0 Then
            Dim strFile2 As String = Server.MapPath("\Uploads\" + tmpfile2 + ".txt")
            'Dim strServerFile2 As String = "\\fileserver\ICT\Web Merchandise Replenishment (WMR)\" + GridView1.Rows(e.RowIndex).Cells(1).Text + "\" + tmpfile2 + ".txt"
            Dim strServerFile2 As String = Server.MapPath("\LocalFolder\" + GridView1.Rows(e.RowIndex).Cells(1).Text + "\" + tmpfile + ".txt")
            'Dim strServerFile2 As String = "\\10.63.1.231\Navision_Shared_Files\Merchandising\Web Merchandise Replenishment (WMR)\" + GridView1.Rows(e.RowIndex).Cells(1).Text + "\" + tmpfile2 + ".txt"

            Dim os2 As New StreamWriter(strFile2)
            For Each dr As DataRow In tmpds2.Tables(0).Rows
                os2.WriteLine(dr("From Code").ToString + vbTab + dr("To Code").ToString +
                    vbTab + dr("Transit Code").ToString + vbTab + Format(dr("Posting Date"), "MM/dd/yyyy").ToString + vbTab + dr("Document No").ToString +
                    vbTab + dr("Transfer Type").ToString + vbTab + dr("TO Reason").ToString + vbTab + dr("Item No").ToString +
                    vbTab + dr("Request Qty").ToString + vbTab + dr("Remarks").ToString)
            Next
            os2.Close()
            Try
                File.Copy(strFile2, strServerFile2)
                contents = contents.Replace("[@DOCNO@]", tmpfile2 + " with " + GridView1.Rows(e.RowIndex).Cells(2).Text + " Brand")
                SendMail("support.mss@primergrp.com", contents, GridView1.Rows(e.RowIndex).Cells(0).Text, strServerFile2)
            Catch ex As Exception
            End Try

            sScript2 = " window.open('Download.aspx?file=" & tmpfile2 & "', '_blank');"
        End If


        'Dim objConn As SqlConnection = New SqlConnection(connstr)
        'Dim sqlcmd As SqlCommand
        'objConn.Open()
        'Dim cmd As String =
        '    <code>
        '        <![CDATA[
        '            BEGIN TRANSACTION;
        '                UPDATE [TOMR].[dbo].[DocDetail]
        '                SET 
        '                    [Status] = 'N'
        '                WHERE [DOCNO] + [ITEMNO] IN
        '                 (
        '                 SELECT [Doc No] + [Item No]
        '                 FROM [TOMR].[dbo].[TO_MR_Collated_Lines]
        '                 WHERE [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
        '                    and [Status] = 'IN-PROCESS'
        '                 )
        '            COMMIT TRANSACTION;

        '            BEGIN TRANSACTION;
        '                UPDATE [TOMR].[dbo].[TO_MR_Collated_Header]
        '                SET 
        '                    [Status] = 'NAV ON QUEUE'
        '                WHERE
        '                    [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
        '                    and [Status] = 'IN-PROCESS'
        '            COMMIT TRANSACTION;

        '            BEGIN TRANSACTION;
        '                UPDATE [TOMR].[dbo].[TO_MR_Collated_Lines]
        '                SET 
        '                    [Status] = 'NAV ON QUEUE'
        '                WHERE
        '                    [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
        '                    and [Status] = 'IN-PROCESS'
        '            COMMIT TRANSACTION;
        '            ]]>
        '    </code>.Value
        'sqlcmd = New SqlCommand(cmd, objConn)
        'sqlcmd.ExecuteNonQuery()
        'objConn.Close()
        'objConn.Dispose()


        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()
        Dim cmd As String =
            <code>
                <![CDATA[
                    
                    BEGIN TRANSACTION;
                        UPDATE [dbo].[TO_MR_Collated_Header_2]
                        SET 
                            [Status] = 'NAV ON QUEUE'
                        WHERE
                            [Document No] = ']]><%= GridView1.Rows(e.RowIndex).Cells(0).Text %><![CDATA['
                    COMMIT TRANSACTION;

                    ]]>
            </code>.Value
        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()


        Dim sMain As String
        sMain = sScript & sScript2
        ClientScript.RegisterStartupScript(Me.GetType(), "script", sMain, True)

        CreateLog(GridView1.Rows(e.RowIndex).Cells(0).Text)

        CreateGrid()
        BindGrid(ddlFields.SelectedValue, dd_status.Text)


        'Dim sPath As String = Server.MapPath("\Uploads\email.txt")
        'Dim objStreamReader As StreamReader
        'objStreamReader = File.OpenText(sPath)
        'Dim contents As String = objStreamReader.ReadToEnd()
        'objStreamReader.Close()
        'contents = contents.Replace("[@DOCNO@]", "Test Mail")
        'SendMail("ireneo.lautillo@primergrp.com;elijah.stodomingo@primergrp.com", contents)
        ''SendMail("elijah.stodomingo@primergrp.com;jacqueline.reaport@primergrp.com;support.mss@primergrp.com", contents)

    End Sub

    Sub gv_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles GridView1.SelectedIndexChanged
        '
    End Sub

    Sub gv_SelectedIndexChanging(ByVal sender As Object, ByVal e As GridViewSelectEventArgs)
        '
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Function SendMail(emailRecipient As String, emailBody As String, tmpsub As String, strFilePath As String) As Integer

        Try


            Dim smtpServer As String = My.Settings.EmailServerNew
            Dim emailSender As String = My.Settings.EmailSender
            Dim emailSenderDisplayName As String = My.Settings.EmailSenderName
            Dim emailSubject As String = tmpsub
            ' Dim client As New SmtpClient(smtpServer, 587)
            Dim client As New SmtpClient(smtpServer, My.Settings.port_new)
            ' client.EnableSsl = True
            client.EnableSsl = False
            ' client.Credentials = New System.Net.NetworkCredential("unified@primergrp.com", "unifiedPa$$w0rd")
            client.Credentials = New System.Net.NetworkCredential("", "")
            Dim [from] As New MailAddress(emailSender, emailSenderDisplayName)
            Dim sMail As String
            Dim sTo As String() = emailRecipient.Split(";")

            Dim [to] As New MailAddress(emailRecipient) '

            Dim message As New MailMessage([from], [to])

            For Each sMail In sTo
                message.Bcc.Add("" & sMail & "")
            Next

            message.IsBodyHtml = True
            message.CC.Add("gilbert.laman@primergrp.com")


            message.Body = emailBody
            message.Subject = emailSubject

            Dim myFile As Net.Mail.Attachment = New Net.Mail.Attachment(strFilePath)
            message.Attachments.Add(myFile)

            client.Send(message)


            Return 0
        Catch ex As Exception

            Dim cmd As String
            Dim ds As New DataSet

            cmd = "INSERT INTO Email_Logs ([ErrorMessage]) values ('" & ex.Message & "')"
            executeQuery(cmd)

            Return -1

        End Try

    End Function


    Protected Sub ddlFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlFields.SelectedIndexChanged
        Dim cmd As String
        Dim ds As New DataSet

        Try
            If ddlFields.SelectedValue = "Company" Then
                cmd = "select distinct [company] from [TO_MR_Collated_Header_2] order by [company] asc"
                ds = executeQuery(cmd)
                dd_status.Items.Clear
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dd_status.Items.Add(ds.Tables(0).Rows(i)("company"))
                Next
            ElseIf ddlFields.SelectedValue = "Brand" Then
                cmd = "select distinct [brand] from [TO_MR_Collated_Header_2] order by [brand] asc"
                ds = executeQuery(cmd)
                dd_status.Items.Clear
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dd_status.Items.Add(ds.Tables(0).Rows(i)("brand"))
                Next
            ElseIf ddlFields.SelectedValue = "Status" Then
                cmd = "select distinct [status] from [TO_MR_Collated_Header_2] order by [status] asc"
                ds = executeQuery(cmd)
                dd_status.Items.Clear()
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dd_status.Items.Add(ds.Tables(0).Rows(i)("status"))
                Next
            End If
        Catch ex As Exception
            'do nothing
        End Try
    End Sub

    'Protected Sub txtFilter_TextChanged(sender As Object, e As EventArgs) Handles txtFilter.TextChanged
    '    BindGrid(ddlFields.SelectedValue, dd_status.Text)
    'End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim status As Integer
        Dim cmd As String
        Dim ds As New DataSet
        If e.CommandName = "Downloadbtn" Then
            Session("selectedDoc") = GridView1.Rows(e.CommandArgument).Cells(0).Text
            cmd = "SELECT top 1 [Document No],[Current User] FROM TO_MR_OpenDoc WHERE [Active] = 1 and [Document No] = '" & Session("selectedDoc") & "' and [Current User] <> '" & Session.Item("tmpuser") & "' order by [Date Time] desc"
            ds = executeQuery(cmd)
            If ds.Tables(0).Rows.Count = 0 Then
                Session("selectedDoc") = Replace(GridView1.Rows(e.CommandArgument).Cells(0).Text, "-", "")
                Session("selectedDocwithoutdash") = GridView1.Rows(e.CommandArgument).Cells(0).Text

                cmd = "insert into TO_MR_OpenDoc ([Document No],[Current User],[Date Time],[Active]) values ('" & GridView1.Rows(e.CommandArgument).Cells(0).Text & "','" & Session.Item("tmpuser") & "','" & DateAndTime.Now & "','1')"
                executeQuery(cmd)

                Response.Redirect("ExportToExcel.aspx")
            Else
                status = 1
            End If
        ElseIf e.CommandName = "Uploadbtn" Then
            Session("selectedDoc") = GridView1.Rows(e.CommandArgument).Cells(0).Text
            cmd = "SELECT top 1 [Document No],[Current User] FROM TO_MR_OpenDoc WHERE [Active] = 1 and [Document No] = '" & Session("selectedDoc") & "' and [Current User] <> '" & Session.Item("tmpuser") & "' order by [Date Time] desc"
            ds = executeQuery(cmd)
            If ds.Tables(0).Rows.Count = 0 Then
                Session("selectedDoc") = GridView1.Rows(e.CommandArgument).Cells(0).Text
                Response.Redirect("Upload.aspx")
            Else
                status = 1
            End If
        ElseIf e.CommandName = "viewbtn" Then
            Session("selectedDoc") = GridView1.Rows(e.CommandArgument).Cells(0).Text
            cmd = "SELECT top 1 [Document No],[Current User] FROM TO_MR_OpenDoc WHERE [Active] = 1 and [Document No] = '" & Session("selectedDoc") & "' and [Current User] <> '" & Session.Item("tmpuser") & "' order by [Date Time] desc"
            ds = executeQuery(cmd)
            If ds.Tables(0).Rows.Count = 0 Then
                tmpdocno = GridView1.Rows(e.CommandArgument).Cells(0).Text
                tmpco = GridView1.Rows(e.CommandArgument).Cells(1).Text
                tmpbrand = GridView1.Rows(e.CommandArgument).Cells(2).Text
                Session("tmpdoc") = GridView1.Rows(e.CommandArgument).Cells(0).Text
                Session("tmpco") = GridView1.Rows(e.CommandArgument).Cells(1).Text
                Session("tmpbrand") = GridView1.Rows(e.CommandArgument).Cells(2).Text
                Session("mode") = GridView1.Rows(e.CommandArgument).Cells(4).Text
                If GridView1.Rows(e.CommandArgument).Cells(4).Text = "NAV ON QUEUE" Or GridView1.Rows(e.CommandArgument).Cells(4).Text = "CANCELLED" Then
                    Session("tmpviewedit") = "VIEW"
                Else
                    Session("tmpviewedit") = "EDIT"
                End If
                Page.Response.Redirect("MRDetails.aspx")
            Else
                status = 1
            End If

        End If

        If status = 1 Then
            lbl_msgbox.InnerText = "Warning! Document is currently being used by"
            lbl_msgbox.Visible = True
            lbl_msgbox.InnerHtml = lbl_msgbox.InnerText & " <strong>" & Replace(ds.Tables(0).Rows(0)("Current User").ToString.ToLower, "primergrp\", "") & "</strong>"
        Else
            lbl_msgbox.Visible = False
        End If

    End Sub

    Private Sub dd_status_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dd_status.SelectedIndexChanged
        BindGrid(ddlFields.SelectedValue, dd_status.Text)
    End Sub

    Protected Sub exportExcel()

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

    Private Sub downloadbtn_Click(sender As Object, e As ImageClickEventArgs) Handles downloadbtn.Click
        Dim cmd_filtered_date As String
        Dim ds_filtered_date As DataSet
        Dim filtered_date As String

        cmd_filtered_date = "select top 1 format(filtered_date,'MM/dd/yyyy') as [filtered_date] from FilteredDate_Settings order by [ID] desc"
        ds_filtered_date = executeQuery(cmd_filtered_date)
        filtered_date = ds_filtered_date.Tables(0).Rows(0)("filtered_date").ToString

        Dim cmd_updater As String
        Dim ds_updater As New DataSet

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

            Dim cmd_download As String
            cmd_download = "insert into TO_MR_Log2 ([user name],[date],[doc no],[action]) values ('" & Session("tmpuser") & "','" & DateAndTime.Now & "','MR LIST','download MR')"
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
            cmd_columns = <code>
                              <![CDATA[
                    select distinct
	                a.[Document No]
                    , a.[Company]
                    , a.[Brand]
                    , Format(a.[Created Date],'MM/dd/yyyy') as [Created Date]
                    , a.[Status]
	                , ISNULL(Format(log2.[date],'MM/dd/yyyy'), '') as [Text File Date]
                    into #tmpMR
                from [dbo].[TO_MR_Collated_Header_2] as a with (nolock)

                left outer join
	                ( select
		                [doc no], MIN([date]) as [date]
	                from [dbo].[TO_MR_Log2]
	                where [action] = 'textfile'
					group by [doc no]
	                ) as log2
	                on a.[Document No] = log2.[doc no]

                where [Company] in (
					select distinct [Company] 
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)
					and [Brand] in (
					select distinct [Brand]
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)
                    
                    select * 
                    into #columnsTemp
                    from #tmpMR
                    
                    select [name] as [COLUMN_NAME]
                    from tempdb.sys.columns
                    where object_id = object_id('tempdb..#columnsTemp')
                    drop table #columnsTemp
                    
                    drop table #tmpMR    
                                ]]>
                          </code>.Value
            ds_columns = executeQuery(cmd_columns)

            '// populate header columns to excel
            Dim x As Integer = 1
            For i = 0 To ds_columns.Tables(0).Rows.Count - 1
                xlsWorkSheet.Cells(1, x) = ds_columns.Tables(0).Rows(i)("COLUMN_NAME")
                x = x + 1
            Next

            Dim cmd_data As String
            Dim ds_data As New DataSet
            cmd_data = <code>
                           <![CDATA[
                select distinct
	                a.[Document No]
                    , a.[Company]
                    , a.[Brand]
                    , Format(a.[Created Date],'MM/dd/yyyy') as [Created Date]
                    , a.[Status]
	                , ISNULL(Format(log2.[date],'MM/dd/yyyy'), '') as [Text File Date]
                from [dbo].[TO_MR_Collated_Header_2] as a with (nolock)

                left outer join
	                ( select
		                [doc no], MIN([date]) as [date]
	                from [dbo].[TO_MR_Log2]
	                where [action] = 'textfile'
					group by [doc no]
	                ) as log2
	                on a.[Document No] = log2.[doc no]

                where [Company] in (
					select distinct [Company] 
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)
					and [Brand] in (
					select distinct [Brand]
					from [dbo].[TO_MR_User_Security]
					where [user name] = ']]><%= Session("tmpuser") %><![CDATA['
					)
                     
                            ]]>
                       </code>.Value

            Select Case ddlFields.SelectedValue
                Case "Brand"
                    cmd_data &= " and [Brand] like '" & dd_status.SelectedItem.Text & "' and [created date] >= '" & filtered_date & "'"
                Case "Company"
                    cmd_data &= " and [Company] like '" & dd_status.SelectedItem.Text & "' and [created date] >= '" & filtered_date & "'"
                Case "Status"
                    cmd_data &= " and [Status] like '" & dd_status.SelectedItem.Text & "' and [created date] >= '" & filtered_date & "'"
            End Select

            cmd_data &= " order by [Status], [Created Date], [Company], [Brand]"
            ds_data = executeQuery(cmd_data)

            Dim rowstart As Integer = 2
            Dim alphabet_increment As String
            Dim zplus As Integer = 0

            '// bold header
            xlsWorkSheet.Range("a1:az1").Font.Bold = True

            For r = 0 To ds_data.Tables(0).Rows.Count - 1 '// for rows
                For c = 0 To ds_columns.Tables(0).Rows.Count - 1 '// for columns
                    If c > 25 Then
                        alphabet_increment = "a" & Chr(Asc("a") + Val(zplus))
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
            xfile = Server.MapPath("~\ExcelExport") & "\" & Format(Date.Now, "MMdd") & "-" & series_ & "-" & "(" & Replace(Replace(Session("tmpuser").ToString.ToLower, "primergrp\", ""), ".", " ") & ")" & ".xlsx"
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
            Response.AppendHeader("Content-Disposition", "attachment; filename=""MR.xlsx")
            Response.TransmitFile(Server.MapPath("~/ExcelExport/" & Format(Date.Now, "MMdd") & "-" & series_ & "-" & "(" & Replace(Replace(Session("tmpuser").ToString.ToLower, "primergrp\", ""), ".", " ") & ")" & ".xlsx"))
            Response.End()
        Catch ex As Exception
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
            Exit Try
        End Try
    End Sub

    Private Sub GridView1_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles GridView1.RowEditing

    End Sub

    'Private Sub reportsBtn_Click(sender As Object, e As EventArgs) Handles reportsBtn.Click
    '    Response.Redirect("Reports.aspx")
    'End Sub
End Class