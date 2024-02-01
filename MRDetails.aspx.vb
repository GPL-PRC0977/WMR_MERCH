Imports System.Drawing
Imports System.Data.SqlClient


Public Class MRDetails
    Inherits System.Web.UI.Page

    Dim connstr As String = [TO].My.Settings.SQLConnection.ToString '"Server='10.63.1.161';Initial Catalog='TOMR';user id='sa';password='Pa$$w0rd';"

    Dim dsmr As DataSet
    Dim dsloc As DataSet
    Dim tmptext As String
    Dim tmpint As Integer = 0
    Dim currow As Integer = -1
    'Dim tmpid As integer = 0
    Dim txtrow As Integer = 0
    Dim tmpTextbox As TextBox
    Dim lbl As Label
    Dim lbl2 As Label
    Dim ltr As LiteralControl
    Dim ltr2 As LiteralControl
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mainnav As Control = Master.FindControl("main_navigation")
        mainnav.Visible = False
        If CType(Session.Item("tmpuser"), String) = "" Then
            Page.Response.Redirect("Default.aspx")
        End If

        div_message_box.Visible = False

        If Session("deletesuccess") = "deleted" Then
            ClientScript.RegisterStartupScript(Me.GetType(), "script", " alert('Item is successfully deleted.'); ", True)
        End If

        Session("deletesuccess") = ""

        'If Session("mode") = "NAV ON QUEUE" Then
        '    buttonSave.Enabled = False
        '    'dropdownItem.Enabled = False
        '    buttonAdd.Enabled = False
        'End If

        Dim dt As DateTime = Now

        DocNoLabel.Text = "DOCUMENT NO. : " + Session("tmpdoc") & " - " & "(" & Session("mode") & ")"

        If Not IsPostBack Then

            'If CheckDocumentSession() = True Then
            'ClientScript.RegisterStartupScript(Me.GetType(), "script", " alert('Document is currently being used by other user.'); window.open('MR.aspx','_self'); ", True)
            'End If

            LoadMainCat()
            LoadItems()
            CreateGrid2()
            BindGrid2("ALL")

            If Session("tmpviewedit") = "EDIT" Then CreateDocumentSession()

        End If

        If Session("tmpviewedit") = "VIEW" Or Session("mode") = "CANCELLED" Then

            buttonSave.Visible = False
            Label1.Visible = False
            dropdownItem.Enabled = False
            txt_itemcode.Enabled = False
            addItems.Visible = False
            buttondelete.Visible = False
            btn_delete_qtyavail_negative.Visible = False
            btn_delete_whs_negative.Visible = False

        End If

        Dim dt2 As DateTime = Now
        Label3.Text = DateDiff(DateInterval.Second, dt, dt2)

        'LoadItems_init()

    End Sub

    Private Function CheckDocumentSession()
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        Dim bCurrentlyOpen As Boolean
        objConn.Open()
        Dim cmd As String = "SELECT Distinct [Document No] FROM TO_MR_OpenDoc WHERE [Active] = 1 and [Document No] = '" & Session("tmpdoc") & "' and [Current User] <> '" & Session.Item("tmpuser") & "' "

        sqlcmd = New SqlCommand(cmd, objConn)
        Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
        bCurrentlyOpen = sqlreader.HasRows = True
        objConn.Close()
        objConn.Dispose()
        'bCurrentlyOpen = True
        Return bCurrentlyOpen
    End Function

    Private Function CreateDocumentSession() As Boolean
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "insert into [dbo].[TO_MR_OpenDoc] ([Document No], [Current User]) values ('" & Session("tmpdoc") & "','" & Session.Item("tmpuser") & "') "

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()
        Return True
    End Function

    Sub EndDocumentSession()
        Dim connstr As String = [TO].My.Settings.SQLConnection.ToString
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "update [dbo].[TO_MR_OpenDoc] SET [Active] = 0 WHERE [Current User] = '" & Session.Item("tmpuser") & "' and [Document No] = '" & Session("tmpdoc") & "'"

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        Page.Response.Redirect("MR.aspx")
    End Sub

    Private Function CreateLog(docno As String) As Boolean
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "insert into [dbo].[TO_MR_Log2] ([user name], [date], [doc no], [action]) values ('" & Session("tmpuser") & "', GetDate(), '" & docno & "', 'save') "

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()
        Return True
    End Function

    Private Function GetData(tmpcat As String) As DataSet
        Dim sqlscript As String =
        <code>
            <![CDATA[

                declare @qtyscript as nvarchar(MAX)
                declare @tmpcat as nvarchar(50)

                set @tmpcat = ']]><%= tmpcat %><![CDATA['
                
                create table #tmp1
                (
					[ID] [int] identity(1,1) NOT NULL,
					[Document No] [nvarchar](50) NULL,
					[Item No] [nvarchar](20) NULL,
					[Item Description] [nvarchar](250) NULL,
					[Item Owner] [nvarchar](10) NULL,
					[Brand] [nvarchar](50) NULL,
					[Main Category] [nvarchar](50) NULL,
                    [SRP] [numeric](20,2) NULL,
                    [Price Type] [nvarchar](20) NULL,
					[Location Code] [nvarchar](20) NULL,
					[Location Name] [nvarchar](200) NULL,
                    [Location Type] [nvarchar](100) NULL,
					[Request Qty] [int] NULL,
					[Doc No] [nvarchar](50) NULL,
					[Doc Date] [date] NULL,
					[Remarks] [nvarchar](250) NULL,
					[Status] [nvarchar](25) NULL,
					[Whse Qty] [int] NULL,
					[Avail Qty] [int] NULL
                )

                insert into #tmp1
                select 
					a.[Document No]
					, a.[Item No]
					, a.[Item Description]
					, a.[Item Owner]
					, a.[Brand]
					, a.[Main Category]
                    , a.[SRP]
                    , a.[Price Type]
					, a.[Location Code]
					, a.[Location Name]
                    , a.[Location Type]
					, a.[Request Qty]
					, a.[Doc No]
					, a.[Doc Date]
					, a.[Remarks]
					, a.[Status]
					, ISNULL(wqty.[Whse Qty], '0') as [Whse Qty]
					, ISNULL((ISNULL(wqty.[Whse Qty], '0') - b.[Req Qty]), '0') as [Avail Qty]
                from [dbo].[TO_MR_Collated_Lines] as a with (nolock)
				left outer join 
					( select 
						[Item Owner]
						, [Item No]
						, SUM([Whse Qty]) as [Whse Qty]
						from [dbo].[TO_MR_Invt_Qty] with (nolock)
						group by [Item Owner], [Item No]
					) as wqty 
					on a.[Item Owner] = wqty.[Item Owner]
					and a.[Item No] = wqty.[Item No]
				left outer join 
					( select 
                        [Document No]
						, [Item Owner]
						, [Item No]
						, SUM([Request Qty]) as [Req Qty]
						from [dbo].[TO_MR_Collated_Lines] with (nolock)
						where [Status] = 'IN-PROCESS'
						group by [Document No], [Item Owner], [Item No]
					) as b
					on a.[Item Owner] = b.[Item Owner]
					and a.[Item No] = b.[Item No]
                    and a.[Document No] = b.[Document No]
                where 
	                a.[Document No] = ']]><%= Session("tmpdoc") %><![CDATA['
                    and a.[Main Category] like (
					CASE
					WHEN @tmpcat = 'ALL' THEN '%'
					ELSE @tmpcat
					END)

                DECLARE @cols AS NVARCHAR(MAX),
                    @query  AS NVARCHAR(MAX);

                SET @cols = STUFF((select ',' + QUOTENAME(tmp1.[Location Name]) 
                            FROM #tmp1 as tmp1
							LEFT OUTER JOIN [dbo].[TO_MR_Store_Sort] as store with (nolock)
								on tmp1.[Location Code] = store.[Location Code]
							GROUP BY tmp1.[Location Name], store.[Ctrl No]
							ORDER BY store.[Ctrl No] ASC
                            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '')

                set @query = 
	                '
	                SELECT
						[Document No]
						, [Item No]
						, [Item Description]
						, [Item Owner]
						, [Brand]
						, [Main Category]
                        , [SRP]
						, [Price Type]
						, [Whse Qty]
						, [Total Req]
						, [Avail Qty]
						, [Gender]
						, [Search Description]
						, [Color]
						, [Size]
		                , ' + @cols + ' 
	                from 
		                (
		                select
							aa.[Document No]
							, aa.[Item No]
							, aa.[Item Description]
							, aa.[Item Owner]
							, aa.[Brand]
							, aa.[Main Category]
                            , aa.[SRP]
							, aa.[Price Type]
							, aa.[Location Name]
							, bb.[Total Req]
							, aa.[Request Qty]
							, aa.[Whse Qty]
							, aa.[Avail Qty]
							, itm.[Gender]
							, itm.[Search Description]
							, itm.[Color]
							, itm.[Size]
		                from #tmp1 as aa
						left outer join 
							( select
								[Item No]
								, SUM([Request Qty]) as [Total Req]
							from #tmp1
							group by [Item No]
							) as bb
							on aa.[Item No] = bb.[Item No]
						left outer join [Reports].[dbo].[t_item_master] as itm with (nolock)
							on aa.[Item No] = itm.[Item No]
						group by 
							aa.[Document No]
							, aa.[Item No]
							, aa.[Item Description]
							, aa.[Item Owner]
							, aa.[Brand]
							, aa.[Main Category]
                            , aa.[SRP]
							, aa.[Price Type]
							, aa.[Location Name]
							, bb.[Total Req]
							, aa.[Request Qty]
							, aa.[Whse Qty]
							, aa.[Avail Qty]
							, itm.[Gender]
							, itm.[Search Description]
							, itm.[Color]
							, itm.[Size]
		                ) as x

	                pivot 
		                (
			                sum([Request Qty])
			                for [Location Name] in (' + @cols + ')
		                ) as p 
										
					order by
						[Brand] ASC
						, [Main Category] ASC
						, [Gender] ASC
						, [Search Description] ASC
						, [Color] ASC
						, [Size] ASC
						, [Price Type] ASC
						, [SRP] ASC
						, [Item No] ASC

	                '

                execute(@query)

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

    Private Function GetLocations() As DataSet
        Dim sqlscript As String =
        <code>
            <![CDATA[

                declare @tmpcat as nvarchar(50)

                set @tmpcat = ']]><%= dropdownMainCat.SelectedItem.Text %><![CDATA['
                
                select loc.[Location Name]
                from [dbo].[TO_MR_Store_Sort] as a with (nolock)
                left outer join 
	                ( select distinct [Location Code], [Location Name]
	                from [Reports].[dbo].[t_locations4] with (nolock)
	                ) as loc
	                on a.[Location Code] = loc.[Location Code]
                where a.[Location Code] in (
	                select [Location Code]
	                from [dbo].[TO_MR_Collated_Lines]
	                where [Document No] = ']]><%= Session("tmpdoc") %><![CDATA['
	                and [Main Category] like (
	                CASE
	                WHEN @tmpcat = 'ALL' THEN '%'
	                ELSE @tmpcat
	                END
	                ))
                order by a.[Ctrl No]
                
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

    Sub LoadMainCat()
        Dim sConnectionString As String
        Dim objConn As SqlConnection
        sConnectionString = [TO].My.Settings.SQLConnection.ToString '"Server='10.63.1.161';Initial Catalog='Reports';user id='sa';password='Pa$$w0rd';"
        objConn = New SqlConnection(sConnectionString)
        objConn.Open()
        Dim sqlcmd As New SqlCommand("SELECT DISTINCT " &
            "[MainCat Description] " &
            "FROM [Reports].[dbo].[t_item_master] " &
            "WHERE [Company Ownership Code] = '" + Session("tmpco") + "' and [Brand Description] = '" + Replace(Replace(Session("tmpbrand"), "&#39;", ""), "amp;", "") + "' " &
            "ORDER by [MainCat Description] ", objConn)
        Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
        dropdownMainCat.Items.Clear()
        dropdownMainCat.Items.Add("ALL")
        Do While sqlreader.Read
            dropdownMainCat.Items.Add(sqlreader("MainCat Description"))
        Loop
        sqlreader.Close()
        objConn.Close()
        objConn.Dispose()
    End Sub

    Sub LoadItems()
        'Dim cmd As String
        Dim ds As New DataSet
        'Dim objConn As SqlConnection
        'objConn = New SqlConnection(connstr)
        'objConn.Open()

        If txt_itemcode.Text <> "" Then
            Dim cmd As String =
<code>
    <![CDATA[

                select distinct
	                itm.[Description]
                from [Reports].[dbo].[t_item_master] as itm with (nolock)

                left outer join [TOMR].[dbo].[TO_MR_Invt_Qty] as qty with (nolock)
	                on itm.[Company Ownership Code] = qty.[Item Owner]
	                and itm.[Item No] = qty.[Item No]

                where itm.[Company Ownership Code] = ']]><%= Session("tmpco") %><![CDATA[' 
                    and Replace(itm.[Brand Description], '''', '') = ']]><%= Replace(Replace(Session("tmpbrand"), "&#39;", ""), "amp;", "") %><![CDATA['
	                and qty.[Whse Qty] <> '0'
	                and itm.[Item No] not in (
		                select distinct [ITEM NO]
		                from [TOMR].[dbo].[]]><%= Strings.Replace(Session("tmpdoc"), "-", "") %><![CDATA[] with (nolock)
	                )
                    and itm.[Item No] like '%]]><%= txt_itemcode.Text %><![CDATA[%'
                order by itm.[Description]

            ]]>
</code>.Value

            ds = executeQuery(cmd)
            dropdownItem.Items.Clear()

            If ds.Tables(0).Rows.Count > 0 Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dropdownItem.Items.Add(ds.Tables(0).Rows(i)("Description"))
                Next
            End If
        End If

        'Dim sqlcmd As New SqlCommand(sqlscript, objConn)

        'Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
        'dropdownItem.Items.Clear()
        'Do While sqlreader.Read
        'dropdownItem.Items.Add(sqlreader("Description"))
        'Loop

        'sqlreader.Close()
        'objConn.Close()
        'objConn.Dispose()

        dropdownItem.Items.Insert(0, New ListItem("-- Select Item --", "0"))


    End Sub
    Sub LoadItems_init()
        'Dim cmd As String
        Dim ds As New DataSet
        'Dim objConn As SqlConnection
        'objConn = New SqlConnection(connstr)
        'objConn.Open()

        Dim cmd As String =
<code>
    <![CDATA[

                select distinct
	                itm.[Description]
                from [Reports].[dbo].[t_item_master] as itm with (nolock)

                left outer join [TOMR].[dbo].[TO_MR_Invt_Qty] as qty with (nolock)
	                on itm.[Company Ownership Code] = qty.[Item Owner]
	                and itm.[Item No] = qty.[Item No]

                where itm.[Company Ownership Code] = ']]><%= Session("tmpco") %><![CDATA[' 
                    and Replace(itm.[Brand Description], '''', '') = ']]><%= Strings.Replace(Session("tmpbrand"), "&#39;", "") %><![CDATA['
	                and qty.[Whse Qty] <> '0'
	                and itm.[Item No] not in (
		                select distinct [ITEM NO]
		                from [TOMR].[dbo].[]]><%= Strings.Replace(Session("tmpdoc"), "-", "") %><![CDATA[] with (nolock)
	                )
                order by itm.[Description]

            ]]>
</code>.Value

            ds = executeQuery(cmd)
            dropdownItem.Items.Clear()

        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                dropdownItem.Items.Add(ds.Tables(0).Rows(i)("Description"))
            Next
        End If

        'Dim sqlcmd As New SqlCommand(sqlscript, objConn)

        'Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
        'dropdownItem.Items.Clear()
        'Do While sqlreader.Read
        'dropdownItem.Items.Add(sqlreader("Description"))
        'Loop

        'sqlreader.Close()
        'objConn.Close()
        'objConn.Dispose()

        dropdownItem.Items.Insert(0, New ListItem("-- Select Item --", "0"))


    End Sub

    Private Sub CreateGrid()

        Dim ii As Integer
        GridView1.Columns.Clear()
        If GridView1.Columns.Count > 1 Then
            For ii = 1 To GridView1.Columns.Count - 1
                'GridView1.Columns.RemoveAt(1)
            Next
        End If

        GridView1.Columns.Clear()

        Dim nameColumn As BoundField

        nameColumn = New BoundField()


        nameColumn.HeaderText = "No."
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        nameColumn.ControlStyle.Width = "10"
        GridView1.Columns.Add(nameColumn)


        nameColumn = New BoundField()
        nameColumn.DataField = "Document No"
        nameColumn.HeaderText = "DOCUMENT NO."
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)


        nameColumn = New BoundField()
        nameColumn.DataField = "Item No"
        nameColumn.HeaderText = "BARCODE"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Item Description"
        nameColumn.HeaderText = "DESCRIPTION"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Whse Qty"
        nameColumn.HeaderText = "WHSE QTY"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Total Req"
        nameColumn.HeaderText = "TOTAL STORE REQUEST"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Avail Qty"
        nameColumn.HeaderText = "QTY AVAILABLE"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Item Owner"
        nameColumn.HeaderText = "COMPANY"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Brand"
        nameColumn.HeaderText = "BRAND"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "MAIN CATEGORY"
        nameColumn.HeaderText = "MAIN CATEGORY"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "SRP"
        nameColumn.HeaderText = "SRP"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        GridView1.Columns.Add(nameColumn)

        dsloc = Me.GetLocations()
        For Each drloc As DataRow In dsloc.Tables(0).Rows

            nameColumn = New BoundField()
            nameColumn.DataField = drloc("Location Name")
            nameColumn.HeaderText = drloc("Location Name")
            nameColumn.ControlStyle.Width = "100"
            GridView1.Columns.Add(nameColumn)

        Next

        'Dim ef As CommandField = New CommandField()
        'ef.ButtonType = ButtonType.Button
        'ef.ShowEditButton = True
        'ef.EditText = "Edit Line"
        'ef.ShowCancelButton = True
        'ef.CancelText = "Cancel Edit"
        'ef.UpdateText = "Update Line"
        'GridView1.Columns.Add(ef)


    End Sub

    Private Sub BindGrid(tmpstr As String)
        Dim sConnectionString As String
        Dim objConn As SqlConnection
        sConnectionString = [TO].My.Settings.SQLConnection.ToString
        objConn = New SqlConnection(sConnectionString)
        objConn.Open()

        Dim sqlscript As String =
        <code>
            <![CDATA[

                declare @qtyscript as nvarchar(MAX)
                declare @tmpcat as nvarchar(50)

                set @tmpcat = ']]><%= tmpstr %><![CDATA['
                
                create table #tmp1
                (
					[ID] [int] identity(1,1) NOT NULL,
					[Document No] [nvarchar](50) NULL,
					[Item No] [nvarchar](20) NULL,
					[Item Description] [nvarchar](250) NULL,
					[Item Owner] [nvarchar](10) NULL,
					[Brand] [nvarchar](50) NULL,
					[Main Category] [nvarchar](50) NULL,
                    [SRP] [numeric](20,2) NULL,
                    [Price Type] [nvarchar](20) NULL,
					[Location Code] [nvarchar](20) NULL,
					[Location Name] [nvarchar](200) NULL,
                    [Location Type] [nvarchar](100) NULL,
					[Request Qty] [int] NULL,
					[Doc No] [nvarchar](50) NULL,
					[Doc Date] [date] NULL,
					[Remarks] [nvarchar](250) NULL,
					[Status] [nvarchar](25) NULL,
					[Whse Qty] [int] NULL,
					[Avail Qty] [int] NULL
                )

                insert into #tmp1
                select 
					a.[Document No]
					, a.[Item No]
					, a.[Item Description]
					, a.[Item Owner]
					, a.[Brand]
					, a.[Main Category]
                    , a.[SRP]
                    , a.[Price Type]
					, a.[Location Code]
					, a.[Location Name]
                    , a.[Location Type]
					, a.[Request Qty]
					, a.[Doc No]
					, a.[Doc Date]
					, a.[Remarks]
					, a.[Status]
					, ISNULL(wqty.[Whse Qty], '0') as [Whse Qty]
					, ISNULL((ISNULL(wqty.[Whse Qty], '0') - ISNULL(b.[Req Qty], '0')), '0') as [Avail Qty]
                from [dbo].[TO_MR_Collated_Lines] as a with (nolock)
				left outer join 
					( select 
						[Item Owner]
						, [Item No]
						, SUM([Whse Qty]) as [Whse Qty]
						from [dbo].[TO_MR_Invt_Qty] with (nolock)
						group by [Item Owner], [Item No]
					) as wqty 
					on a.[Item Owner] = wqty.[Item Owner]
					and a.[Item No] = wqty.[Item No]
				left outer join 
					( select 
                        [Document No]
						, [Item No]
						, SUM([Request Qty]) as [Req Qty]
						from [dbo].[TO_MR_Collated_Lines] with (nolock)
						--where [Status] = 'IN-PROCESS'
						group by [Document No], [Item No]
					) as b
					on a.[Document No] = b.[Document No]
					and a.[Item No] = b.[Item No]
                where 
	                a.[Document No] = ']]><%= Session("tmpdoc") %><![CDATA['
                    and a.[Main Category] like (
					CASE
					WHEN @tmpcat = 'ALL' THEN '%'
					ELSE @tmpcat
					END)

                DECLARE @cols AS NVARCHAR(MAX),
                    @query  AS NVARCHAR(MAX);

                SET @cols = STUFF((select ',' + QUOTENAME(tmp1.[Location Name]) 
                            FROM #tmp1 as tmp1
							LEFT OUTER JOIN [dbo].[TO_MR_Store_Sort] as store with (nolock)
								on tmp1.[Location Code] = store.[Location Code]
                            where tmp1.[Document No] = ']]><%= Session("tmpdoc") %><![CDATA['
							GROUP BY tmp1.[Location Name], store.[Ctrl No]
							ORDER BY store.[Ctrl No] ASC
                            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '')

                set @query = 
	                '
	                SELECT
						[Document No]
						, [Item No]
						, [Item Description]
						, [Item Owner]
						, [Brand]
						, [Main Category]
                        , [SRP]
						, [Price Type]
						, [Whse Qty]
						, [Total Req]
						, [Avail Qty]
						, [Gender]
						, [Search Description]
						, [Color]
						, [Size]
		                , ' + @cols + ' 
	                from 
		                (
		                select
							aa.[Document No]
							, aa.[Item No]
							, aa.[Item Description]
							, aa.[Item Owner]
							, aa.[Brand]
							, aa.[Main Category]
                            , aa.[SRP]
							, aa.[Price Type]
							, aa.[Location Name]
							, bb.[Total Req]
							, aa.[Request Qty]
							, aa.[Whse Qty]
							, aa.[Avail Qty]
							, itm.[Gender]
							, itm.[Search Description]
							, itm.[Color]
							, itm.[Size]
		                from #tmp1 as aa
						left outer join 
							( select
								[Item No]
								, SUM([Request Qty]) as [Total Req]
							from #tmp1
							group by [Item No]
							) as bb
							on aa.[Item No] = bb.[Item No]
						left outer join [Reports].[dbo].[t_item_master] as itm with (nolock)
							on aa.[Item No] = itm.[Item No]
						group by 
							aa.[Document No]
							, aa.[Item No]
							, aa.[Item Description]
							, aa.[Item Owner]
							, aa.[Brand]
							, aa.[Main Category]
                            , aa.[SRP]
							, aa.[Price Type]
							, aa.[Location Name]
							, bb.[Total Req]
							, aa.[Request Qty]
							, aa.[Whse Qty]
							, aa.[Avail Qty]
							, itm.[Gender]
							, itm.[Search Description]
							, itm.[Color]
							, itm.[Size]
		                ) as x

	                pivot 
		                (
			                sum([Request Qty])
			                for [Location Name] in (' + @cols + ')
		                ) as p 
										
					order by
						[Brand] ASC
						, [Main Category] ASC
						, [Gender] ASC
						, [Search Description] ASC
						, [Color] ASC
						, [Size] ASC
						, [Price Type] ASC
						, [SRP] ASC
						, [Item No] ASC

	                '

                execute(@query)

                drop table #tmp1

            ]]>
        </code>.Value

        Dim sqlcmd As New SqlCommand(sqlscript, objConn)
        GridView1.DataSource = sqlcmd.ExecuteReader()
        GridView1.DataBind()

        objConn.Close()
        objConn.Dispose()

    End Sub

    Private Function GetColumns() As DataSet
        Dim sqlscript As String = "exec sp_collect_columns '" & Replace(Session("selectedDoc"), "-", "") & "','',''"
        Dim cmd As New SqlCommand(sqlscript)
        Using con As New SqlConnection(connstr)
            Using sda As New SqlDataAdapter()
                cmd.CommandTimeout = 0
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

    Private Sub CreateGrid2()

        Dim stores As New List(Of String)()
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()
        Dim cmd As String =
            <code>
                <![CDATA[
                    select distinct
                        b.[Location Name]
                    from [dbo].[TO_MR_Collated_Header_2] as a with (nolock)
                    left outer join [Reports].[dbo].[t_locations2] as b with (nolock)
                        on LEFT(a.[Doc No], 9) = b.[Location Code]
                        and a.[Company] = b.[Company Code]
                    where LTRIM(RTRIM(a.[Remarks])) <> '' and
                         a.[Document No] = ']]><%= Session("tmpdoc") %><![CDATA['
                ]]>
            </code>.Value
        sqlcmd = New SqlCommand(cmd, objConn)
        Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
        If sqlreader.HasRows = True Then
            While sqlreader.Read()
                stores.Add(sqlreader("Location Name").ToString())
            End While
        End If
        objConn.Close()
        objConn.Dispose()


        GridView1.Columns.Clear()

        Dim nameColumn As BoundField
        'Dim colcheckbox As CheckBoxField

        dsloc = Me.GetColumns()
        For Each drloc As DataRow In dsloc.Tables(0).Rows
            'colcheckbox = New CheckBoxField
            nameColumn = New BoundField()

            'colcheckbox.HeaderText = "Select"

            nameColumn.DataField = drloc("COLUMN_NAME")
            nameColumn.HeaderText = drloc("COLUMN_NAME")
            'nameColumn.ControlStyle.Width = "100"

            Dim val As String = stores.Find(Function(value As String)
                                                Return value.Equals(drloc("COLUMN_NAME"))
                                            End Function)

            If val <> "" Then nameColumn.HeaderStyle.ForeColor = Color.Red


            GridView1.Columns.Add(nameColumn)

        Next

    End Sub

    Private Sub BindGrid2(tmpstr As String)
        Dim sConnectionString As String
        Dim objConn As SqlConnection
        sConnectionString = [TO].My.Settings.SQLConnection.ToString
        objConn = New SqlConnection(sConnectionString)
        objConn.Open()

        Dim sqlscript As String
        If Session("mode") = "NAV ON QUEUE" Or Session("mode") = "CANCELLED" Then
            sqlscript =
        <code>
            <![CDATA[

            DECLARE @docno AS NVARCHAR(50) = replace(']]><%= Session("tmpdoc") %><![CDATA[','-','')
                  DECLARE @DYNAMICCOLS AS NVARCHAR(max)
                  DECLARE @query AS NVARCHAR(MAX)
				  
                  SET @DYNAMICCOLS = stuff((select ',' + '[' + column_name + ']' from INFORMATION_SCHEMA.COLUMNS 
                  where TABLE_NAME = @docno and ORDINAL_POSITION >= 12 FOR XML PATH('')),1,1,'')

                  set @query = '
				  declare @tmpcat as nvarchar(50) = ''ALL''
                  select
                  a.[NO]
                  ,a.[DOCUMENT NO]
                  ,a.[ITEM NO]
                  ,a.[ITEM DESCRIPTION]
                  ,a.[WHSE QTY]
                  ,a.[TOTAL REQ]
                  ,a.[QTY AVAIL]
                  ,a.[ITEM OWNER]
                  ,a.[BRAND]
                  ,a.[MAIN CATEGORY]
				  ,a.[SRP]
                  , (select [Year-Season] from [reports].[dbo].[t_item_master] where [ITEM NO] = a.[ITEM NO]) as [YEAR-SEASON]
                  , (select [Search Description] from [reports].[dbo].[t_item_master] where [ITEM NO] = a.[ITEM NO]) as [SEARCH DESCRIPTION]
                  ,' + @DYNAMICCOLS  + '
                  into #tmp1
                  FROM ' + @docno + ' as a

                  select * from #tmp1 where [MAIN CATEGORY] like (
                       CASE
		               WHEN @tmpcat = ''ALL'' THEN ''%''
		               ELSE @tmpcat
		               END) order by [NO]
                  drop table #tmp1
                  '
                  exec(@query)

            ]]>
        </code>.Value
        Else
            sqlscript =
        <code>
            <![CDATA[

                declare @tablename as nvarchar(100) = REPLACE(']]><%= Session("tmpdoc") %><![CDATA[', '-', '')
                
                select 
	            a.[COLUMN_NAME]
                into #tmpstores
                from INFORMATION_SCHEMA.COLUMNS as a
                left outer join
	            ( select distinct
		        [Location Name]
	            from [Reports].[dbo].[t_locations2] with (nolock)
	            ) as loc
	            on a.[COLUMN_NAME] = loc.[Location Name]
                where TABLE_NAME = @tablename
	            and loc.[Location Name] is not null

                ---------------------------------------------------------------------------------

                declare @cols as nvarchar(max)
                set @cols = STUFF((select '+' + 'ISNULL(' + QUOTENAME([COLUMN_NAME]) + ', ''0'')'
                    FROM #tmpstores 
                    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '')

                -----------------------------------------------------------------------------------

                declare @sqlreq as nvarchar(max)
                set @sqlreq = '
	                select
		                [Item No]
		                , SUM([Whse Qty]) as [Whse Qty]
	                into #tmpinvt
	                from [TOMR].[dbo].[TO_MR_Invt_Qty]
	                where [Item No] in 
		                (select [ITEM NO] from [dbo].[' + @tablename + '])
                        --and [Location Code] = (select top 1 [Warehouse Code] from [TOMR].[dbo].[TO_MR_Warehouse] where [Company Ownership Code] = '']]><%= Session("tmpco") %><![CDATA['')
	                group by [Item No]

	                select 
		                [ITEM NO], ' + @cols + ' as [TOTAL REQ]
	                into #tmpreq
	                from [dbo].[' + @tablename + ']

	                update a
		                set a.[WHSE QTY] = ISNULL(b.[WHSE QTY], ''0'')
		                , a.[TOTAL REQ] = ISNULL(c.[TOTAL REQ], ''0'')
		                , a.[QTY AVAIL] = (ISNULL(b.[WHSE QTY], ''0'') - ISNULL(c.[TOTAL REQ], ''0''))
	                from [TOMR].[dbo].[' + @tablename + '] as a
	                left outer join #tmpinvt as b
		                on a.[ITEM NO] = b.[Item No]
	                left outer join #tmpreq as c
		                on a.[ITEM NO] = c.[ITEM NO]
	                drop table #tmpinvt
	                drop table #tmpreq
	                '
                exec(@sqlreq)

                drop table #tmpstores

                ------------------------------------------------------------------------------------------

                  DECLARE @docno AS NVARCHAR(50) = replace(']]><%= Session("tmpdoc") %><![CDATA[','-','')
                  DECLARE @DYNAMICCOLS AS NVARCHAR(max)
                  DECLARE @query AS NVARCHAR(MAX)
				  
                  SET @DYNAMICCOLS = stuff((select ',' + '[' + column_name + ']' from INFORMATION_SCHEMA.COLUMNS 
                  where TABLE_NAME = @docno and ORDINAL_POSITION >= 12 FOR XML PATH('')),1,1,'')

                  set @query = '
				  declare @tmpcat as nvarchar(50) = '']]><%= tmpstr %><![CDATA[''
                  select
                      a.[NO]
                      ,a.[DOCUMENT NO]
                      ,a.[ITEM NO]
                      ,a.[ITEM DESCRIPTION]
                      ,a.[WHSE QTY]
                      ,a.[TOTAL REQ]
                      ,a.[QTY AVAIL]
                      ,a.[ITEM OWNER]
                      ,a.[BRAND]
                      ,a.[MAIN CATEGORY]
				      ,a.[SRP]
                      , (select [Year-Season] from [reports].[dbo].[t_item_master] where [ITEM NO] = a.[ITEM NO]) as [YEAR-SEASON]
                      , (select [Search Description] from [reports].[dbo].[t_item_master] where [ITEM NO] = a.[ITEM NO]) as [SEARCH DESCRIPTION]
                      ,' + replace(@DYNAMICCOLS,'&amp;','&')  + '
                  into #tmp1
                  FROM ' + @docno + ' as a

                  select * from #tmp1 where [MAIN CATEGORY] like (
                       CASE
		               WHEN @tmpcat = ''ALL'' THEN ''%''
		               ELSE @tmpcat
		               END) order by [NO]
                  drop table #tmp1
                  '
                  exec(@query)
                 

  

            ]]>
        </code>.Value
        End If



        Dim sqlcmd As New SqlCommand(sqlscript, objConn)
        GridView1.DataSource = sqlcmd.ExecuteReader()
        GridView1.DataBind()



        objConn.Close()
        objConn.Dispose()
        LoadItems_init()
    End Sub

    Protected Sub gv_RowDataBound(sender As Object, e As GridViewRowEventArgs)


        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.ToolTip = e.Row.Cells(3).Text & vbCr & e.Row.Cells(4).Text & vbCr & "WHSE QTY:  " & e.Row.Cells(5).Text & vbCr & "TOTAL REQ: " & e.Row.Cells(6).Text & vbCr & "QTY AVAIL:  " & e.Row.Cells(7).Text
            Dim availcell As TableCell = e.Row.Cells(7)
            Dim availqty As Integer = Integer.Parse(availcell.Text)
            If availqty < 0 Then availcell.BackColor = ColorTranslator.FromHtml("#ff8080")
            'If availqty < 0 Then availcell.ForeColor = ColorTranslator.FromHtml("#ff8080")
            'If availqty < 0 Then availcell.Font.Bold = True


            e.Row.Cells(0).CssClass = "Locked"
            e.Row.Cells(1).CssClass = "Locked"
            e.Row.Cells(2).CssClass = "Locked"
            e.Row.Cells(3).CssClass = "Locked"
        End If

        If Session("mode") = "NAV ON QUEUE" Or Session("mode") = "CANCELLED" Then
            e.Row.Cells(0).Visible = False
            upload_btn.Visible = False
        End If

    End Sub

    Protected Sub gv_RowCreated(sender As Object, e As GridViewRowEventArgs)

        Dim ix As Integer = 0

        If Session("tmpviewedit") = "VIEW" Then
            Dim row As Integer = 0
            row = GridView1.Rows.Count - 1
            If row > -1 Then
                Dim lblCtr As New Label
                lblCtr.Text = row + 1
                GridView1.Rows(row).Cells(0).Controls.Add(lblCtr)
            End If
        End If

        If Session("tmpviewedit") = "EDIT" Then

            Dim tmpid As Integer = 0
            Dim gvcolcnt As Integer = GridView1.Columns.Count - 1
            Dim row As Integer = 0
            Dim col As Integer

            row = GridView1.Rows.Count - 1
            If row > -1 Then
                'Dim lblCtr As New Label
                'lblCtr.Text = row + 1
                'GridView1.Rows(row).Cells(0).Controls.Add(lblCtr)

                col = 13
                Do While gvcolcnt + 1 > col
                    col = col + 1
                    tmpid = tmpid + 1
                    tmpTextbox = New TextBox
                    lbl = New Label
                    'lbl2 = New Label

                    lbl.ID = "lbl_" & ix
                    'lbl2.ID = "lbl2_" & ix
                    ix += 1

                    ltr = New LiteralControl
                    'ltr2 = New LiteralControl
                    tmpTextbox.Width = "100"


                    If [TO].My.Settings.ViewDetails = 1 Then

                        lbl.Text = "OH: --"
                        'lbl2.Text = "Ave Sales : --"

                        Dim objConn As SqlConnection = New SqlConnection(connstr)
                        Dim sqlcmd As SqlCommand
                        objConn.Open()
                        Dim cmd As String =
                           <code>
                               <![CDATA[
                    Select top 1
                    ISNULL([Store Qty], '0') as [Store Qty]
                    From [dbo].[TO_MR_Invt_Sales]
                    where [Location Name] = ']]><%= GridView1.Columns(col - 1).HeaderText %><![CDATA['
                                   And [Item No] = ']]><%= GridView1.Rows(row).Cells(3).Text %><![CDATA[' --and [Store Qty] >= '0'
                          ]]>
                           </code>.Value

                        sqlcmd = New SqlCommand(cmd, objConn)
                        sqlcmd.ExecuteNonQuery()

                        Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader

                        If sqlreader.HasRows = True Then
                            Do While sqlreader.Read
                                If Val(sqlreader("Store Qty").ToString) < 0 Then
                                    lbl.Text = "<i class='onhand_css'>" & 0 & "</i>"
                                Else
                                    lbl.Text = "<i class='onhand_css'>" & sqlreader("Store Qty").ToString & "</i>"
                                    'lbl2.Text = "Ave Sales : " & sqlreader("Store Sales").ToString
                                End If
                            Loop
                        Else
                            lbl.Text = "<i class='onhand_css'>" & "0" & "</i>"
                            'lbl2.Text = "Ave Sales : 0.0"
                        End If

                        objConn.Close()

                    End If

                    tmpTextbox.ID = "txt" & tmpid.ToString

                    'tmpTextbox.Attributes.Add("onkeypress", "javascript:tab isNumberKey(event);")

                    ltr.Text = ""
                    'ltr2.Text = " <br/> "
                    GridView1.Rows(row).Cells(col).Controls.Add(tmpTextbox)
                    GridView1.Rows(row).Cells(col).Controls.Add(ltr)
                    GridView1.Rows(row).Cells(col).Controls.Add(lbl)
                    'GridView1.Rows(row).Cells(col).Controls.Add(ltr2)
                    'GridView1.Rows(row).Cells(col).Controls.Add(lbl2)
                    GridView1.Rows(row).Cells(3).BackColor = ColorTranslator.FromHtml("#e6f2ff")
                    GridView1.Rows(row).Cells(3).CssClass = "FixedCol"

                    'GridView1.Rows(row).Cells(5).BackColor = ColorTranslator.FromHtml("#e6f2ff")
                    If GridView1.Rows(row).Cells(col).Text = "&nbsp;" Or Trim(GridView1.Rows(row).Cells(col).Text) = "" Then
                        'tmpTextbox.Text = tmpTextbox.ID.ToString
                    tmpTextbox.Text = ""
                        tmpTextbox.Style.Add("font-family", "Verdana, Geneva, Tahoma, sans-serif")
                        tmpTextbox.Style.Add("font-size", "10px")
                        tmpTextbox.Style.Add("font-weight", "bold")
                        tmpTextbox.Style.Add("text-align", "center")
                        tmpTextbox.Style.Add("background-color", "#e6f2ff")
                        tmpTextbox.Style.Add("border", "solid 1px lightgray")
                        tmpTextbox.Style.Add("width", "80px")
                    ElseIf CInt(GridView1.Rows(row).Cells(col).Text) = 0 Then
                        tmpTextbox.Text = ""
                        tmpTextbox.Style.Add("font-family", "Verdana, Geneva, Tahoma, sans-serif")
                        tmpTextbox.Style.Add("font-size", "10px")
                        tmpTextbox.Style.Add("font-weight", "bold")
                        tmpTextbox.Style.Add("text-align", "center")
                        tmpTextbox.Style.Add("background-color", "#e6f2ff")
                        tmpTextbox.Style.Add("border", "solid 1px lightgray")
                        tmpTextbox.Style.Add("width", "80px")
                        'tmpTextbox.Style.Add("border-radius", "5px")
                    Else
                        'tmpTextbox.Text = tmpTextbox.ID.ToString
                        tmpTextbox.Style.Add("font-family", "Verdana, Geneva, Tahoma, sans-serif")
                        tmpTextbox.Style.Add("font-size", "10px")
                        tmpTextbox.Style.Add("font-weight", "bold")
                        tmpTextbox.Style.Add("text-align", "center")
                        tmpTextbox.Style.Add("border", "solid 1px lightgray")
                        tmpTextbox.Style.Add("width", "80px")
                        tmpTextbox.Text = GridView1.Rows(row).Cells(col).Text
                        GridView1.Rows(row).Cells(4).Width = 300

                    End If

                Loop
            End If

        End If

        'If e.Row.RowType = DataControlRowType.DataRow Then
        '    e.Row.Attributes("ondblclick") = Page.ClientScript.GetPostBackClientHyperlink(GridView1, "Select$" & e.Row.RowIndex)
        '    e.Row.Attributes("style") = "cursor:pointer"
        'End If

    End Sub

    Protected Sub gv_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)

        '

    End Sub

    Protected Sub gv_CancelEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        GridView1.EditIndex = -1
        BindGrid2(dropdownMainCat.Text)
    End Sub

    Protected Sub gv_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        '
    End Sub

    Protected Sub gv_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)


        'Dim selectedDoc As String = GridView1.Rows(e.RowIndex).Cells(2).Text
        'Dim selectedItem As String = GridView1.Rows(e.RowIndex).Cells(3).Text
        'Dim selectedBrand As String = GridView1.Rows(e.RowIndex).Cells(9).Text

        Session("selected_doc") = GridView1.Rows(e.RowIndex).Cells(2).Text
        Session("selected_item") = GridView1.Rows(e.RowIndex).Cells(3).Text
        Session("selected_brand") = GridView1.Rows(e.RowIndex).Cells(9).Text
        Session("deleteType") = "item"

        msgbox2_label.Text = "Are you sure you want to delete this item?"
        div_message_box.Visible = True


    End Sub

    Protected Sub gv_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles GridView1.SelectedIndexChanged

        'tmplabel.Text = GridView1.SelectedRow.Cells(1).Text

        Dim gvr As GridViewRow
        Dim hdr As GridViewRow

        gvr = GridView1.SelectedRow
        hdr = GridView1.HeaderRow
        Dim item As String = gvr.Cells(3).Text

        Dim _lbl As New Label
        'Dim _lbl2 As New Label


        If [TO].My.Settings.ViewDetails = 0 Then

            For i = 14 To gvr.Cells.Count - 1
                Dim objConn As SqlConnection = New SqlConnection(connstr)
                Dim sqlcmd As SqlCommand
                objConn.Open()
                Dim cmd As String =
                        <code>
                            <![CDATA[
                                select top 1
                                 ISNULL([Store Qty], '0') as [Store Qty]
                                 , ISNULL([Store Sales], '0') as [Store Sales]
                                from [dbo].[TO_MR_Invt_Sales]
                                where [Location Name] = ']]><%= hdr.Cells(i).Text %><![CDATA['
                                    and [Item No] = ']]><%= gvr.Cells(3).Text %><![CDATA['
                            ]]>
                        </code>.Value

                sqlcmd = New SqlCommand(cmd, objConn)
                sqlcmd.ExecuteNonQuery()

                Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
                _lbl = CType(gvr.Cells(i).FindControl("lbl_" & i - 14), Label)
                '_lbl2 = CType(gvr.Cells(i).FindControl("lbl2_" & i - 11), Label)
                'CType(e.Row.Cells(1).FindControl("PayRateAmount"), Label).Text
                If _lbl Is Nothing Then
                Else

                    If sqlreader.HasRows = True Then
                        Do While sqlreader.Read
                            _lbl.Text = "On Hand :  " & sqlreader("Store Qty").ToString
                            '_lbl2.Text = "Ave Sales : " & sqlreader("Store Sales").ToString

                        Loop
                    Else
                        _lbl.Text = "On Hand : 0"
                        '_lbl2.Text = "Ave Sales : 0.0"
                    End If
                End If


                objConn.Close()
                objConn.Dispose()



                'ltr.Text = " <br/> "
                'ltr2.Text = " <br/> "
                'gvr.Cells(i).Controls.Add(ltr)
                'gvr.Cells(i).Controls.Add(lbl)
                'gvr.Cells(i).Controls.Add(ltr2)
                'gvr.Cells(i).Controls.Add(lbl2)

            Next



            'tmpTextbox.ID = "txt" & tmpid.ToString

        End If

        'lblItemNo.Text = "Item No. : " + GridView1.SelectedRow.Cells(1).Text

        'AddTextbox()

    End Sub

    Protected Sub gv_SelectedIndexChanging(ByVal sender As Object, ByVal e As GridViewSelectEventArgs)
        '
    End Sub

    Protected Sub AddItem(tmpitem As String)

        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String =
            <code>
                <![CDATA[

                    update [dbo].[]]><%= Strings.Replace(Session("tmpdoc"), "-", "") %><![CDATA[]
                    set [NO] = [NO] + 1
                
                    insert into [dbo].[]]><%= Strings.Replace(Session("tmpdoc"), "-", "") %><![CDATA[]
                    ([NO], [DOCUMENT NO], [ITEM NO], [ITEM DESCRIPTION], [WHSE QTY], [TOTAL REQ], [QTY AVAIL], [ITEM OWNER], [BRAND], [MAIN CATEGORY], [SRP])
                    select
	                    '1'
	                    , ']]><%= Session("tmpdoc") %><![CDATA['
	                    , itm.[Item No]
	                    , ']]><%= tmpitem %><![CDATA['
	                    , (select sum([Whse Qty]) from [dbo].[TO_MR_Invt_Qty] where [Item No] = itm.[Item No])
	                    , '0'
	                    , (select sum([Whse Qty]) from [dbo].[TO_MR_Invt_Qty] where [Item No] = itm.[Item No])
	                    , itm.[Company Ownership Code]
                        , itm.[Brand Description]
	                    , itm.[MainCat Description]
	                    , itm.[Unit Price]
                    from [Reports].[dbo].[t_item_master] as itm with (nolock)

                    where [Description] = ']]><%= tmpitem %><![CDATA['
                       
                ]]>
            </code>.Value

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()

        GridView1.EditIndex = -1
        BindGrid2(dropdownMainCat.Text)

        LoadItems()

        '(select count([NO]) from [TOMR].[dbo].[]]><%= Strings.Replace(Session("tmpdoc"), "-", "") %><![CDATA[]) + 1

    End Sub

    'Protected Sub buttonAdd_Click(sender As Object, e As EventArgs) Handles buttonAdd.Click
    '    If dropdownItem.Text <> "0" Then
    '        AddItem(dropdownItem.Text)
    '    End If
    '    'LoadItems()
    'End Sub

    Private Sub dropdownMainCat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dropdownMainCat.SelectedIndexChanged
        'GridView1.EditIndex = -1
        '

        CreateGrid2()
        BindGrid2(dropdownMainCat.Text)
    End Sub

    Private Sub AddTextbox()

        ''Dim tmpTextbox As TextBox
        'Dim tmpid As integer = 0
        'Dim gvcolcnt As integer = GridView1.Columns.Count - 1
        'Dim gvrowcnt As integer = GridView1.Rows.Count - 1
        'Dim row As integer = 0
        'Dim col As integer

        'row = -1
        'Do While gvrowcnt > row
        '    row = row + 1
        '    col = 10
        '    Do While gvcolcnt > col
        '        col = col + 1
        '        tmpid = tmpid + 1
        '        tmpTextbox = New TextBox()
        '        tmpTextbox.ID = "txt" + tmpid.ToString
        '        'tmpTextbox.AutoPostBack = True
        '        GridView1.Rows(row).Cells(col).Controls.Add(tmpTextbox)
        '        If GridView1.Rows(row).Cells(col).Text = "&nbsp;" Then
        '            tmpTextbox.Text = "0"
        '        Else
        '            tmpTextbox.Text = GridView1.Rows(row).Cells(col).Text
        '        End If
        '    Loop
        'Loop

    End Sub

    Private Sub AssociateTextboxEventHandler()
        'For Each c As Control In Me.Controls
        '    If TypeOf c Is TextBox Then
        '        AddHandler c, AddressOf tb_TextChanged
        '    End If
        'Next
    End Sub

    Private Sub tb_TextChanged()
        '
    End Sub

    Protected Sub buttonSave_Click(sender As Object, e As EventArgs) Handles buttonSave.Click

        Session("deletionType") = ""
        Session("negativeType") = ""
        Session("deleteType") = ""
        Session("isEnter") = "yes"
        msgbox2_label.Text = "Save the document now?"
        button1.Text = "Save"
        div_message_box.Visible = True
        'button1.UseSubmitBehavior = True

        'Dim objConn As SqlConnection = New SqlConnection(connstr)
        'Dim sqlcmd As SqlCommand
        'objConn.Open()

        'Dim tmpid As Integer = 0
        'Dim gvcolcnt As Integer = GridView1.Columns.Count - 1
        'Dim gvrowcnt As Integer = GridView1.Rows.Count - 1
        'Dim row As Integer = 0
        'Dim col As Integer
        'Dim reqt As Integer = 0
        'Dim availt As Integer = 0


        'Try
        '    row = -1
        '    Do While gvrowcnt > row
        '        Dim colvalue As Integer = 0
        '        tmpid = 0
        '        row = row + 1
        '        col = 13
        '        Do While gvcolcnt + 1 > col
        '            col = col + 1
        '            tmpid = tmpid + 1

        '            If Trim(CType(GridView1.Rows(row).Cells(col).FindControl("txt" + tmpid.ToString), TextBox).Text) = Nothing Then
        '                colvalue = 0
        '            Else
        '                colvalue = Convert.ToInt64(CType(GridView1.Rows(row).Cells(col).FindControl("txt" + tmpid.ToString), TextBox).Text)
        '            End If

        '            Dim a As Integer = 13
        '            reqt += colvalue

        '            a = a + 1

        '            availt = Convert.ToInt64(Trim(GridView1.Rows(row).Cells(5).Text)) - reqt

        '            Dim cmd As String =
        '                        <code>
        '                            <![CDATA[
        '                    BEGIN TRANSACTION;
        '                        UPDATE [dbo].]]><%= Strings.Replace(Session("tmpdoc"), "-", "") %><![CDATA[
        '                        SET
        '                            [TOTAL REQ] = ']]><%= reqt %><![CDATA['
        '                            , [QTY AVAIL] =  ']]><%= availt %><![CDATA['
        '                            , []]><%= Trim(GridView1.Columns(col - 1).HeaderText) %><![CDATA[] = ']]><%= colvalue %><![CDATA['
        '                        WHERE
        '                            [Item No] = ']]><%= Trim(GridView1.Rows(row).Cells(3).Text) %><![CDATA['
        '                    COMMIT TRANSACTION;
        '                    ]]>
        '                        </code>.Value

        '            sqlcmd = New SqlCommand(cmd, objConn)
        '            sqlcmd.ExecuteNonQuery()

        '        Loop
        '        reqt = 0
        '        availt = 0
        '    Loop

        '    objConn.Close()
        '    objConn.Dispose()


        '    Dim cmd_docdetail_updateQty_select As String
        '    Dim ds_docdetail_updateQty_select As New DataSet

        '    cmd_docdetail_updateQty_select = "EXEC sp_selectDocDetail 0,'" & Session("tmpdoc") & "',''"
        '    ds_docdetail_updateQty_select = executeQuery(cmd_docdetail_updateQty_select)

        '    If ds_docdetail_updateQty_select.Tables(0).Rows.Count > 0 Then
        '        Dim cmd_docdetail_updateQty As String
        '        Dim ds_docdetail_updateQty As New DataSet
        '        For i = 0 To ds_docdetail_updateQty_select.Tables(0).Rows.Count - 1
        '            cmd_docdetail_updateQty = "update [dbo].[DocDetail] set [ALLOCATEDQTY] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ALLOCATEDQTY") & "'
        '                                    where [DOCNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("Doc No") & "'
        '                                   and [ITEMNO] = '" & ds_docdetail_updateQty_select.Tables(0).Rows(i)("ITEMNO") & "'"
        '            ds_docdetail_updateQty = executeQuery(cmd_docdetail_updateQty)
        '        Next
        '    End If


        '    Dim cmddocdetail_cancel As String
        '    Dim cmddocheader_cancel As String
        '    Dim ds As New DataSet

        '    cmddocdetail_cancel = <code>
        '                              <![CDATA[
        '               update DocDetail set [STATUS] = 'C' 
        '               where [docno] in (select [Doc No] from [dbo].[TO_MR_Collated_Header_2] where [Document No] = ']]><%= Session("tmpdoc") %><![CDATA[')
        '               and [ITEMNO] in (select [ITEM NO] from [dbo].[]]><%= Replace(Session("tmpdoc"), "-", "") %><![CDATA[] where [TOTAL REQ] = 0)
        '               ]]>
        '                          </code>.Value
        '    executeQuery(cmddocdetail_cancel)


        '    cmddocheader_cancel = "SELECT sum([TOTAL REQ]) as [totalsum] FROM [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "]"
        '    ds = executeQuery(cmddocheader_cancel)
        '    If ds.Tables(0).Rows(0)("totalsum").ToString = 0 Then
        '        cmddocheader_cancel = "UPDATE DocHeader SET STATUS = 'C' WHERE [DOCNO] in (select [Doc NO] from [TO_MR_Collated_Header_2] where [Document No] = '" & Session("tmpdoc") & "')"
        '        executeQuery(cmddocheader_cancel)

        '        cmddocheader_cancel = "UPDATE TO_MR_Collated_Header_2 SET [Status] = 'CANCELLED' where [Document No] ='" & Session("tmpdoc") & "'"
        '        executeQuery(cmddocheader_cancel)
        '    End If

        '    CreateLog(Session("tmpdoc"))

        '    BindGrid2(dropdownMainCat.Text)
        'Catch ex As Exception
        '    ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
        '    Exit Try
        'End Try
    End Sub

    Protected Sub onEnterSave()
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim tmpid As Integer = 0
        Dim gvcolcnt As Integer = GridView1.Columns.Count - 1
        Dim gvrowcnt As Integer = GridView1.Rows.Count - 1
        Dim row As Integer = 0
        Dim col As Integer
        Dim reqt As Integer = 0
        Dim availt As Integer = 0


        Try
            row = -1
            Do While gvrowcnt > row
                Dim colvalue As Integer = 0
                tmpid = 0
                row = row + 1
                col = 13
                Do While gvcolcnt + 1 > col
                    col = col + 1
                    tmpid = tmpid + 1

                    If Trim(CType(GridView1.Rows(row).Cells(col).FindControl("txt" + tmpid.ToString), TextBox).Text) = Nothing Then
                        'CType(GridView1.Rows(row).Cells(col).FindControl("txt" + tmpid.ToString), TextBox).Text = 0
                        colvalue = 0
                    Else
                        colvalue = Convert.ToInt64(CType(GridView1.Rows(row).Cells(col).FindControl("txt" + tmpid.ToString), TextBox).Text)
                        'Convert.ToInt64(CType(GridView1.Rows(row).Cells(col).FindControl("txt" + tmpid.ToString), TextBox).Text)
                    End If

                    Dim a As Integer = 13
                    reqt += colvalue 'Convert.ToInt64(CType(GridView1.Rows(row).Cells(a).FindControl("txt" + tmpid.ToString), TextBox).Text)

                    a = a + 1

                    availt = Convert.ToInt64(Trim(GridView1.Rows(row).Cells(5).Text)) - reqt

                    Dim cmd As String =
                                <code>
                                    <![CDATA[
                            BEGIN TRANSACTION;
                                UPDATE [dbo].]]><%= Strings.Replace(Session("tmpdoc"), "-", "") %><![CDATA[
                                SET
                                    [TOTAL REQ] = ']]><%= reqt %><![CDATA['
                                    , [QTY AVAIL] =  ']]><%= availt %><![CDATA['
                                    , []]><%= Trim(GridView1.Columns(col - 1).HeaderText) %><![CDATA[] = ']]><%= colvalue %><![CDATA['
                                WHERE
                                    [Item No] = ']]><%= Trim(GridView1.Rows(row).Cells(3).Text) %><![CDATA['
                            COMMIT TRANSACTION;
                            ]]>
                                </code>.Value

                    sqlcmd = New SqlCommand(cmd, objConn)
                    sqlcmd.ExecuteNonQuery()

                Loop
                reqt = 0
                availt = 0
            Loop

            objConn.Close()
            objConn.Dispose()


            Dim cmd_docdetail_updateQty_select As String
            Dim ds_docdetail_updateQty_select As New DataSet

            cmd_docdetail_updateQty_select = "EXEC sp_selectDocDetail 0,'" & Session("tmpdoc") & "',''"
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


            Dim cmddocdetail_cancel As String
            Dim cmddocheader_cancel As String
            Dim ds As New DataSet

            cmddocdetail_cancel = <code>
                                      <![CDATA[
                       update DocDetail set [STATUS] = 'C',[ALLOCATEDQTY] = '0'
                       where [docno] in (select [Doc No] from [dbo].[TO_MR_Collated_Header_2] where [Document No] = ']]><%= Session("tmpdoc") %><![CDATA[')
                       and [ITEMNO] in (select [ITEM NO] from [dbo].[]]><%= Replace(Session("tmpdoc"), "-", "") %><![CDATA[] where [TOTAL REQ] = 0) and [Brand] = ']]><%= Session("tmpbrand") %><![CDATA['
                       ]]>
                                  </code>.Value
            executeQuery(cmddocdetail_cancel)


            cmddocheader_cancel = "SELECT sum([TOTAL REQ]) as [totalsum] FROM [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "]"
            ds = executeQuery(cmddocheader_cancel)
            If ds.Tables(0).Rows(0)("totalsum").ToString = 0 Then
                cmddocheader_cancel = "UPDATE DocHeader SET STATUS = 'C' WHERE [DOCNO] in (select [Doc NO] from [TO_MR_Collated_Header_2] where [Document No] = '" & Session("tmpdoc") & "')"
                executeQuery(cmddocheader_cancel)

                cmddocheader_cancel = "UPDATE TO_MR_Collated_Header_2 SET [Status] = 'CANCELLED' where [Document No] ='" & Session("tmpdoc") & "'"
                executeQuery(cmddocheader_cancel)
            End If

            CreateLog(Session("tmpdoc"))

            BindGrid2(dropdownMainCat.Text)
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('Changes successfully saved.');", True)
        Catch ex As Exception
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
            Exit Try
        End Try
    End Sub

    Private Sub txt_itemcode_TextChanged(sender As Object, e As EventArgs) Handles txt_itemcode.TextChanged
        LoadItems()
    End Sub

    Private Sub GridView1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridView1.RowDataBound

    End Sub

    Private Sub buttondelete_Click(sender As Object, e As EventArgs) Handles buttondelete.Click
        'buttondelete.OnClientClick = "if ( ! CreateFile()) return false;"
        Session("isEnter") = ""
        Session("deletionType") = ""
        Session("negativeType") = ""
        Session("deleteType") = "document"
        msgbox2_label.Text = "Are you sure you want to cancel this document?"
        div_message_box.Visible = True

    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        Dim cmd As String
        Dim ds As New DataSet

        If Session("deletionType") = "negativeQuantity" Then
            DeleteAllNegative()
        End If

        If Session("isEnter") = "yes" Then
            onEnterSave()
        End If

        If Session("deleteType") = "item" Then
            Try
                cmd = "select * from [dbo].[" & Replace(Session("selected_doc"), "-", "") & "]"
                ds = executeQuery(cmd)
                If ds.Tables(0).Rows.Count = 1 Then
                    ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('Document cant be empty. You can delete the document instead.');", True)
                Else
                    executeQuery("DELETE FROM [dbo].[" & Replace(Session("selected_doc"), "-", "") & "] where [ITEM NO] = '" & Session("selected_item") & "'")
                    executeQuery("UPDATE [dbo].[DocDetail] SET [STATUS] = 'C', [ALLOCATEDQTY] = '0' where [DOCNO] in (select [Doc No] from [dbo].[TO_MR_Collated_Header_2] where [Document No] = '" & Session("selected_doc") & "') and [ITEMNO] = '" & Session("selected_item") & "' and [BRAND] = '" & Session("selected_brand") & "' ")
                    executeQuery("INSERT INTO TO_MR_Log2 ([user name],[date],[doc no],[action]) values ('" & Session("tmpuser") & "','" & DateAndTime.Now & "','" & Session("selected_doc") & " / " & Session("selected_item") & "','delete item')")
                    Session("deletesuccess") = "deleted"
                    Response.Redirect("MRDetails.aspx")
                End If

            Catch ex As Exception
                ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
                Exit Try
            End Try
        ElseIf Session("deleteType") = "document" Then
            Try

                cmd = "UPDATE DocDetail SET [STATUS] = 'C', [ALLOCATEDQTY] = '0' WHERE [DOCNO] IN (SELECT [Doc No] FROM [dbo].[TO_MR_Collated_Header_2] WHERE [Document No] = '" & Session("tmpdoc") & "')
                AND [ITEMNO] IN (SELECT [ITEM NO] FROM [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "]) and [BRAND] = '" & Session("tmpbrand") & "'"
                executeQuery(cmd)

                cmd = "UPDATE DocHeader SET STATUS = 'C' WHERE [DOCNO] IN (SELECT [Doc NO] FROM [TO_MR_Collated_Header_2] WHERE [Document No] = '" & Session("tmpdoc") & "')"
                executeQuery(cmd)

                cmd = "UPDATE TO_MR_Collated_Header_2 set [STATUS] = 'CANCELLED' where [Document No] = '" & Session("tmpdoc") & "'"
                executeQuery(cmd)

                cmd = <code>
                          <![CDATA[
                            declare @docnumber as nvarchar(MAX) = replace(']]><%= Session("tmpdoc") %><![CDATA[','-','')

                            IF OBJECT_ID('tempdb..#TEMP_GETDATA') IS NOT NULL
                            DROP TABLE #TEMP_GETDATA

                            CREATE TABLE #TEMP_GETDATA (
                             ID INT NOT NULL IDENTITY(1,1),
                             [LOCATION NAME] VARCHAR(MAX)
                            )

							INSERT INTO #TEMP_GETDATA
							select COLUMN_NAME
							from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = @docnumber and ORDINAL_POSITION >= 12

							declare @count as int
							declare @location as nvarchar(MAX)
							declare @str as varchar(MAX)

							set @count = (select count([LOCATION NAME]) from #TEMP_GETDATA)

							while @count <> 0
							begin
							set @location = (select [LOCATION NAME] from #TEMP_GETDATA where [ID] = @count)
							set @str = 'UPDATE ' + @docnumber + ' set [' + @location + '] = 0, [TOTAL REQ] = 0, [QTY AVAIL] = [WHSE QTY]'
							exec(@str)
							set @count = @count -1
							end
                            ]]>
                      </code>.Value
                executeQuery(cmd)


                cmd = "INSERT INTO TO_MR_Log2 ([user name],[date],[doc no],[action]) values ('" & Session("tmpuser") & "','" & Date.Now & "','" & Session("tmpdoc") & "','delete')"
                executeQuery(cmd)

                Session("deletesuccess") = "deleted"
                Response.Redirect("MR.aspx")
            Catch ex As Exception
                ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
                Exit Try
            End Try
        End If

        Session("isEnter") = ""
        Session("selected_doc") = ""
        Session("selected_item") = ""
        'Session("tmpbrand") = ""
        Session("deleteType") = ""
        div_message_box.Visible = False
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        button1.Text = "Delete"
    End Sub

    Private Sub btnlogout_Click(sender As Object, e As EventArgs) Handles btnlogout.Click



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

    Private Sub BtnHome_Click(sender As Object, e As EventArgs) Handles BtnHome.Click
        Response.Redirect("MR.aspx")
    End Sub

    Private Sub addItems_Click(sender As Object, e As ImageClickEventArgs) Handles addItems.Click
        If dropdownItem.Text <> "0" Then
            AddItem(dropdownItem.Text)
        Else
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('Please select Item to add first.');", True)
        End If
    End Sub

    Private Sub export_btn_Click(sender As Object, e As ImageClickEventArgs) Handles export_btn.Click
        Session("selectedDocwithoutdash") = Session("tmpdoc")
        Session("selectedDoc") = Replace(Session("tmpdoc"), "-", "")
        Response.Redirect("ExportToExcel.aspx")
    End Sub

    Private Sub upload_btn_Click(sender As Object, e As ImageClickEventArgs) Handles upload_btn.Click
        Session("selectedDoc") = Session("tmpdoc")
        Response.Redirect("Upload.aspx")
    End Sub

    Private Sub btn_delete_whs_negative_Click(sender As Object, e As EventArgs) Handles btn_delete_whs_negative.Click

        Session("isEnter") = ""
        Session("deleteType") = ""
        Session("deletionType") = "negativeQuantity"
        Session("negativeType") = "WHSE"
        button1.Text = "Delete"
        msgbox2_label.Text = "Delete all items with negative WHSE QTY?"
        div_message_box.Visible = True

    End Sub

    Private Sub btn_delete_qtyavail_negative_Click(sender As Object, e As EventArgs) Handles btn_delete_qtyavail_negative.Click

        Session("isEnter") = ""
        Session("deleteType") = ""
        Session("deletionType") = "negativeQuantity"
        Session("negativeType") = "QTYAVAIL"
        button1.Text = "Delete"
        msgbox2_label.Text = "Delete all items with negative QTY AVAIL?"
        div_message_box.Visible = True

    End Sub

    Protected Sub DeleteAllNegative()
        Dim cmd As String
        Dim ds As New DataSet

        If Session("negativeType") = "WHSE" Then

            Try
                cmd = "select * from [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "] where [WHSE QTY] <= 0"
                ds = executeQuery(cmd)

                If ds.Tables(0).Rows.Count > 0 Then
                    cmd = "UPDATE DocDetail SET [STATUS] = 'C', [ALLOCATEDQTY] = '0' WHERE [DOCNO] IN (SELECT [Doc No] FROM [dbo].[TO_MR_Collated_Header_2] WHERE [Document No] = '" & Session("tmpdoc") & "')
                AND [ITEMNO] IN (SELECT [ITEM NO] FROM [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "] where [WHSE QTY] <= 0) and [BRAND] = '" & Session("tmpbrand") & "'"
                    executeQuery(cmd)

                    cmd = "delete from [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "] where [WHSE QTY] <= 0"
                    executeQuery(cmd)

                    ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('All items with negative and zero WHSE QTY successfully deleted.');location.href = 'MRDetails.aspx'", True)
                Else
                    ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('There is no negative or zero warehouse quantity in this document.');", True)
                End If
            Catch ex As Exception
                ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
            End Try

        ElseIf Session("negativeType") = "QTYAVAIL" Then

            Try
                cmd = "select * from [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "] where [QTY AVAIL] < 0"
                ds = executeQuery(cmd)

                If ds.Tables(0).Rows.Count > 0 Then
                    cmd = "UPDATE DocDetail SET [STATUS] = 'C', [ALLOCATEDQTY] = '0' WHERE [DOCNO] IN (SELECT [Doc No] FROM [dbo].[TO_MR_Collated_Header_2] WHERE [Document No] = '" & Session("tmpdoc") & "')
                AND [ITEMNO] IN (SELECT [ITEM NO] FROM [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "] where [QTY AVAIL] < 0) and [BRAND] = '" & Session("tmpbrand") & "'"
                    executeQuery(cmd)

                    cmd = "delete from [dbo].[" & Replace(Session("tmpdoc"), "-", "") & "] where [QTY AVAIL] < 0"
                    executeQuery(cmd)

                    ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('All items with negative QTY AVAIL successfully deleted.');location.href = 'MRDetails.aspx'", True)
                Else
                    ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('There is no negative QTY AVAIL in this document.');", True)
                End If
            Catch ex As Exception
                ScriptManager.RegisterStartupScript(Me, Page.GetType(), " Thenalert;", "alert('" & ex.Message & "');", True)
            End Try

        End If

        Session("deletionType") = ""
        Session("negativeType") = ""

    End Sub

    Private Sub dropdownItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dropdownItem.SelectedIndexChanged

    End Sub
End Class