Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel

Imports System.Xml.Linq

Public Class Users
    Inherits System.Web.UI.Page

    Dim connstr As String = "Server='VM_RS';Initial Catalog='TOMR';user id='sa';password='Pa$$w0rd';"

    Dim dsmr As DataSet
    Dim dsloc As DataSet

    Dim tmptext As String
    Dim tmpint As Int16 = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If _Default.statusLog <> "TO" Then
            Page.Response.Redirect("Default.aspx")
        End If
        If Not Page.IsPostBack Then
            CreateGrid()
            BindGrid()
        End If
    End Sub

    Private Function GetBarcode() As DataSet
        Dim sqlscript As String =
        <code>
            <![CDATA[

                select
	                *
                from [Reports].[dbo].[t_user_security]
                where [filter] = 'BRAND'
                order by [user name], [owner], [value]
                
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
        End Using
    End Function

    Private Sub CreateGrid()

        Dim nameColumn As BoundField

        nameColumn = New BoundField()
        nameColumn.DataField = "ctrl no"
        nameColumn.HeaderText = "NO."
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvUser.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "user name"
        nameColumn.HeaderText = "USER NAME"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvUser.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "owner"
        nameColumn.HeaderText = "COMPANY"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvUser.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "value"
        nameColumn.HeaderText = "BRAND"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvUser.Columns.Add(nameColumn)


    End Sub

    Private Sub BindGrid()

        dsmr = Me.GetBarcode()
        If dsmr.Tables.Count <> 0 Then
            gvUser.DataSource = dsmr
            gvUser.DataBind()
        End If

    End Sub

    Protected Sub gvUser_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
        gvUser.PageIndex = e.NewPageIndex
        Me.BindGrid()
    End Sub

End Class