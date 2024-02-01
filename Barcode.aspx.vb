Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel

Imports System.Xml.Linq

Public Class Barcode
    Inherits System.Web.UI.Page

    Dim connstr As String = [TO].My.Settings.SQLConnection.ToString '"Server='VM_RS';Initial Catalog='TOMR';user id='sa';password='Pa$$w0rd';"

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
                from [TOMR].[dbo].[TO_MR_Barcode_Types] as a with (nolock)
                
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
        nameColumn.DataField = "Ctrl No"
        nameColumn.HeaderText = "NO."
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvBarcode.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Item Owner"
        nameColumn.HeaderText = "ITEM OWNER"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvBarcode.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Brand"
        nameColumn.HeaderText = "BRAND"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvBarcode.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Main Category"
        nameColumn.HeaderText = "MAIN CATEGORY"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvBarcode.Columns.Add(nameColumn)

        nameColumn = New BoundField()
        nameColumn.DataField = "Barcode Size"
        nameColumn.HeaderText = "BARCODE SIZE"
        nameColumn.InsertVisible = False
        nameColumn.ReadOnly = True
        gvBarcode.Columns.Add(nameColumn)



    End Sub

    Private Sub BindGrid()

        dsmr = Me.GetBarcode()
        If dsmr.Tables.Count <> 0 Then
            gvBarcode.DataSource = dsmr
            gvBarcode.DataBind()
        End If

    End Sub




End Class