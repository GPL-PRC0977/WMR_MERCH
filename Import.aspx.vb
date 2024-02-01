Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Imports Excel = Microsoft.Office.Interop.Excel

Public Class Import
    Inherits Page

    Sub LoadCompanies()
        Dim sConnectionString As String
        Dim objConn As SqlConnection
        sConnectionString = "Server='VM_RS';Initial Catalog='Reports';user id='sa';password='Pa$$w0rd';"
        objConn = New SqlConnection(sConnectionString)
        objConn.Open()
        Dim sqlcmd As New SqlCommand("SELECT DISTINCT " & _
            "[Company Code] " & _
            "FROM [t_company] " & _
            "ORDER BY [Company Code]", objConn)
        Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
        ddlCompany.Items.Clear()
        ddlCompany.Items.Add("")
        Do While sqlreader.Read
            ddlCompany.Items.Add(sqlreader("Company Code"))
        Loop
        sqlreader.Close()
        objConn.Close()
    End Sub

    Sub LoadStores()
        Dim sConnectionString As String
        Dim objConn As SqlConnection
        sConnectionString = "Server='VM_RS';Initial Catalog='Reports';user id='sa';password='Pa$$w0rd';"
        objConn = New SqlConnection(sConnectionString)
        objConn.Open()
        Dim sqlcmd As New SqlCommand("SELECT DISTINCT " &
            "([Location Name] + ' - ' + [Location Code]) AS [Store] " &
            "FROM [t_locations4] " &
            "WHERE [Company Ownership Code] = '' or [Company Ownership Code] = '" & ddlCompany.Text & "'" &
            "ORDER by [Store] ", objConn)
        Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
        ddlStore.Items.Clear()
        ddlStore.Items.Add("")
        Do While sqlreader.Read
            ddlStore.Items.Add(sqlreader("Store"))
        Loop
        sqlreader.Close()
        objConn.Close()
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If _Default.statusLog <> "IN" Then
            Page.Response.Redirect("Default.aspx")
        End If
        If Not Page.IsPostBack Then
            LoadCompanies()
        End If
    End Sub

    Protected Sub buttonImport_Click(sender As Object, e As EventArgs) Handles buttonImport.Click
        Dim sConnectionString As String
        Dim objConn As SqlConnection

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim a As Integer = 0
        Dim tmpcount As Integer = 0
        Dim tmpStoreCode As String = ""
        Dim tmpStoreName As String = ""
        Dim tmpFileName As String = ""
        Dim tmpsql As String = ""
        Dim sqlcmd As SqlCommand

        Dim tmpItemNo As String = ""
        Dim tmpQty As Integer = 0

        If fileuploadExcel.HasFile And Trim(ddlStore.Text) <> "" And Trim(textboxDate.Text <> "") Then

            tmpFileName = fileuploadExcel.FileName '"Up" + Replace(Replace(Replace(Now(), " ", ""), ":", ""), "/", "")

            Try
                fileuploadExcel.SaveAs("C:\Uploads\" &
                    fileuploadExcel.FileName)
                labelMessage.Text = "File Name: " & _
                    fileuploadExcel.PostedFile.FileName & "<br>" & _
                    "File Size: " & _
                    fileuploadExcel.PostedFile.ContentLength & " kb<br>" & _
                    "Content Type: " & _
                    fileuploadExcel.PostedFile.ContentType

                sConnectionString = "Server='VM_RS';Initial Catalog='Reports';user id='sa';password='Pa$$w0rd';"
                objConn = New SqlConnection(sConnectionString)
                objConn.Open()

                xlApp = New Excel.Application
                xlWorkBook = xlApp.Workbooks.Open("C:\Uploads\" & tmpFileName)
                xlWorkSheet = xlWorkBook.Worksheets(1)

                Dim rowno As Integer = 9

                If InStr(1, ddlStore.Text, "-") <> 0 Then
                    tmpStoreCode = Strings.Right(ddlStore.Text, Len(ddlStore.Text) - (InStr(1, ddlStore.Text, "-", CompareMethod.Binary) + 1))
                End If

                If InStr(1, ddlStore.Text, "-") <> 0 Then
                    tmpStoreName = Strings.Left(ddlStore.Text, (InStr(1, ddlStore.Text, "-", CompareMethod.Binary) - 1))
                End If

                'sqlcmd = New SqlCommand("SELECT COUNT(*) AS [Count] FROM [t_invt_actual] " & _
                '    "WHERE [Location Code] = '" & tmpStoreCode & "' " & _
                '    "AND [Count Date] = '" & textboxDate.Text & "'", objConn)
                'Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
                'Do While sqlreader.Read
                '    tmpcount = sqlreader("Count")
                'Loop
                'sqlreader.Close()

                sqlcmd = New SqlCommand("DELETE FROM [t_invt_actual] " & _
                    "WHERE [Location Code] = '" & tmpStoreCode & "' " & _
                    "AND [Count Date] = '" & textboxDate.Text & "'", objConn)
                sqlcmd.ExecuteNonQuery()

                'If tmpcount = 0 Then
                Do
                    a = a + 1

                    tmpItemNo = xlWorkSheet.Cells(rowno, 1).Value()
                    tmpQty = CInt(xlWorkSheet.Cells(rowno, 6).Value())

                    tmpsql = "INSERT INTO [t_invt_actual] " & _
                        "([Location Code]" & _
                        ",[Location Name]" & _
                        ",[Count Date]" & _
                        ",[Ctrl No]" & _
                        ",[Item No]" & _
                        ",[Quantity])" & _
                    "VALUES " & _
                        "('" & tmpStoreCode & "'" & _
                        ",'" & tmpStoreName & "'" & _
                        ",'" & textboxDate.Text & "'" & _
                        ",'" & a & "'" & _
                        ",'" & tmpItemNo & "'" & _
                        ",'" & tmpQty & "')"
                    sqlcmd = New SqlCommand(tmpsql, objConn)
                    sqlcmd.ExecuteNonQuery()

                    tmpItemNo = ""
                    tmpQty = 0
                    rowno = rowno + 1

                Loop Until Trim(xlWorkSheet.Cells(rowno, 1).Value) = ""
                labelMessage.Text = "File import done."
                'Else
                'labelMessage.Text = "Location and date already uploaded. Overwrite?"
                'If msgbox = DialogResult.Yes Then
                '    sqlcmd = New SqlCommand("DELETE FROM [t_invt_actual] " & _
                '        "WHERE [Location Code] = '" & tmpStoreCode & "' " & _
                '        "AND [Count Date] = '" & dtpCountDate.Text & "'", objConn)
                '    sqlcmd.ExecuteNonQuery()
                '    Do
                '        a = a + 1
                '        tmpItemNo = xlWorkSheet.Cells(rowno, 1).Value()
                '        tmpQty = CInt(xlWorkSheet.Cells(rowno, 6).Value())
                '        tmpsql = "INSERT INTO [t_invt_actual] " & _
                '           "([Location Code]" & _
                '           ",[Location Name]" & _
                '           ",[Count Date]" & _
                '           ",[Ctrl No]" & _
                '           ",[Item No]" & _
                '           ",[Quantity])" & _
                '        "VALUES " & _
                '           "('" & tmpStoreCode & "'" & _
                '           ",'" & tmpStoreName & "'" & _
                '           ",'" & dtpCountDate.Text & "'" & _
                '           ",'" & a & "'" & _
                '           ",'" & tmpItemNo & "'" & _
                '           ",'" & tmpQty & "')"
                '        sqlcmd = New SqlCommand(tmpsql, objConn)
                '        sqlcmd.ExecuteNonQuery()
                '        tmpItemNo = ""
                '        tmpQty = 0
                '        rowno = rowno + 1
                '    Loop Until Trim(xlWorkSheet.Cells(rowno, 1).Value) = ""
                '    labelMessage.Text = "File import done."
                'Else
                '    labelMessage.Text = "File import cancelled."
                'End If
                'End If

                xlWorkBook.Close()
                xlApp.Quit()
                objConn.Close()

                My.Computer.FileSystem.DeleteFile("C:\Uploads\" & tmpFileName)
                ddlCompany.Text = ""
                ddlStore.Text = ""
                textboxDate.Text = ""
                'labelMessage.Text = ""
                tmpFileName = ""


            Catch ex As Exception
                labelMessage.Text = "ERROR: " & ex.Message.ToString()
            End Try

        Else
            labelMessage.Text = "Please fill-up all fields."
        End If

    End Sub

    Protected Sub ddlCompany_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCompany.SelectedIndexChanged
        LoadStores()
    End Sub

End Class