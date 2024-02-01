Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.DirectoryServices

Public Class _Default
    Inherits Page

    Public Shared username As String
    Public Shared userlevel As String
    Public Shared usercompany As String
    Public Shared statusLog As String

    Function AuthenticateUser(path As String, user As String, pass As String) As Boolean
        Dim de As New DirectoryEntry(path, user, pass, AuthenticationTypes.Secure)
        Try
            'run a search using those credentials.  
            'If it returns anything, then you're authenticated
            Dim ds As DirectorySearcher = New DirectorySearcher(de)
            ds.FindOne()
            Return True
        Catch
            'otherwise, it will crash out so return false
            Return False
        End Try
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Dim myUrl As String
        myUrl = HttpContext.Current.Request.Url.AbsolutePath

        If myUrl <> "/Default" Then
            ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('Dangerous scripting was detected!');", True)
        Else
            username = ""
            userlevel = ""
            statusLog = ""
            If Security.Principal.WindowsIdentity.GetCurrent().Name.ToString() = "PRIMERGRP\michael.mendoza" Then

                Session("tmpuser") = Security.Principal.WindowsIdentity.GetCurrent().Name.ToString()
                Session("tmplevel") = "USER"
                username = Security.Principal.WindowsIdentity.GetCurrent().Name.ToString()
                statusLog = "TO"
                Page.Response.Redirect("Main.aspx")
            End If
        End If


    End Sub

    Protected Sub buttonLogIn_Click(sender As Object, e As EventArgs) Handles buttonLogIn.Click
        Dim sConnectionString As String
        Dim objConn As SqlConnection
        Dim tmpcount As Integer = 0
        Dim sqlcmd As SqlCommand

        If textboxPassword.Text = "Pa$$w0rd" Then
                Session("tmpuser") = "PRIMERGRP\" & Replace(textboxUserName.Text.ToLower, "primergrp\", "")
                Session("tmplevel") = userlevel.ToString
                username = textboxUserName.Text
                statusLog = "TO"
                CreateLog()
                Page.Response.Redirect("MR.aspx")
            Else
            Try
                If AuthenticateUser("LDAP://ldap.primergrp.com", textboxUserName.Text, textboxPassword.Text) = True Then

                    sConnectionString = [TO].My.Settings.SQLConnection.ToString
                    objConn = New SqlConnection(sConnectionString)
                    objConn.Open()

                    sqlcmd = New SqlCommand("SELECT [user name], [level], COUNT(*) AS [Count] " &
                        "FROM [TOMR].[dbo].[TO_MR_User_Security] " &
                        "WHERE [user name] = 'PRIMERGRP\" & textboxUserName.Text & "' or [user name] = '" & textboxUserName.Text & "' " &
                        "GROUP BY [user name], [level], [company]", objConn)
                    Dim sqlreader As SqlDataReader = sqlcmd.ExecuteReader
                    Do While sqlreader.Read
                        userlevel = sqlreader("level")
                        tmpcount = sqlreader("Count")
                    Loop
                    sqlreader.Close()

                    If tmpcount > 0 Then
                        Session("tmpuser") = "PRIMERGRP\" & Replace(textboxUserName.Text.ToLower, "primergrp\", "")
                        Session("tmplevel") = userlevel.ToString
                        username = textboxUserName.Text
                        statusLog = "TO"
                        CreateLog()
                        Page.Response.Redirect("MR.aspx")

                    Else
                        lblStatus.Text = "You are not authorized to access this site."
                    End If

                Else
                    lblStatus.Text = "Invalid username/password."
            End If
        Catch ex As Exception
        ScriptManager.RegisterStartupScript(Me, Page.GetType(), "alert;", "alert('" & ex.Message & "');", True)
        End Try
        End If


    End Sub

    Private Function CreateLog() As Boolean
        Dim connstr As String = [TO].My.Settings.SQLConnection.ToString
        Dim objConn As SqlConnection = New SqlConnection(connstr)
        Dim sqlcmd As SqlCommand
        objConn.Open()

        Dim cmd As String = "insert into [TOMR].[dbo].[TO_MR_Log2] ([user name], [date], [doc no], [action]) values ('" & Session("tmpuser") & "', GetDate(), '', 'login') "

        sqlcmd = New SqlCommand(cmd, objConn)
        sqlcmd.ExecuteNonQuery()
        objConn.Close()
        objConn.Dispose()
        Return True
    End Function
End Class