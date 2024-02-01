Imports System.Data.SqlClient

Public Class SiteMaster
    Inherits MasterPage
    Private Const AntiXsrfTokenKey As String = "__AntiXsrfToken"
    Private Const AntiXsrfUserNameKey As String = "__AntiXsrfUserName"
    Private _antiXsrfTokenValue As String

    Protected Sub Page_Init(sender As Object, e As EventArgs)
        ' The code below helps to protect against XSRF attacks
        Dim requestCookie = Request.Cookies(AntiXsrfTokenKey)
        Dim requestCookieGuidValue As Guid
        If requestCookie IsNot Nothing AndAlso Guid.TryParse(requestCookie.Value, requestCookieGuidValue) Then
            ' Use the Anti-XSRF token from the cookie
            _antiXsrfTokenValue = requestCookie.Value
            Page.ViewStateUserKey = _antiXsrfTokenValue
        Else
            ' Generate a new Anti-XSRF token and save to the cookie
            _antiXsrfTokenValue = Guid.NewGuid().ToString("N")
            Page.ViewStateUserKey = _antiXsrfTokenValue

            Dim responseCookie = New HttpCookie(AntiXsrfTokenKey) With {
                 .HttpOnly = True,
                 .Value = _antiXsrfTokenValue
            }
            If FormsAuthentication.RequireSSL AndAlso Request.IsSecureConnection Then
                responseCookie.Secure = True
            End If
            Response.Cookies.[Set](responseCookie)
        End If

        AddHandler Page.PreLoad, AddressOf master_Page_PreLoad
    End Sub

    Protected Sub master_Page_PreLoad(sender As Object, e As EventArgs)
        If Not IsPostBack Then
            ' Set Anti-XSRF token
            ViewState(AntiXsrfTokenKey) = Page.ViewStateUserKey
            ViewState(AntiXsrfUserNameKey) = If(Context.User.Identity.Name, [String].Empty)
        Else
            ' Validate the Anti-XSRF token
            If DirectCast(ViewState(AntiXsrfTokenKey), String) <> _antiXsrfTokenValue OrElse DirectCast(ViewState(AntiXsrfUserNameKey), String) <> (If(Context.User.Identity.Name, [String].Empty)) Then
                Throw New InvalidOperationException("Validation of Anti-XSRF token failed.")
            End If
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Unnamed_LoggingOut(sender As Object, e As LoginCancelEventArgs)
        Context.GetOwinContext().Authentication.SignOut()
    End Sub

    Sub EndDocumentSession(sender As Object, e As EventArgs)
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