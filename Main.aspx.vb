Public Class Main
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        If CType(Session.Item("tmpuser"), String) = "" Then
            Page.Response.Redirect("Default.aspx")
        End If

    End Sub

End Class