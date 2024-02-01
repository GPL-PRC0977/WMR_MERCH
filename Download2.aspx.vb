Public Class Download2
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim tmpfile As String = Request.QueryString("file")

        'If Session("xFile1") <> "" Then
        '    tmpfile = Session("xFile1")
        'End If


        If Not IsPostBack Then
            If tmpfile <> "" Then
                Dim strFile As String = Server.MapPath("\Uploads\" + tmpfile + ".csv")
                Dim Header As String = "Attachment; Filename=" + tmpfile + ".csv"
                Dim Dfile As New System.IO.FileInfo(strFile)

                Response.ContentType = "text/plain"
                Response.AddHeader("Content-Disposition", Header)
                Response.WriteFile(strFile)



                'Response.TransmitFile(Dfile.FullName)
                ''Don't forget to add the following line
                'Response.Flush()
                'Response.SuppressContent = True



                'If Session("xFile2").ToString <> "" Then
                '    Session("xFile1") = Session("xFile2")
                '    sMain = "window.open('Download.aspx'Download.aspx?id=" & Now.TimeOfDay.TotalMilliseconds & "', '_blank');"
                '    ClientScript.RegisterStartupScript(Me.GetType(), "script", sMain, True)
                'End If

                'My.Computer.FileSystem.DeleteFile(strFile)
                'Response.End()
                'ClientScript.RegisterStartupScript(Me.GetType(), "script", " alert('Successfully created text file(s)');", True)

            End If

        End If
    End Sub

End Class