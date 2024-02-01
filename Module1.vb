Imports System.Data.SqlClient
Module Module1
    Public con As New SqlConnection
    Public conn As New SqlConnection
    'Public conString As String = My.Settings.SQLConnection
    Public DA As New SqlDataAdapter
    Public cmd As New SqlCommand
    Public Function executeQuery(strSQL As String) As DataSet
        Dim ds As New DataSet

        Dim cn As New SqlConnection(My.Settings.SQLConnection.ToString)
        Dim cmd As New SqlCommand(strSQL)
        Dim da As New SqlDataAdapter(cmd)

        Try
            cmd.Connection = cn
            cmd.CommandTimeout = 0
            cn.Open()
            da.Fill(ds)

        Catch ex As Exception
            ds = Nothing
        End Try
        cn.Close()
        cn = Nothing
        Return ds
    End Function

    Public Function executeQueryRS(strSQL As String) As DataSet
        Dim ds As New DataSet

        Dim cn As New SqlConnection(My.Settings.RSconn.ToString)
        Dim cmd As New SqlCommand(strSQL)
        Dim da As New SqlDataAdapter(cmd)

        Try
            cmd.Connection = cn
            cmd.CommandTimeout = 0
            cn.Open()
            da.Fill(ds)

        Catch ex As Exception
            ds = Nothing
        End Try
        cn.Close()
        cn = Nothing
        Return ds
    End Function
    Function GetSelectedValues(ByVal lstBoxText As ListBox) As String
        Dim selectedValues As String = String.Empty

        For Each li As ListItem In lstBoxText.Items
            If li.Selected = True Then
                selectedValues += li.Text & ","
            End If
        Next

        selectedValues = selectedValues.ToString().Remove(selectedValues.ToString().LastIndexOf(","))

        Return selectedValues.ToString()
    End Function
End Module
