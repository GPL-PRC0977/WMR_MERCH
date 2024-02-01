Imports System.Data.SqlClient
Public Class test
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim cmd As String
        Dim ds As New DataSet

        Try
            cmd = "IF OBJECT_ID('tempdb..#tmp1') IS NOT NULL
DROP TABLE #tmp1

IF OBJECT_ID('tempdb..#tmp2') IS NOT NULL
DROP TABLE #tmp2

SELECT COLUMN_NAME 
into #tmp1
FROM INFORMATION_SCHEMA.COLUMNS as colname 
WHERE TABLE_NAME = 'WMRLTS1904260' and [ORDINAL_POSITION] > 11
select row_number() over(order by [COLUMN_NAME] asc) as [ID],[COLUMN_NAME] into #tmp2 from #tmp1 as a

declare @ID as int
declare @strsql as nvarchar(max)
set @ID = (select count([COLUMN_NAME]) from #tmp1)
	while @ID <> 0
		begin
	
		declare @location as nvarchar(max) = (select [COLUMN_NAME] from #tmp2 where [ID]=@ID)
		set @strsql = 'select ''' + @location + ''' as [LOCATION],(select sum(['+ @location +']) from WMRLTS1904260) as [TOTAL QTY SERVED]'
		exec(@strsql)

		set @ID = @ID - 1
		end
"
            ds = executeQuery(cmd)
            dg.DataSource = ds.Tables(0)
            dg.DataBind()

        Catch ex As Exception

        End Try
    End Sub

End Class