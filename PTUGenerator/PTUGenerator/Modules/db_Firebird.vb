Imports System.Data.Odbc

Module db_Firebird
    Public con As OdbcConnection

    Friend dbName As String = "D:\CaDeAtH\Documents\RAW\PTU Gen\RobGensan_PHOTO.FDB"
    Friend fbUser As String = "SYSDBA"
    Friend fbPass As String = "masterkey"
    Friend fbDataSet As New DataSet

    Private language() As String = _
        {"Connection error failed."}
    Private conStr As String = String.Empty

    Public Sub dbOpen()
        conStr = "DRIVER=Firebird/InterBase(r) driver;User=" & fbUser & ";Password=" & fbPass & ";Database=" & dbName & ";"

        con = New OdbcConnection(conStr)
        Try
            con.Open()
        Catch ex As Exception
            con.Dispose()
            MsgBox(language(0) + vbCrLf + ex.Message.ToString, vbCritical, "Error")
            Exit Sub
        End Try
    End Sub

    Public Sub dbClose()
        con.Close()
    End Sub

    Public Function getColumns(ByVal tblName As String) As List(Of String)
        Try
            Dim restrictions As String() = New String() {Nothing, Nothing, tblName, Nothing}

            dbOpen()
            Dim dtbl As DataTable = con.GetSchema("Columns", restrictions)
            dbClose()

            Dim tmpStr As New List(Of String)
            For Each dataRow As DataRow In dtbl.Rows
                tmpStr.Add(dataRow("Column_name"))
            Next

            Return tmpStr
        Catch ex As Exception
            dbClose()
            MsgBox("Error in retrieving table columns", MsgBoxStyle.Critical)
            Return Nothing
        End Try
    End Function

    Public Function LoadSQL(ByVal mySql As String) As DataSet
        Try
            dbOpen()
            Dim da As OdbcDataAdapter
            Dim fillData As String = "CustomSQL"
            Dim ds As New DataSet

            ds.Clear()
            da = New OdbcDataAdapter(mySql, con)
            da.Fill(ds, fillData)

            dbClose()

            Return ds
        Catch ex As Exception
            Console.WriteLine("SQL: " & mySql & "| ERR: " & ex.Message)
            MsgBox("Cannot do some queries", MsgBoxStyle.Critical, "Load SQL")
            dbClose()
            Return Nothing
        End Try
    End Function

#Region "SPECIALIZED FUNCTIONS"

#End Region
End Module
