Imports System.Data.Odbc

Module db_Firebird
    Public con As OdbcConnection

    Friend dbName As String = "Database.FDB"
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

    Public Function getTables(ByVal tblName As String) As List(Of String)
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
            MsgBox("Error in retrieving tables", MsgBoxStyle.Critical)
            Return Nothing
        End Try
    End Function

    Public Function getSQL(ByVal mySql As String) As DataSet
        dbOpen()



        dbClose()
    End Function

#Region "SPECIALIZED FUNCTIONS"

#End Region
End Module
