Public Class diagOptions
    Dim iniFile As New IniFile

    Private Sub diagOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboCompany.SelectedIndex = 0
        LoadConfig()
    End Sub

    Private Sub LoadConfig()
        On Error Resume Next

        With iniFile
            .Load(frmMain2.configFile)
            mod_extract.BranchCode = .GetSection("Extractor").GetKey("Branch").Value
            mod_extract.AreaCode = .GetSection("Extractor").GetKey("Area").Value
            mod_extract.Company = .GetSection("Extractor").GetKey("Company").Value
            db_Firebird.dbName = IsError(.GetSection("Extractor").GetKey("DB").Value)

            txtArea.Text = mod_extract.AreaCode
            txtBranch.Text = mod_extract.BranchCode
            cboCompany.Text = mod_extract.Company
            txtDatabase.Text = db_Firebird.dbName
        End With
    End Sub
End Class