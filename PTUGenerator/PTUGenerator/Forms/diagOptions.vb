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
            mod_extract.SaveUrl = .GetSection("Extractor").GetKey("Path").Value
            db_Firebird.dbName = .GetSection("Extractor").GetKey("DB").Value

            txtArea.Text = mod_extract.AreaCode
            txtBranch.Text = mod_extract.BranchCode
            cboCompany.Text = mod_extract.Company
            txtSave.Text = mod_extract.SaveUrl
            txtDatabase.Text = db_Firebird.dbName
        End With
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Not System.IO.File.Exists(frmMain2.configFile) Then
            System.IO.File.Create(frmMain2.configFile).Dispose()
            mod_extract.CodesInit()
        End If

        With iniFile
            .Load(frmMain2.configFile)
            .AddSection("Extractor").AddKey("Branch").Value = txtBranch.Text
            .AddSection("Extractor").AddKey("Area").Value = txtArea.Text
            .AddSection("Extractor").AddKey("Company").Value = cboCompany.Text
            .AddSection("Extractor").AddKey("Path").Value = txtSave.Text
            .AddSection("Extractor").AddKey("DB").Value = txtDatabase.Text

            .Save(frmMain2.configFile)
        End With

        LoadConfig()
        MsgBox("Configuration Saved", MsgBoxStyle.Information)
        Me.Close()
    End Sub

    Private Sub txtSave_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSave.DoubleClick
        txtSave.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    End Sub
End Class