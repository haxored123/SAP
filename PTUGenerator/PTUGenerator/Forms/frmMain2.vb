Public Class frmMain2
    Friend configFile As String = "Aerauxel.ini"
    Friend iniFile As New IniFile
    Dim adShow As Boolean = False

    Private Sub lblSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSite.Click
        Process.Start("http://adf.ly/1OHMIb")
    End Sub

    Private Sub btnGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGen.Click
        Frozen(1)


    End Sub

    Private Sub Frozen(ByVal st As Boolean)
        btnGen.Enabled = Not st
        btnExit.Enabled = Not st
    End Sub

    Private Sub LoadConfig()
        With iniFile
            .Load(configFile)
            mod_extract.BranchCode = IIf(IsError(.GetSection("Extractor").GetKey("Branch").Value), "", .GetSection("Extractor").GetKey("Branch").Value)
            mod_extract.AreaCode = IIf(IsError(.GetSection("Extractor").GetKey("Area").Value), "", .GetSection("Extractor").GetKey("Area").Value)
            mod_extract.Company = IIf(IsError(.GetSection("Extractor").GetKey("Company").Value), "", .GetSection("Extractor").GetKey("Company").Value)
            db_Firebird.dbName = IIf(IsError(.GetSection("Extractor").GetKey("DB").Value), "", .GetSection("Extractor").GetKey("DB").Value)
            If mod_extract.BranchCode = "" Or _
                mod_extract.AreaCode = "" Or _
                mod_extract.Company = "" Or _
                db_Firebird.dbName = "" Then
                diagOptions.Show()
            End If

            txtArea.Text = mod_extract.AreaCode
            txtBranch.Text = mod_extract.BranchCode
            txtCompany.Text = mod_extract.Company
        End With
    End Sub

    Private Sub frmMain2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = My.Application.Info.Title & "|" & mod_extract.Company & " by IT Department 2015 | Version " & Me.GetType.Assembly.GetName.Version.ToString
        wbAds.Navigate("http://adf.ly/7104086/banner/pgc-itdept.org/software/ptu-generator/")
    End Sub
End Class