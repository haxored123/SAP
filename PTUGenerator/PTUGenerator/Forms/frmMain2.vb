Public Class frmMain2
    Friend configFile As String = "Aerauxel.ini"
    Friend iniFile As New IniFile
    Dim adShow As Boolean = False

    Private Sub lblSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSite.Click
        Process.Start("http://adf.ly/1OHMIb")
    End Sub

    Private Sub btnGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGen.Click
        tmrAds.Enabled = True
        mod_extract.TransDate = mcSales.SelectionRange.Start
        If devMode Then mod_extract.TransDate = "10/01/2015"
        Frozen(1)

        fetchingSales(mod_extract.TransDate)

        MsgBox(mod_extract.TransDate & " extracted", MsgBoxStyle.Information)
        Frozen(0)
    End Sub

    Private Sub Frozen(ByVal st As Boolean)
        btnGen.Enabled = Not st
        btnExit.Enabled = Not st
        pbLoad.Visible = st
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

    Private Sub frmMain2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DoubleClick
        diagOptions.Show()
    End Sub

    Private Sub frmMain2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = My.Application.Info.Title & "|" & mod_extract.Company & " by IT Department 2015 | Version " & Me.GetType.Assembly.GetName.Version.ToString
        wbAds.Navigate("http://adf.ly/7104086/banner/pgc-itdept.org/software/ptu-generator/")

        LoadConfig()

        If devMode Then MsgBox("DEVELOPER MODE", MsgBoxStyle.Information)
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub fetchingSales(ByVal st As Date)
        Try
            Dim dailySalesOR As New DataSet
            Dim mySql As String
            mySql = "SELECT ID, TRANSDATE "
            mySql &= vbCr & "FROM POSENTRY  "
            mySql &= vbCr & String.Format("WHERE POSTYPE = 'Sales' AND TRANSDATE = '{0}'", st.ToString("M/d/yyyy"))

            dailySalesOR = LoadSQL(mySql, "GetAll")
            Dim MaxRow As Integer = dailySalesOR.Tables("GetAll").Rows.Count

            ProcessBarInit(0, MaxRow)
            For idx As Integer = 0 To MaxRow - 1
                With dailySalesOR
                    Dim EntryID As String = .Tables("GetAll").Rows(idx).Item("ID")
                    mySql = "SELECT ITM.POSENTRYID, ITM.LINENO, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"" "
                    mySql &= vbCrLf & "FROM POSITEM ITM INNER JOIN ITEMMASTER ITMM ON ITMM.ITEMNO = ITM.ITEMNO "
                    mySql &= vbCrLf & String.Format("WHERE ITM.POSENTRYID = '{0}'", EntryID)

                    Dim tmpSales As DataSet = LoadSQL(mySql)
                    Dim tmpMaxCount As Integer = tmpSales.Tables(0).Rows.Count
                    Dim tblName As String = tmpSales.Tables(0).Rows(tmpMaxCount - 1).Item("ItemCode")

                    mySql &= vbCrLf & " AND ITM.QTY > 0"
                    tmpSales = LoadSQL(mySql, tblName)

                    If Not dailySalesOR.Tables.Contains(tblName) Then
                        dailySalesOR.Tables.Add(tmpSales.Tables(tblName).Copy)
                    Else
                        dailySalesOR.Merge(tmpSales.Tables(tblName))
                    End If
                End With
                Console.WriteLine("-------------------======================================TransNum: " & idx)
                AddProcessBar()
                Application.DoEvents()
            Next

            Console.WriteLine("Tables: " & dailySalesOR.Tables.Count)
            Dim str As String = ""
            For cnt As Integer = 0 To dailySalesOR.Tables.Count - 1
                str &= dailySalesOR.Tables(cnt).TableName & vbTab
            Next
            Console.WriteLine(str)
            mod_extract.GeneratePTUFileV2(dailySalesOR)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Fetching")
        End Try
    End Sub

    Friend Sub AddProcessBar()
        On Error Resume Next
        pbLoad.Value += 1
    End Sub

    Friend Sub ProcessBarInit(ByVal st As Integer, ByVal en As Integer)
        pbLoad.Minimum = st
        pbLoad.Maximum = en
        Application.DoEvents()
    End Sub

    Private Sub wbAds_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wbAds.DocumentCompleted
        adShow = True
    End Sub

    Private Sub tmrAds_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrAds.Tick
        Static cnt As Integer
        cnt += 1
        If cnt >= 5 And adShow Then
            tmrAds.Enabled = False
            pbIT.Visible = False
        End If
    End Sub
End Class