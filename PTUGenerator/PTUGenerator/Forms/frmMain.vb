Public Class frmMain

    Private configFile As String = "Aerauxel.ini"
    Private iniFile As New IniFile

    Private Sub btnGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGen.Click
        Try
            btnGen.Enabled = False

            Dim objSales As New DataSet
            Dim objSerial As New DataSet
            Dim tmpCC As New DataSet
            Dim objCC As New DataSet
            mod_extract.TransDate = dtpExtract.Value.ToString("M/d/yyyy")
            If devMode Then mod_extract.TransDate = "12/12/2014"

            'Sales
            Dim mySql As String
            mySql = "SELECT ENT.ID, ENT.TRANSDATE, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"""
            mySql &= vbCrLf & "FROM POSENTRY ENT "
            mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
            mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO"
            mySql &= vbCrLf & "WHERE "
            mySql &= vbCrLf & "POSTYPE = 'Sales' AND ITM.ITEMNO <> 'CASHD' AND ITM.ITEMNO <> 'CASHO' AND ITM.ITEMNO <> 'CARD3' AND "
            mySql &= vbCrLf & String.Format("ENT.TRANSDATE = '{0}'", mod_extract.TransDate)
            mySql &= vbCrLf & "ORDER BY ITM.ITEMNO ASC"
            objSales = LoadSQL(mySql)
            Console.WriteLine("Output: " & objSales.Tables(0).Rows.Count & vbCrLf & mySql)

            'For CC
            mySql = "SELECT ITM.POSENTRYID, ENT.NAME"
            mySql &= vbCrLf & "FROM POSENTRY ENT "
            mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
            mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO"
            mySql &= vbCrLf & "WHERE "
            mySql &= vbCrLf & String.Format("POSTYPE = 'Sales' AND ITM.ITEMNO = 'CARD3' AND ENT.TRANSDATE = '{0}'", mod_extract.TransDate)
            tmpCC = LoadSQL(mySql)


            Dim entry As String = "0"
            If tmpCC.Tables(0).Rows.Count = 1 Then entry = tmpCC.Tables(0).Rows(0).Item(0).ToString

            mySql = "SELECT ENT.ID, ENT.TRANSDATE, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"""
            mySql &= vbCrLf & "FROM POSENTRY ENT "
            mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
            mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO "
            mySql &= vbCrLf & "WHERE "
            mySql &= vbCrLf & String.Format("POSTYPE = 'Sales' AND ITM.ITEMNO <> 'CARD3' AND ITM.POSENTRYID = '{0}'", entry)
            Console.WriteLine("CC: " & mySql)
            objCC = LoadSQL(mySql)
            Console.WriteLine("Output: " & objCC.Tables(0).Rows.Count)


            If objSales.Tables(0).Rows.Count < 1 Then
                If objCC.Tables(0).Rows.Count < 1 Then
                    MsgBox("No Sales recorded", MsgBoxStyle.Information)
                    btnGen.Enabled = True
                    Exit Sub
                End If
            End If

            ' Serial
            mySql = "SELECT ENT.ID, ENT.TRANSDATE, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"""
            mySql &= vbCrLf & "FROM POSENTRY ENT "
            mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
            mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO"
            mySql &= vbCrLf & "WHERE "
            mySql &= vbCrLf & "ITM.ITEMNO <> 'CASHD' AND ITM.ITEMNO <> 'CASHO' AND ITM.SERIALNO <> '' AND "
            mySql &= vbCrLf & String.Format("ENT.TRANSDATE = '{0}'", dtpExtract.Value.ToString("M/d/yyyy"))
            mySql &= vbCrLf & "ORDER BY ITM.ITEMNO ASC"

            objSerial = LoadSQL(mySql)
            GeneratePTUFile(objSales, objSerial, objCC)
            MsgBox("SAP Sales extracted", MsgBoxStyle.Information)
            btnGen.Enabled = True
        Catch ex As Exception
            btnGen.Enabled = True
            MsgBox(ex.Message.ToString, MsgBoxStyle.Information, "Extract Error")
        End Try
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next 'Uncomment on Final

        advertising.InitializedAds()
        wbAds.Navigate(DisplayAds)

        If Not System.IO.File.Exists(configFile) Then
            System.IO.File.Create(configFile).Dispose()

            With iniFile
                .Load(configFile)

                .AddSection("Extractor").AddKey("DB").Value = db_Firebird.dbName
                .AddSection("Extractor").AddKey("Branch").Value = mod_extract.BranchCode
                .AddSection("Extractor").AddKey("Area").Value = mod_extract.AreaCode
                .AddSection("Extractor").AddKey("Customer").Value = mod_extract.CustomerCode

                .Save(configFile)
            End With
        End If

        With iniFile
            .Load(configFile)
            mod_extract.CustomerCode = .GetSection("Extractor").GetKey("Customer").Value
            mod_extract.BranchCode = .GetSection("Extractor").GetKey("Branch").Value
            mod_extract.AreaCode = .GetSection("Extractor").GetKey("Area").Value
            db_Firebird.dbName = .GetSection("Extractor").GetKey("DB").Value
        End With

        txtArea.Text = mod_extract.AreaCode
        txtBranch.Text = mod_extract.BranchCode
        txtCustomer.Text = mod_extract.CustomerCode

        txtArea.ReadOnly = True
        txtBranch.ReadOnly = True
        txtCustomer.ReadOnly = True
    End Sub

    Private Sub wbAds_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wbAds.DocumentCompleted
        pbIT.Visible = True
    End Sub
End Class
