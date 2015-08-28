Public Class frmMain

    Private configFile As String = "Aerauxel.ini"
    Private iniFile As New IniFile

    Private Sub btnGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGen.Click
        Dim objSales As New DataSet
        Dim objSerial As New DataSet
        mod_extract.TransDate = dtpExtract.Value.ToString("M/d/yyyy")
        If devMode Then mod_extract.TransDate = "6/3/2015"

        Dim mySql As String
        mySql = "SELECT ENT.ID, ENT.TRANSDATE, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"""
        mySql &= vbCrLf & "FROM POSENTRY ENT "
        mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
        mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO"
        mySql &= vbCrLf & "WHERE "
        mySql &= vbCrLf & "ITM.ITEMNO <> 'CASHD' AND ITM.ITEMNO <> 'CASHO' AND"
        mySql &= vbCrLf & String.Format("ENT.TRANSDATE = '{0}'", mod_extract.TransDate)
        mySql &= vbCrLf & "ORDER BY ITM.ITEMNO ASC"

        Console.WriteLine("SQL: " & mySql)
        objSales = LoadSQL(mySql)
        If objSales.Tables(0).Rows.Count < 1 Then
            MsgBox("No Sales recorded", MsgBoxStyle.Information)
            Exit Sub
        End If

        mySql = "SELECT ENT.ID, ENT.TRANSDATE, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"""
        mySql &= vbCrLf & "FROM POSENTRY ENT "
        mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
        mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO"
        mySql &= vbCrLf & "WHERE "
        mySql &= vbCrLf & "ITM.ITEMNO <> 'CASHD' AND ITM.ITEMNO <> 'CASHO' AND ITM.SERIALNO <> '' AND "
        mySql &= vbCrLf & String.Format("ENT.TRANSDATE = '{0}'", dtpExtract.Value.ToString("M/d/yyyy"))
        mySql &= vbCrLf & "ORDER BY ITM.ITEMNO ASC"

        objSerial = LoadSQL(mySql)
        GeneratePTUFile(objSales, objSerial)
        MsgBox("SAP Sales extracted", MsgBoxStyle.Information)
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next 'Uncomment on Final

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
End Class
