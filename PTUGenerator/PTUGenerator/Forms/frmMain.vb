Public Class frmMain

    Friend configFile As String = "Aerauxel.ini"
    Friend iniFile As New IniFile
    Dim adShow As Boolean = False

    Private Sub fetchingSales(ByVal st As Date)
        Try
            Dim dailySalesOR As New DataSet
            Dim mySql As String
            mySql = "SELECT ID, TRANSDATE "
            mySql &= vbCr & "FROM POSENTRY  "
            mySql &= vbCr & String.Format("WHERE POSTYPE = 'Sales' AND TRANSDATE = '{0}'", st.ToString("M/d/yyyy"))

            dailySalesOR = LoadSQL(mySql, "GetAll")
            Dim MaxRow As Integer = dailySalesOR.Tables("GetAll").Rows.Count
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

                Application.DoEvents()
            Next

            Console.WriteLine("Tables: " & dailySalesOR.Tables.Count)
            Dim str As String = ""
            For cnt As Integer = 0 To dailySalesOR.Tables.Count - 1
                str &= dailySalesOR.Tables(cnt).TableName & vbTab
            Next
            Console.WriteLine(str)

            'Display DataSets
            'Dim MaxTableCnt As Integer = dailySalesOR.Tables.Count
            'For idx = 0 To MaxTableCnt - 1
            '    Console.WriteLine("Table: " & dailySalesOR.Tables(idx).TableName)
            '    Console.WriteLine("Number of Column: " & dailySalesOR.Tables(idx).Columns.Count)

            '    For RoxIdx As Integer = 0 To dailySalesOR.Tables(idx).Rows.Count - 1
            '        With dailySalesOR.Tables(idx).Rows(RoxIdx)
            '            Dim str As String = ""
            '            For ColIdx As Integer = 0 To dailySalesOR.Tables(idx).Columns.Count - 1
            '                str &= .Item(ColIdx) & vbTab
            '            Next
            '            Console.WriteLine(str)
            '        End With
            '    Next
            'Next

            mod_extract.GeneratePTUFileV2(dailySalesOR)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Fetching")
        End Try
    End Sub

    Private Sub btnGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGen.Click
        mod_extract.TransDate = dtpExtract.Value.ToString("M/d/yyyy")
        If devMode Then mod_extract.TransDate = "9/28/2015"
        fetchingSales(mod_extract.TransDate)

        Exit Sub
        Try
            btnGen.Enabled = False
            tmrAds.Enabled = True

            Dim objSales As New DataSet
            Dim objSerial As New DataSet
            Dim tmpCC As New DataSet
            Dim objCC As New DataSet
            mod_extract.TransDate = dtpExtract.Value.ToString("M/d/yyyy")

            'Sales
            Dim mySql As String
            mySql = "SELECT ENT.ID, ENT.TRANSDATE, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"""
            mySql &= vbCrLf & "FROM POSENTRY ENT "
            mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
            mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO"
            mySql &= vbCrLf & "WHERE "
            mySql &= vbCrLf & "POSTYPE = 'Sales' AND ITM.QTY > 0"
            mySql &= vbCrLf & String.Format("ENT.TRANSDATE = '{0}'", mod_extract.TransDate)
            mySql &= vbCrLf & "ORDER BY ITM.ITEMNO ASC"
            objSales = LoadSQL(mySql)
            Console.WriteLine("Output: " & objSales.Tables(0).Rows.Count & vbCrLf & mySql)

            'For no ITEMNAME

            'For CC
            'mySql = "SELECT ITM.POSENTRYID, ENT.NAME"
            'mySql &= vbCrLf & "FROM POSENTRY ENT "
            'mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
            'mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO"
            'mySql &= vbCrLf & "WHERE "
            'mySql &= vbCrLf & String.Format("POSTYPE = 'Sales' AND ITM.ITEMNO = 'CARD3' AND ENT.TRANSDATE = '{0}'", mod_extract.TransDate)
            'tmpCC = LoadSQL(mySql)

            'Dim entry As String = "0"
            'If tmpCC.Tables(0).Rows.Count = 1 Then entry = tmpCC.Tables(0).Rows(0).Item(0).ToString

            'mySql = "SELECT ENT.ID, ENT.TRANSDATE, ITM.ITEMNO as ""ItemCode"", ITMM.ITEMNAME as ""ItemName"", ITM.QTY as ""Quantity"", ITM.UNITPRICE as ""Price"", ITM.SERIALNO as ""IntrSerial"""
            'mySql &= vbCrLf & "FROM POSENTRY ENT "
            'mySql &= vbCrLf & "INNER JOIN POSITEM ITM ON ENT.ID = ITM.POSENTRYID"
            'mySql &= vbCrLf & "LEFT JOIN ITEMMASTER ITMM ON ITM.ITEMNO = ITMM.ITEMNO "
            'mySql &= vbCrLf & "WHERE "
            'mySql &= vbCrLf & String.Format("POSTYPE = 'Sales' AND ITM.ITEMNO <> 'CARD3' AND ITM.POSENTRYID = '{0}'", entry)
            'Console.WriteLine("CC: " & mySql)
            'objCC = LoadSQL(mySql)
            'Console.WriteLine("Output: " & objCC.Tables(0).Rows.Count)


            'If objSales.Tables(0).Rows.Count < 1 Then
            '    If objCC.Tables(0).Rows.Count < 1 Then
            '        MsgBox("No Sales recorded", MsgBoxStyle.Information)
            '        btnGen.Enabled = True
            '        Exit Sub
            '    End If
            'End If

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
        If devMode Then MsgBox("DEVELOPER MODE", MsgBoxStyle.Information)

        On Error Resume Next 'Uncomment on Final

        advertising.InitializedAds()
        wbAds.Navigate("http://adf.ly/7104086/banner/pgc-itdept.org/software/ptu-generator/")

        LoadConfig()

        txtArea.Text = mod_extract.AreaCode
        txtBranch.Text = mod_extract.BranchCode

        Me.Text = My.Application.Info.Title & "|" & mod_extract.Company & " by IT Department 2015"

        txtArea.ReadOnly = True
        txtBranch.ReadOnly = True
        txtCustomer.ReadOnly = True
    End Sub

    Private Sub LoadConfig()
        Try
            If Not System.IO.File.Exists(configFile) Then
                System.IO.File.Create(configFile).Dispose()

                With iniFile
                    .Load(configFile)

                    .AddSection("Extractor").AddKey("DB").Value = db_Firebird.dbName
                    .AddSection("Extractor").AddKey("Branch").Value = mod_extract.BranchCode
                    .AddSection("Extractor").AddKey("Area").Value = mod_extract.AreaCode
                    .AddSection("Extractor").AddKey("Customer").Value = mod_extract.CustomerCode
                    .AddSection("Extractor").AddKey("Company").Value = mod_extract.Company

                    .Save(configFile)
                End With
            End If

            With iniFile
                .Load(configFile)
                mod_extract.CustomerCode = .GetSection("Extractor").GetKey("Customer").Value
                mod_extract.BranchCode = .GetSection("Extractor").GetKey("Branch").Value
                mod_extract.AreaCode = .GetSection("Extractor").GetKey("Area").Value
                mod_extract.Company = .GetSection("Extractor").GetKey("Company").Value
                db_Firebird.dbName = .GetSection("Extractor").GetKey("DB").Value
            End With
        Catch ex As Exception
            MsgBox("Configuration Corrupted" & vbCrLf & ex.Message.ToString & vbCr & "System Closing...", MsgBoxStyle.Critical, "Load Config Failed")
            End
        End Try
    End Sub

    Private Sub wbAds_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wbAds.DocumentCompleted
        adShow = True
    End Sub

    Private Sub tmrAds_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrAds.Tick
        Static cnt As Integer
        cnt += 1
        If cnt >= 5 Then
            If Not adShow Then
                tmrAds.Enabled = False
                pbIT.Visible = False
            End If
            cnt = 0
        End If
    End Sub
End Class
