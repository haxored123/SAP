Public Class frmMain
    Dim options As DataSet
    Dim adsShow As Boolean = False

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        StartAds()

        If lvJE.Items.Count <= 0 Then
            Exit Sub
        End If

        Frozen(True)
        CheckingValidity()
        Frozen(False)
        MsgBox("Files are loaded and ready to be converted.", MsgBoxStyle.Information)
    End Sub

    Private Sub AddItem(ByVal branchCode As String, ByVal url As String)
        Console.WriteLine("Branch Code: " & branchCode)
        Console.WriteLine("URL: " & url)

        Dim lv As ListViewItem = lvJE.Items.Add(branchCode)
        lv.SubItems.Add(url)
    End Sub

    Private Sub StartAds()
        If Not adsShow Then
            wbAds.Navigate("http://adf.ly/7104086/banner/google.com")
            adsShow = True
        End If
    End Sub

    Private Sub CheckingValidity()
        Dim hasMisty As Boolean = False

        BranchJElist.Clear()
        Dim confRF As DataSet = loadConfig(SheetNum.RevolvingFunds)
        Dim distRole As DataSet = loadConfig(SheetNum.DistRole)

        For Each JEFile As ListViewItem In lvJE.Items

            If confRF.Tables(fillData).Select("BranchCode = '" & JEFile.Text & "'").Count = 0 Then
                JEFile.BackColor = Color.MistyRose
                hasMisty = True
            Else
                Dim tmpBJE As New BranchJE
                With tmpBJE
                    .BranchCode = JEFile.Text
                    .BranchName = distRole.Tables(fillData).Select("BranchCode = '" & JEFile.Text & "'")(0).Item("Name").ToString
                    .AreaCode = distRole.Tables(fillData).Select("BranchCode = '" & JEFile.Text & "'")(0).Item("Area").ToString
                    .FileURL = JEFile.SubItems(1).Text
                    .DocDate = getDate(getFilename(JEFile.SubItems(1).Text))
                    .RevolvingFund = confRF.Tables(fillData).Select("BranchCode = '" & JEFile.Text & "'")(0).Item("COA").ToString = ""
                End With
                BranchJElist.Add(tmpBJE)
                JEFile.BackColor = Color.LightGreen
            End If
            Application.DoEvents()
        Next
        If hasMisty Then MsgBox("Misty Rose row means invalid entries" & vbCrLf & "Meaning, the Branch Code is missing at the Configuration File and it will be EXCLUDED.", MsgBoxStyle.Information)
    End Sub

    Private Sub Frozen(ByVal st As Boolean)
        Dim s As Boolean = Not st
        btnConvert.Enabled = s
    End Sub

    Private Function getDate(ByVal filename As String) As Date
        Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyyMMdd"}
        Dim edate As String = filename.Substring(4, 8)
        Dim expenddt As Date = Date.ParseExact(edate, format,
            System.Globalization.DateTimeFormatInfo.InvariantInfo,
            Globalization.DateTimeStyles.None)

        Return expenddt
    End Function

    Private Function isValid(ByVal filename As String) As Boolean
        Dim tmpFN As String = getFilename(filename)
        If tmpFN.Substring(0, 4) = "JRNL" Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function getFilename(ByVal str As String) As String
        Dim pos As Integer = str.LastIndexOf("\")
        str = str.Substring(pos + 1) 'Filename
        pos = str.LastIndexOf(".")
        Return str.Substring(0, pos) 'No Extension
    End Function

    Private Function getBranchCode(ByVal filename As String) As String
        Dim tmpCode As String
        filename = getFilename(filename)
        tmpCode = filename.Substring(filename.Length - 3)
        Return tmpCode
    End Function

    Private Function fileExist(ByVal url As String) As Boolean
        If System.IO.File.Exists(url) Then
            Return True
        End If
        Return False
    End Function

    Private Sub loadOption()
        options = loadConfig(SheetNum.Options)
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        mod_excel.ConfigFile = "Configuration.xls"
        loadOption()
        Frozen(1)

        Dim tmpDr() As DataRow
        tmpDr = options.Tables(fillData).Select("Key = 'SaveURL'")
        If tmpDr.Count > 0 Then
            saveAsPath = tmpDr(0).Item("Value").ToString
            If saveAsPath = "" Then saveAsPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End If

        lblSaveURL.Text = saveAsPath
        Console.WriteLine("Path: " & saveAsPath)
    End Sub

    Private Sub btnConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvert.Click
        loadOption()

        Frozen(1)
        ConvertNow()
        Frozen(False)
        MsgBox("All files has been converted", MsgBoxStyle.Information)
    End Sub

    Private Sub btnConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConfig.Click
        If Not fileExist("Configuration.xls") Then
            MsgBox("Configuration file is missing", vbCritical, "ERROR")
            Exit Sub
        End If

        MsgBox("Please edit this file and save it", MsgBoxStyle.Information)
        Process.Start("Configuration.xls")
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Dim result As DialogResult = ofdJE.ShowDialog

        If result = Windows.Forms.DialogResult.OK Then
            For Each fileName In ofdJE.FileNames
                If isValid(fileName) Then AddItem(getBranchCode(fileName), fileName)
            Next
        Else
            Console.WriteLine("Canceled")
        End If
    End Sub

    Private Sub lvJE_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lvJE.KeyDown
        If e.KeyCode = Keys.Delete Then
            For Each i As ListViewItem In lvJE.SelectedItems
                lvJE.Items.Remove(i)
            Next
        End If
    End Sub

    Private Sub wbAds_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wbAds.DocumentCompleted
        tmrAds.Enabled = True
    End Sub

    Private Sub tmrAds_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrAds.Tick
        Static cnt As Integer
        Console.WriteLine("Tick Tok")
        If cnt >= 5 Then
            pbLogo.Visible = False
            tmrAds.Enabled = False
        End If
        cnt += 1
    End Sub
End Class
