'Add Referrence
'COM Microsoft Excel 14.0 Object Library
Imports Microsoft.Office.Interop

Module mod_extract
    Public devMode As Boolean = True

    Friend fileFormat As String = "fileformat.xlsx"
    Friend CustomerCode As String = "CTPF 90001"
    Friend CreditCardCode As String = "CTPF 90050"
    Friend TransDate As String = "12/12/2014"
    Friend BranchCode As String = "ROG"
    Friend AreaCode As String = "GSC"
    Friend Company As String = "Perfecom"
    Friend SaveUrl As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

    Private AccntCodes() As String

    ''' <summary>
    ''' Extract OR SALES FILES from POS
    ''' </summary>
    ''' <param name="dsTrans">Dataset of Sales</param>
    ''' <param name="dsSerial">Dataset of Serials</param>
    ''' <param name="dsCC">Dataset of CC</param>
    ''' <remarks></remarks>
    Public Sub GeneratePTUFile(ByVal dsTrans As DataSet, ByVal dsSerial As DataSet, ByVal dsCC As DataSet)
        'Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oWB = oXL.Workbooks.Open(Application.StartupPath() & "\" & fileFormat)

        'Sheet 1
        oSheet = oWB.Worksheets(1)

        'Sales
        Dim salesCnt As Integer = 1
        If dsTrans.Tables(0).Rows.Count > 0 Then
            oSheet.Cells(salesCnt + 1, 1) = salesCnt
            oSheet.Cells(salesCnt + 1, 2) = CustomerCode
            oSheet.Cells(salesCnt + 1, 3) = TransDate

            salesCnt += 1
        End If

        'CC
        If dsCC.Tables(0).Rows.Count > 0 Then
            oSheet.Cells(salesCnt + 1, 1) = salesCnt
            oSheet.Cells(salesCnt + 1, 2) = CreditCardCode
            oSheet.Cells(salesCnt + 1, 3) = TransDate
        End If

        'Sheet 2
        oSheet = oWB.Worksheets(2)
        For cnt As Integer = 0 To dsTrans.Tables(0).Rows.Count - 1 'Sales
            oSheet.Cells(cnt + 2, 1) = 1

            With dsTrans.Tables(0).Rows(cnt)
                Dim prc As Double = .Item("Price") / 1.12 'remove VAT
                oSheet.Cells(cnt + 2, 2) = .Item("ItemCode")
                oSheet.Cells(cnt + 2, 3) = .Item("ItemName")
                oSheet.Cells(cnt + 2, 4) = .Item("Quantity")
                oSheet.Cells(cnt + 2, 5) = prc 'Fetch from DB or recompute
                oSheet.Cells(cnt + 2, 6) = 0 '.Item("Discount") 'Change Discount into 0. to be verified where should it came from.
                oSheet.Cells(cnt + 2, 7) = BranchCode 'Perfecom BBA|Photo BAB
                oSheet.Cells(cnt + 2, 8) = IIf(Company = "Perfecom", BranchCode, AreaCode)
                oSheet.Cells(cnt + 2, 9) = IIf(Company = "Perfecom", AreaCode, BranchCode)
                oSheet.Cells(cnt + 2, 10) = "OPE"
                oSheet.Cells(cnt + 2, 13) = "OVAT-N"
            End With
            Application.DoEvents() 'No Hang
        Next
        Dim ccCnt As Integer = dsTrans.Tables(0).Rows.Count
        For cnt = 0 To dsCC.Tables(0).Rows.Count - 1 'CC
            oSheet.Cells(ccCnt + cnt, 1) = 2 'Next recordKey
            With dsCC.Tables(0).Rows(cnt)
                Dim prc As Double = .Item("Price") / 1.12 'remove VAT
                oSheet.Cells(ccCnt + cnt, 2) = .Item("ItemCode")
                oSheet.Cells(ccCnt + cnt, 3) = .Item("ItemName")
                oSheet.Cells(ccCnt + cnt, 4) = .Item("Quantity")
                oSheet.Cells(ccCnt + cnt, 5) = prc 'Fetch from DB or recompute
                oSheet.Cells(ccCnt + cnt, 6) = 0 '.Item("Discount") 'Change Discount into 0. to be verified where should it came from.
                oSheet.Cells(ccCnt + cnt, 7) = BranchCode
                oSheet.Cells(ccCnt + cnt, 8) = AreaCode
                oSheet.Cells(ccCnt + cnt, 9) = BranchCode
                oSheet.Cells(ccCnt + cnt, 10) = "OPE"
                oSheet.Cells(ccCnt + cnt, 13) = "OVAT-N"
            End With
            Application.DoEvents() 'No Hang
        Next

        'Sheet 3
        oSheet = oWB.Worksheets(3)
        For cnt = 0 To dsSerial.Tables(0).Rows.Count - 1
            oSheet.Cells(cnt + 2, 1) = 1
            With dsSerial.Tables(0).Rows(cnt)
                Dim tmpStr As String = .Item("IntrSerial")

                If .Item("IntrSerial").ToString.Length > 0 Then Console.WriteLine("Serial Found")
                If IsNumeric(.Item("IntrSerial")) Then
                    tmpStr = "'" & .Item("IntrSerial")
                End If
                oSheet.Cells(cnt + 2, 2) = .Item("ItemCode")
                oSheet.Cells(cnt + 2, 3) = .Item("Quantity")
                oSheet.Cells(cnt + 2, 4) = BranchCode
                oSheet.Cells(cnt + 2, 5) = tmpStr
            End With
        Next

        'Switch to Sheet 1
        oSheet = oWB.Worksheets(1)
        'Save
        oWB.SaveAs(SaveUrl & "\" & BranchCode & CDate(TransDate).ToString("MMddyyyy") & ".PTU")

        'Quit
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing
    End Sub

    Friend Sub GeneratePTUFileV2(ByVal dsSales As DataSet)
        Dim recordKey As Integer = 1
        'Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        Try
            dsSales.Tables.Remove("GetAll")
            GetCodes(dsSales)
            oWB = oXL.Workbooks.Open(Application.StartupPath() & "\" & fileFormat)

            'Sheet 1
            oSheet = oWB.Worksheets(1)
            For Each code In AccntCodes
                With oSheet
                    .Cells(recordKey + 1, 1) = recordKey
                    .Cells(recordKey + 1, 2) = code
                    .Cells(recordKey + 1, 3) = TransDate
                End With
                recordKey += 1
            Next

            'Seet 2
            oSheet = oWB.Worksheets(2)
            Dim rowCnt As Integer = 0
            For tblCnt As Integer = 0 To dsSales.Tables.Count - 1
                For rowIdx As Integer = 0 To dsSales.Tables(tblCnt).Rows.Count - 1
                    oSheet.Cells(rowCnt + 2, 1) = tblCnt + 1 'RecordKey
                    With dsSales.Tables(tblCnt).Rows(rowIdx)
                        Dim prc As Double = .Item("Price") / 1.12
                        oSheet.Cells(rowCnt + 2, 2) = .Item("ItemCode")
                        oSheet.Cells(rowCnt + 2, 3) = .Item("ItemName")
                        oSheet.Cells(rowCnt + 2, 4) = .Item("Quantity")
                        oSheet.Cells(rowCnt + 2, 5) = prc 'Fetch from DB or recompute
                        oSheet.Cells(rowCnt + 2, 6) = 0 '.Item("Discount") 'Change Discount into 0. to be verified where should it came from.
                        oSheet.Cells(rowCnt + 2, 7) = BranchCode
                        oSheet.Cells(rowCnt + 2, 8) = IIf(Company = "Perfecom", BranchCode, AreaCode)
                        oSheet.Cells(rowCnt + 2, 9) = IIf(Company = "Perfecom", AreaCode, BranchCode)
                        oSheet.Cells(rowCnt + 2, 10) = "OPE"
                        oSheet.Cells(rowCnt + 2, 13) = "OVAT-N"
                    End With

                    rowCnt += 1
                    Console.WriteLine("RowNo: " & rowCnt)
                Next
            Next

            'Sheet 3 - Serial
            oSheet = oWB.Worksheets(3)
            rowCnt = 0
            For tblCnt = 0 To dsSales.Tables.Count - 1
                Dim result() As DataRow = dsSales.Tables(tblCnt).Select("IntrSerial <> ''")
                Console.WriteLine("hasSerial: " & result.Count)
                For rowIdx As Integer = 0 To result.Count - 1
                    With result(rowIdx)
                        oSheet.Cells(rowCnt + 2, 1) = tblCnt + 1 'RecordKey
                        oSheet.Cells(rowCnt + 2, 2) = .Item("ItemCode") 'ItemCode
                        oSheet.Cells(rowCnt + 2, 3) = .Item("Quantity") 'Quantity
                        oSheet.Cells(rowCnt + 2, 4) = BranchCode  'WhsCode
                        oSheet.Cells(rowCnt + 2, 5) = _
                            If(IsNumeric(.Item("IntrSerial")), "'" & .Item("IntrSerial"), .Item("IntrSerial")) 'IntrSerial
                    End With
                    rowCnt += 1
                Next
            Next

            oSheet = oWB.Worksheets(1)
            oSheet.Activate()

            'Save
            oWB.SaveAs(SaveUrl & "\" & BranchCode & CDate(TransDate).ToString("MMddyyyy") & ".PTU")
            MsgBox("DONE!", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox(ex.Message.ToUpper, MsgBoxStyle.Critical, "Error Generating")
        End Try

        'Quit
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing
    End Sub

    Private Sub GetCodes(ByVal dsSales As DataSet)
        On Error Resume Next

        With frmMain.iniFile
            .Load(frmMain.configFile)
            ReDim AccntCodes(dsSales.Tables.Count - 1)

            For cnt As Integer = 0 To dsSales.Tables.Count - 1
                Dim tblName As String = dsSales.Tables(cnt).TableName
                AccntCodes(cnt) = .GetSection("CODES").GetKey(tblName).Value
                If AccntCodes(cnt) = Nothing Then
                    AccntCodes(cnt) = dsSales.Tables(cnt).TableName
                End If
            Next
        End With
    End Sub
End Module
