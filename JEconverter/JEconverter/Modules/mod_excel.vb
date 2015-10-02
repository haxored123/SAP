Imports Microsoft.Office.Interop 'Add Referrence #COM Microsoft Excel 14.0 Object Library

Module mod_excel
    Public BranchJElist As New BranchJECollection
    Friend ConfigFile As String = String.Empty
    Friend fillData As String = "JournalEntries"
    Friend saveAsPath As String = Application.StartupPath

    Enum SheetNum As Integer
        COA = 1
        DistRole = 2
        RevolvingFunds = 3
        Options = 4
    End Enum

    ''' <summary>
    ''' Get Config Value for a specific Branch
    ''' </summary>
    ''' <param name="branchCode">Three (3) letter code representing the branch</param>
    ''' <param name="configSheet">Identify which Sheet do you want to get the result</param>
    ''' <param name="colNum">Return the column Number entered. Default 2</param>
    ''' <returns>Information of the Branch, 0 if non</returns>
    ''' <remarks></remarks>
    Friend Function getBranchInfo(ByVal branchCode As String, ByVal configSheet As SheetNum, Optional ByVal colNum As Integer = 2) As String
        If Not fileExist(ConfigFile) Then
            MsgBox("Config file missing", MsgBoxStyle.Critical)
            Return 0
        End If

        Dim infoResult As String = Nothing

        'Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oWB = oXL.Workbooks.Open(Application.StartupPath & "\" & ConfigFile)
        oSheet = oWB.Worksheets(configSheet)

        Dim MaxEntries As Integer
        With oSheet
            MaxEntries = .Cells(.Rows.Count, 1).End(Excel.XlDirection.xlUp).row
        End With

        For rowIdx As Integer = 2 To MaxEntries
            If oSheet.Cells(rowIdx, 1).value = branchCode Then
                infoResult = oSheet.Cells(rowIdx, colNum).value
                Exit For
            End If
        Next

        'Memory Unload
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing

        Return infoResult
    End Function

    Friend Sub ConvertNow()
        For Each branch As BranchJE In BranchJElist
            'Create Virtualize Table
            branch.JournalEntries = getJE(branch)

            'Convert
            ConvertingToSAP(branch)
        Next
    End Sub

    Private Function cleanPath(ByVal str As String) As String
        If Right(str, 1) = "\" Then
            str = str.Substring(0, Len(str) - 1)
        End If

        Console.WriteLine("Path: " & str)
        Return str
    End Function

    Private Sub ConvertingToSAP(ByVal dsJE As BranchJE)
        Dim options As DataSet = loadConfig(SheetNum.Options)
        Dim cnt As Integer = 0

        'Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        Try
            oWB = oXL.Workbooks.Open(Application.StartupPath & "\Format.xlsx")
            oSheet = oWB.Worksheets(1) 'Document

            'Adding One Row
            oSheet.Cells(3, 1) = 1 'JDT_NUM
            oSheet.Cells(3, 2) = dsJE.DocDate.ToString("yyyyMMdd") 'RefDate
            oSheet.Cells(3, 8) = dsJE.DocDate.ToString("yyyyMMdd") 'TaxDate
            oSheet.Cells(3, 10) = "tNO" 'UseAutoStorno
            oSheet.Cells(3, 12) = dsJE.DocDate.ToString("yyyyMMdd") 'VatDate
            oSheet.Cells(3, 14) = "tNO" 'StampTax

            oSheet = oWB.Worksheets(2) 'DocumentLines
            Dim ds As DataSet = dsJE.JournalEntries
            Dim MaxRow As Integer = ds.Tables(fillData).Rows.Count
            For cnt = 0 To MaxRow - 1
                With ds.Tables(fillData).Rows(cnt)
                    oSheet.Cells(cnt + 3, 1) = 1 'ParentKey
                    oSheet.Cells(cnt + 3, 2) = cnt 'LineNum
                    oSheet.Cells(cnt + 3, 4) = .Item("AccountNo").ToString 'AccountCode
                    oSheet.Cells(cnt + 3, 5) = .Item("Debit") 'Debit
                    oSheet.Cells(cnt + 3, 6) = .Item("Credit") 'Credit
                    oSheet.Cells(cnt + 3, 19) = dsJE.AreaCode 'ProfitCode
                    oSheet.Cells(cnt + 3, 32) = dsJE.BranchCode 'OcrCode2
                    oSheet.Cells(cnt + 3, 33) = "OPE" 'OcrCode3
                End With

                Application.DoEvents()
            Next

            'Save As
            Dim saveAsFilename As String = "CONV-"
            Dim saveType As String = "Excel"
            If options.Tables.Count > 0 Then
                Dim dr As DataRow = options.Tables(fillData).Select("Key = 'SaveType'")(0)
                saveType = dr.Item("Value")
                saveAsPath = cleanPath(saveAsPath)
                Select Case saveType.ToUpper
                    Case "EXCEL"
                        saveAsFilename &= dsJE.BranchCode & dsJE.DocDate.ToString("MMddyyyy") & ".xlsx"
                        oWB.SaveAs(saveAsPath & "\" & saveAsFilename)
                    Case "TAB"
                        oXL.DisplayAlerts = False

                        oSheet = oWB.Worksheets(1)
                        oSheet.Activate() 'Document

                        saveAsFilename = "doc_" & dsJE.BranchCode & dsJE.DocDate.ToString("MMddyyyy") & ".txt"
                        oWB.SaveAs(saveAsPath & "\" & saveAsFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlCurrentPlatformText)

                        oSheet = oWB.Worksheets(2)
                        oSheet.Activate() 'DocumentsLines
                        saveAsFilename = "docLn_" & dsJE.BranchCode & dsJE.DocDate.ToString("MMddyyyy") & ".txt"
                        oWB.SaveAs(saveAsPath & "\" & saveAsFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlCurrentPlatformText)
                    Case "CSV"
                        oXL.DisplayAlerts = False

                        oSheet = oWB.Worksheets(1)
                        oSheet.Activate() 'Document
                        saveAsFilename = "doc_" & dsJE.BranchCode & dsJE.DocDate.ToString("MMddyyyy") & ".csv"
                        oWB.SaveAs(saveAsPath & "\" & saveAsFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV)

                        oSheet = oWB.Worksheets(2)
                        oSheet.Activate() 'DocumentsLines
                        saveAsFilename = "docLn_" & dsJE.BranchCode & dsJE.DocDate.ToString("MMddyyyy") & ".csv"
                        oWB.SaveAs(saveAsPath & "\" & saveAsFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV)
                End Select
            Else
                saveAsFilename &= dsJE.BranchCode & dsJE.DocDate.ToString("MMddyyyy") & ".xlsx"
                oWB.SaveAs(saveAsPath & "\" & saveAsFilename)
            End If

            'Memory Unload
            oSheet = Nothing
            oWB = Nothing
            oXL.Quit()
            oXL = Nothing
        Catch ex As Exception
            'Memory Unload
            oSheet = Nothing
            oWB = Nothing
            oXL.Quit()
            oXL = Nothing

            MsgBox("Failed to convert SAP File" & vbCr & "FileURL: " & dsJE.FileURL & vbCr & "Row Number: " & cnt & vbCr & _
                   ex.Message.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Function getJE(ByVal branch As BranchJE) As DataSet
        'Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        Dim ds As New DataSet, dt As New DataTable(fillData)
        Dim accCol As DataColumn, debitCol As DataColumn, creditCol As DataColumn
        Dim dsNewRow As DataRow
        Dim rowIdx As Integer = 0

        Try
            oWB = oXL.Workbooks.Open(branch.FileURL)
            oSheet = oWB.Worksheets(1)

            Dim MaxEntries As Integer
            With oSheet
                MaxEntries = .Cells(.Rows.Count, 1).End(Excel.XlDirection.xlUp).row
            End With

            accCol = New DataColumn("AccountNo", GetType(String))
            debitCol = New DataColumn("Debit", GetType(Integer))
            creditCol = New DataColumn("Credit", GetType(Integer))

            ds.Tables.Add(dt)
            ds.Tables(fillData).Columns.Add(accCol)
            ds.Tables(fillData).Columns.Add(debitCol)
            ds.Tables(fillData).Columns.Add(creditCol)

            Dim COA As DataSet = loadConfig(SheetNum.COA)
            Dim Rfund As DataSet = loadConfig(SheetNum.RevolvingFunds)
            Dim tmpStr As String
            Dim tmpDr() As DataRow

            For rowIdx = 2 To MaxEntries
                tmpStr = oSheet.Cells(rowIdx, 3).value 'Change to column 3(AccountName) from column 2(AccountNo)
                If Left(tmpStr, 16) = "REVOLVING FUND -" Then 'Change for Revolving Fund Codes
                    tmpDr = Rfund.Tables(fillData).Select("BranchCode = '" & branch.BranchCode & "'")
                    tmpStr = "NO SAP COA - RF (" & tmpStr & "|" & branch.BranchCode & ")"
                    If tmpDr.Count > 0 Then tmpStr = tmpDr(0).Item("COA")
                Else
                    tmpDr = COA.Tables(fillData).Select("AccountNo = '" & tmpStr & "'")
                    tmpStr = "NO SAP COA(" & tmpStr & ")"
                    If tmpDr.Count > 0 Then tmpStr = tmpDr(0).Item("SAPCOA")
                End If

                dsNewRow = ds.Tables(fillData).NewRow
                With dsNewRow
                    .Item("AccountNo") = tmpStr
                    .Item("Debit") = IIf(oSheet.Cells(rowIdx, 4).value = Nothing, 0, oSheet.Cells(rowIdx, 4).value)
                    .Item("Credit") = IIf(oSheet.Cells(rowIdx, 5).value = Nothing, 0, oSheet.Cells(rowIdx, 5).value)
                End With
                ds.Tables(fillData).Rows.Add(dsNewRow)

                Application.DoEvents()
            Next

            'Memory Unload
            oSheet = Nothing
            oWB = Nothing
            oXL.Quit()
            oXL = Nothing

        Catch ex As Exception
            'Memory Unload
            oSheet = Nothing
            oWB = Nothing
            oXL.Quit()
            oXL = Nothing

            MsgBox("Error Occurred. Please check the following" & vbCrLf & _
                "File: " & branch.FileURL & vbCrLf & "RowNumber: " & rowIdx, MsgBoxStyle.Critical, "GET JE")
            ds = Nothing
        End Try

        Return ds
    End Function

    Private Function fileExist(ByVal url As String) As Boolean
        If System.IO.File.Exists(url) Then
            Return True
        End If
        Return False
    End Function

    Friend Function loadConfig(ByVal sh As SheetNum) As DataSet
        Dim ds As New DataSet, dt As New DataTable(fillData)
        Dim dsNewRow As DataRow

        'Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oWB = oXL.Workbooks.Open(Application.StartupPath & "\" & ConfigFile)
        oSheet = oWB.Worksheets(sh)

        Dim MaxColumn As Integer = 0
        With oSheet
            MaxColumn = .Cells(1, .Columns.Count).End(Excel.XlDirection.xlToLeft).column
        End With

        ds.Tables.Add(dt)
        For colIdx As Integer = 0 To MaxColumn - 1
            Dim tmpCol As DataColumn
            tmpCol = New DataColumn(oSheet.Cells(1, colIdx + 1).value, GetType(String))
            ds.Tables(fillData).Columns.Add(tmpCol)
        Next

        Dim MaxEntries As Integer
        With oSheet
            MaxEntries = .Cells(.Rows.Count, 1).End(Excel.XlDirection.xlUp).row
        End With

        For rowIdx As Integer = 0 To MaxEntries - 1
            dsNewRow = ds.Tables(fillData).NewRow
            With dsNewRow
                For colIdx As Integer = 0 To MaxColumn - 1
                    .Item(colIdx) = IIf(oSheet.Cells(rowIdx + 1, colIdx + 1).value = Nothing, "", oSheet.Cells(rowIdx + 1, colIdx + 1).value)
                Next
            End With
            ds.Tables(fillData).Rows.Add(dsNewRow)
        Next

        'Memory Unload
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing

        Return ds
    End Function
End Module
