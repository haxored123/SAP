'Add Referrence
'COM Microsoft Excel 14.0 Object Library
Imports Microsoft.Office.Interop

Module mod_extract

    Friend fileFormat As String = "fileformat.xlsx"
    Friend CustomerCode As String = "CTPF 90001"
    Friend TransDate As String = "8/13/2015"
    Friend BranchCode As String = "ROX"
    Friend AreaCode As String = "GSC"
    Friend SaveUrl As String = "Desktop"

    Public Sub GeneratePTUFile(ByVal dsTrans As DataSet, ByVal dsSerial As DataSet)
        'Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oWB = oXL.Workbooks.Open(fileFormat)

        'Sheet 1
        oSheet = oWB.Worksheets(0)
        oSheet.Cells(2, 1) = 1
        oSheet.Cells(2, 2) = CustomerCode
        oSheet.Cells(2, 3) = TransDate

        'Sheet 2
        oSheet = oWB.Worksheets(1)
        For cnt As Integer = 0 To dsTrans.Tables(0).Rows.Count - 1
            oSheet.Cells(cnt + 2, 1) = 1
            With dsTrans.Tables(0).Rows(cnt)
                oSheet.Cells(cnt + 2, 2) = .Item("ItemCode")
                oSheet.Cells(cnt + 2, 3) = .Item("ItemName")
                oSheet.Cells(cnt + 2, 4) = .Item("Quantity")
                oSheet.Cells(cnt + 2, 5) = .Item("Price") 'Fetch from DB or recompute
                oSheet.Cells(cnt + 2, 6) = .Item("Discount")
                oSheet.Cells(cnt + 2, 7) = BranchCode
                oSheet.Cells(cnt + 2, 8) = AreaCode
                oSheet.Cells(cnt + 2, 9) = BranchCode
                oSheet.Cells(cnt + 2, 10) = "OPE"
                oSheet.Cells(cnt + 2, 13) = "OVAT-N"
            End With
        Next

        'Sheet 3
        oSheet = oWB.Worksheets(2)
        For cnt = 0 To dsSerial.Tables(0).Rows.Count - 1
            oSheet.Cells(cnt + 2, 1) = 1
            With dsSerial.Tables(0).Rows(cnt)
                oSheet.Cells(cnt + 2, 2) = .Item("ItemCode")
                oSheet.Cells(cnt + 2, 3) = .Item("Quantity")
                oSheet.Cells(cnt + 2, 4) = .Item("WhsCode")
                oSheet.Cells(cnt + 2, 5) = .Item("IntrSerial")
            End With
        Next

        'Save
        oWB.SaveAs(SaveUrl & "\" & BranchCode & CDate(TransDate).ToString("mmddyyyy") & ".PTU")

        'Quit
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing
    End Sub
End Module
