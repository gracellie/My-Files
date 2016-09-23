Imports Microsoft.Office.Interop

Module ExcelManipulator
    Dim dsIMD As New DataSet
    Const MASTERDATA As String = "\TEMPLATES\IMD.xlsx"
    Const TEMPLATE As String = "\TEMPLATES\ESKIE_TEMPLATE.PTU"
    Const LF_RIGHTSALES As Integer = 50 'LOOKING FOR *** TOTAL SALES REPORT BY MENU ITEM *** EXPECTED ROW IS 7
    Dim CUSTOMER As String = ""
    Dim WareHouse As String = ""

    Const WHS As String = "SMI"
    Const OCR1 As String = "VIS"
    Const OCR3 As String = "OPE"
    Const TAX As String = "OVAT-N"

    Friend Sub ConvertToPTU(src As String, dest As String)
        If src = "" Then Exit Sub
        If dest = "" Then Exit Sub

        ' Excel Files
        Dim iSales As Excel.Application
        Dim ItemMasterData As Excel.Application
        Dim PopulateTemplate As New Excel.Application

        Dim saleDate As Date
        Dim MaxEntries As Integer

        'VIRTUALIZING SALES FILE===================
        iSales = OpenExcel(src)
        Dim iBook As Excel.Workbook
        Dim iSheet As Excel.Worksheet

        iBook = iSales.Workbooks(1)
        iSheet = iBook.Worksheets(1)
        'ItemMasterData = OpenExcel(MASTERDATA)
        'PopulateTemplate = OpenExcel(TEMPLATE)

        ' Checking
        ' *** TOTAL SALES REPORT BY MENU ITEM ***
        Dim cnt As Integer = 0
        For cnt = 1 To LF_RIGHTSALES
            Console.WriteLine(String.Format("Row {0}: ", cnt) & iSheet.Cells(cnt, 1).value)
            If IsNothing(iSheet.Cells(cnt, 1).value) Then Continue For

            If iSheet.Cells(cnt, 1).value.ToString = "*** TOTAL SALES REPORT BY MENU ITEM ***" Then
                Console.WriteLine("RIGHT SALES FILE!")
                Console.WriteLine("Row: " & cnt)
                Exit For
            End If
        Next

        If cnt >= LF_RIGHTSALES Then
            CloseExcel(iSales)
            MsgBox("Invalid File", MsgBoxStyle.Critical)
            Exit Sub
        End If

        saleDate = ParsingDate(iSheet.Cells(cnt + 1, 1).value)
        Console.WriteLine("Sales File Date is " & saleDate.ToString("MMM dd, yyyy"))

        Dim dsNewRow As DataRow
        Dim dsSales As New DataSet
        dsSales.Tables.Add(tbl_Import)

        Dim rowSales As Integer = 12
        While Not IsNothing(iSheet.Cells(rowSales, 2).value)
            'VAT 12% Included - Don't Save, End of Line
            If iSheet.Cells(rowSales, 2).value = "VAT 12% Included" Then Exit While

            dsNewRow = dsSales.Tables(0).NewRow
            With dsNewRow
                .Item("Description") = Trim(iSheet.Cells(rowSales, 2).value)
                .Item("Qty") = iSheet.Cells(rowSales, 5).value
                Dim prc As Double = iSheet.Cells(rowSales, 3).value / iSheet.Cells(rowSales, 5).value
                .Item("Price") = prc / 1.12
            End With
            dsSales.Tables(0).Rows.Add(dsNewRow)

            rowSales += 1
        End While

        'SVC 00001 - Service Charge
        dsNewRow = dsSales.Tables(0).NewRow
        With dsNewRow
            .Item("Description") = "Service Charge"
            .Item("Qty") = 1
            .Item("Price") = iSheet.Cells(rowSales + 3, 3).Value
        End With
        dsSales.Tables(0).Rows.Add(dsNewRow)
        'END - VIRTUALIZING SALES FILE===================

        If frmMain.cmbWareHouse.Text = "COFFEE LOVER" Then
            CUSTOMER = "CTGH 00001"
            WareHouse = "COL"
        ElseIf frmMain.cmbWareHouse.Text = "FOODCOURT" Then
            CUSTOMER = "CTGH 00002"
            WareHouse = "FOC"
        ElseIf frmMain.cmbWareHouse.Text = "PTU KTV" Then
            CUSTOMER = "CTGH 00003"
            WareHouse = "PTU KTV"
        ElseIf frmMain.cmbWareHouse.Text = "CHIC-BOY" Then
            CUSTOMER = "CTGH 00004"
            WareHouse = "CHIC"
        ElseIf frmMain.cmbWareHouse.Text = "PBA SPORTS CAFE" Then
            CUSTOMER = "CTGH 00005"
            WareHouse = "PBA"
        ElseIf frmMain.cmbWareHouse.Text = "PTU WAVE RESTO" Then
            CUSTOMER = "CTGH 00006"
            WareHouse = "WAV"
        End If


        'LOADING IMD ==================
        ItemMasterData = OpenExcel(Application.StartupPath & MASTERDATA)
        Dim dsIMD As New DataSet
        dsIMD.Tables.Add(tbl_IMD)

        Dim imdSheet As Excel.Worksheet
        Dim imdBook As Excel.Workbook
        imdBook = ItemMasterData.Workbooks(1)
        imdSheet = imdBook.Worksheets(1)

        With imdSheet
            MaxEntries = .Cells(.Rows.Count, 1).End(Excel.XlDirection.xlUp).row
        End With

        Console.WriteLine("MaxEntries : " & MaxEntries)
  
        For rowIdx As Integer = 2 To MaxEntries
            dsNewRow = dsIMD.Tables(0).NewRow
            With dsNewRow
                .Item("ItemCode") = Trim(imdSheet.Cells(rowIdx, 1).value)
                .Item("Description") = Trim(imdSheet.Cells(rowIdx, 2).value)
                .Item("Branch") = Trim(imdSheet.Cells(rowIdx, 3).value)
            End With
            dsIMD.Tables(0).Rows.Add(dsNewRow)
        Next

      
       
        CloseExcel(ItemMasterData)
        Console.WriteLine("Database Records: " & dsIMD.Tables(0).Rows.Count)
        'END - LOADING IMD ==================

        For j As Integer = dsIMD.Tables(0).Rows.Count - 1 To 0 Step -1
            Dim dr As DataRow = dsIMD.Tables(0).Rows(j)
            If dr("Branch") <> frmMain.cmbWareHouse.Text Then
                dr.Delete()
            End If
        Next


        'POPULATING THE TEMPLATE==================
        PopulateTemplate = OpenExcel(Application.StartupPath & TEMPLATE)
        Dim temSheet As Excel.Worksheet
        Dim temBook As Excel.Workbook
        temBook = PopulateTemplate.Workbooks(1)
        temSheet = temBook.Worksheets(1) 'Sheet 1 - Documents


        'Adding RecordKey
        temSheet.Cells(2, 1).value = 1
        temSheet.Cells(2, 2).value = CUSTOMER
        temSheet.Cells(2, 3).value = saleDate.ToString("M/d/yyyy")

        temSheet = temBook.Worksheets(2) 'Sheet 2 - DocumentLines
        Dim i As Integer = 2
        For Each dr As DataRow In dsSales.Tables(0).Rows
            temSheet.Cells(i, 1).value = 1                                      ' RecordKey
            temSheet.Cells(i, 2).value = getItemCode(dr("Description"), dsIMD)  ' ItemCode
            temSheet.Cells(i, 3).value = dr("Description")                      ' Description
            temSheet.Cells(i, 4).value = dr("Qty")                              ' Qty
            temSheet.Cells(i, 5).value = dr("Price")                            ' Price
            temSheet.Cells(i, 6).value = 0                                      ' Discount
            temSheet.Cells(i, 7).value = WareHouse                              ' WhsCode
            temSheet.Cells(i, 8).value = WareHouse              '               'OcrCode
            temSheet.Cells(i, 9).value = ""                                     ' OcrCode2
            temSheet.Cells(i, 10).value = ""                                    ' OcrCode3
            temSheet.Cells(i, 13).value = TAX                                   ' TaxCode

            i += 1
        Next

        temBook.SaveAs(dest)
        CloseExcel(PopulateTemplate)
        'END - POPULATING THE TEMPLATE==================


        CloseExcel(iSales)
        Console.WriteLine("Finished")
    End Sub

#Region "Initialization"
    Private Function tbl_IMD() As DataTable
        Dim tbl As New DataTable("ITEMMASTER")
        With tbl.Columns
            .Add("ItemCode")
            .Add("Description")
            .Add("Branch")
        End With

        Return tbl
    End Function

    Private Function tbl_Import() As DataTable
        Dim tbl As New DataTable("IMPORT")
        With tbl.Columns
            .Add("Description")
            .Add("Qty")
            .Add("Price")
        End With

        Return tbl
    End Function
#End Region


#Region "Procedures and Functions"
    Private Function OpenExcel(src As String) As Excel.Application
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook

        oWB = oXL.Workbooks.Open(src)

        Return oXL
    End Function

    Private Sub CloseExcel(OpenExcel As Excel.Application)
        OpenExcel.Quit()
        OpenExcel = Nothing

        ReleaseObject(OpenExcel)
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            Dim intRel As Integer = 0
            Do
                intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            Loop While intRel > 0
            'MsgBox("Final Released obj # " & intRel)
        Catch ex As Exception
            'MsgBox("Error releasing object" & ex.ToString)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Function ParsingDate(str As String) As Date
        'From 7/15/2016 To 7/15/2016
        Dim strDate As String = str.Split("To")(0)

        Return strDate.Replace("From ", "")
    End Function

    Private Function getItemCode(ByVal desc As String, ByVal ds As DataSet) As String
        Dim dr() As DataRow = ds.Tables("ITEMMASTER").Select(String.Format("Description = '{0}' ", desc))

        If dr.Count > 1 Then MsgBox("We found multi entries in your IMD", MsgBoxStyle.Critical, dr(0)("Description"))
        If dr.Count = 0 Then Return "* " & desc

        Return dr(0)("Itemcode")
    End Function


#End Region


End Module