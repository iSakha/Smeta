Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO
Module myFunc
    '===================================================================================
    '             === check expiration date ===
    '===================================================================================
    Sub checkExpirationDate()

        Dim currentDate As Date = Date.Now
        Dim lastRunDate As Date = My.Settings.lastRun
        Dim daysStayed As Int32 = My.Settings.expireDate.Subtract(currentDate).Days

        mainForm.menuItem_department.Enabled = False
        mainForm.menuItem_company.Enabled = False

        If lastRunDate.Subtract(currentDate).Days > 0 Then
            MsgBox("Check date and time settings!")
            mainForm.Close()
        Else
            My.Settings.lastRun = currentDate
            My.Settings.Save()
        End If

        If daysStayed > 0 Then
            Return
        Else
            MsgBox("This app has expired!")
            mainForm.Close()
        End If
    End Sub

    '===================================================================================
    '             === Load database ===
    '===================================================================================

    Sub load_db()

        Dim key As Integer = 0

        Select Case mainForm.cancelFlag

            Case False

                ' Show the FolderBrowserDialog.
                mainForm.FBD.SelectedPath = Directory.GetCurrentDirectory()
                Dim result As DialogResult = mainForm.FBD.ShowDialog()
                If (result = DialogResult.OK) Then
                    mainForm.sDir = mainForm.FBD.SelectedPath
                Else
                    mainForm.cancelFlag = True
                End If

                mainForm.i_superPivotDict = New Dictionary(Of Integer, Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable)))
                mainForm.i_pivotTableDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))
                mainForm.i_pivot_wsDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelWorksheet))


                mainForm.filePath = New Collection
                mainForm.fileNames = New Collection

                Try
                    For Each foundFile In My.Computer.FileSystem.GetFiles _
            (mainForm.sDir, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.omdb")
                        'Console.WriteLine(foundFile)
                        mainForm.filePath.Add(foundFile)
                        'Console.WriteLine(foundFile)
                        Dim dIndex = StrReverse(foundFile).IndexOf("\")

                        Dim name As String
                        name = Right(foundFile, dIndex)

                        name = Left(name, Len(name) - 5)
                        mainForm.fileNames.Add(name)

                    Next

                Catch
                End Try

                For Each fPath As String In mainForm.filePath

                    '   Create collection of Excel files workSheets

                    Dim ws As ExcelWorksheet
                    Dim excelFile = New FileInfo(fPath)
                    'Console.WriteLine(mainForm.sFilePath)
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                    Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

                    key = key + 1

                    mainForm.i_wsDict = New Dictionary(Of Integer, ExcelWorksheet)

                    For i As Integer = 0 To Excel.Workbook.Worksheets.Count - 1

                        mainForm.i_xlTableDict = New Dictionary(Of Integer, ExcelTable)

                        ws = Excel.Workbook.Worksheets(i)

                        mainForm.i_wsDict.Add(i, ws)

                        Dim k As Integer = 0
                        For Each tbl As ExcelTable In ws.Tables

                            mainForm.i_xlTableDict.Add(k, tbl)
                            k = k + 1
                        Next tbl

                        mainForm.i_pivotTableDict.Add(i, mainForm.i_xlTableDict)
                    Next i

                    mainForm.i_pivot_wsDict.Add(key - 1, mainForm.i_wsDict)
                    mainForm.i_superPivotDict.Add(key - 1, mainForm.i_pivotTableDict)
                    mainForm.i_pivotTableDict = New Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))

                Next fPath
        End Select
    End Sub

    '===================================================================================
    '             === Create dataset ===
    '===================================================================================
    Sub create_dataset(_iDEpartment As Integer, _iCategory As Integer)

        Dim dt As DataTable

        Dim xlTable As ExcelTable
        Dim adr As String
        Dim row As DataRow
        Dim ws As ExcelWorksheet
        Dim r_xlTable, c_xlTable As Integer
        Dim rng As ExcelRange


        mainForm.dts = New DataSet

        ws = mainForm.i_pivot_wsDict(_iDEpartment)(_iCategory)

        For k As Integer = 0 To ws.Tables.Count - 1

            xlTable = ws.Tables(k)
            c_xlTable = xlTable.Address.Columns
            r_xlTable = xlTable.Address.Rows

            adr = xlTable.Address.Address
            rng = ws.Cells(adr)

            Select Case k

                Case = 0

                    dt = New DataTable
                    dt.TableName = xlTable.Name

                    'Adding the Columns
                    For i = 0 To c_xlTable - 1
                        dt.Columns.Add(rng.Value(0, i))
                    Next i

                    dt.Columns(0).DataType = System.Type.GetType("System.Int32")               ' #
                    dt.Columns(1).DataType = System.Type.GetType("System.String")              ' Fixture
                    dt.Columns(2).DataType = System.Type.GetType("System.Int32")               ' Q-ty
                    dt.Columns(3).DataType = System.Type.GetType("System.Int32")               ' BelImlight
                    dt.Columns(4).DataType = System.Type.GetType("System.Int32")               ' PRLightigTouring
                    dt.Columns(5).DataType = System.Type.GetType("System.Int32")               ' BlackOut
                    dt.Columns(6).DataType = System.Type.GetType("System.Int32")               ' Vision
                    dt.Columns(7).DataType = System.Type.GetType("System.Int32")               ' Stage
                    dt.Columns(8).DataType = System.Type.GetType("System.Int32")               ' Weight
                    If mainForm.iDepartment = 3 Then
                        dt.Columns(9).DataType = System.Type.GetType("System.String")          ' Power/length
                    Else
                        dt.Columns(9).DataType = System.Type.GetType("System.Int32")           ' Power/length
                    End If

                    dt.Columns(10).DataType = System.Type.GetType("System.Int32")              ' Price
                    dt.Columns.Add()
                    dt.Columns(11).DataType = System.Type.GetType("System.Int32")              ' Result
                    dt.Columns(11).ColumnName = "Result"


                    For i = 1 To r_xlTable - 1

                        row = dt.Rows.Add()

                        For j = 0 To c_xlTable - 1

                            row.Item(j) = rng.Value(i, j)

                        Next j

                        Dim val, val_bel, val_pr, val_black, val_vis, val_st As Integer

                        val = row.Item(2)
                        val_bel = row.Item(3)
                        val_pr = row.Item(4)
                        val_black = row.Item(5)
                        val_vis = row.Item(6)
                        val_st = row.Item(7)

                        row.Item(c_xlTable) = val - (val_bel + val_pr + val_black + val_vis + val_st)

                    Next i

                Case > 0

                    dt = New DataTable
                    dt.TableName = xlTable.Name

                    'Adding the Columns
                    For i = 0 To c_xlTable - 1
                        dt.Columns.Add(rng.Value(0, i))
                    Next i

                    dt.Columns(0).DataType = System.Type.GetType("System.Int32")
                    dt.Columns(1).DataType = System.Type.GetType("System.String")
                    dt.Columns(2).DataType = System.Type.GetType("System.Int32")
                    dt.Columns(3).DataType = System.Type.GetType("System.String")
                    dt.Columns(4).DataType = System.Type.GetType("System.Int32")
                    dt.Columns(5).DataType = System.Type.GetType("System.String")
                    dt.Columns(6).DataType = System.Type.GetType("System.Int32")
                    dt.Columns(7).DataType = System.Type.GetType("System.String")
                    dt.Columns(8).DataType = System.Type.GetType("System.Int32")


                    'Add Rows from Excel table

                    For i = 1 To r_xlTable - 1
                        row = dt.Rows.Add()

                        For j = 0 To c_xlTable - 1

                            If rng.Value(i, j) = Nothing Then
                                Select Case j
                                    Case 3
                                        row.Item(j) = ""
                                    Case 4
                                        row.Item(j) = 0
                                    Case 5
                                        row.Item(j) = ""
                                    Case 6
                                        row.Item(j) = 0
                                    Case 7
                                        row.Item(j) = ""
                                    Case 8
                                        row.Item(j) = 0
                                End Select
                            Else
                                row.Item(j) = rng.Value(i, j)
                            End If

                        Next j
                    Next i

            End Select

            mainForm.dts.Tables.Add(dt)
        Next k

    End Sub

    '===================================================================================
    '             === fillDGV ===
    '===================================================================================
    Sub fillDGV(_department As String, _sender As Object)

        create_dataset(mainForm.iDepartment, mainForm.iCategory)
        mainForm.DGV.DataSource = mainForm.dts.Tables(0)
        format_DGV()

    End Sub

    '===================================================================================
    '             === Format DGV ===
    '===================================================================================

    Sub format_DGV()

        Dim col() As Color

        col = {Color.FromArgb(252, 228, 214), Color.FromArgb(221, 235, 247), Color.FromArgb(237, 237, 237),
            Color.FromArgb(226, 239, 218), Color.FromArgb(237, 226, 246)}

        mainForm.DGV.Columns(0).Width = 55                ' #
        mainForm.DGV.Columns(1).Width = 230               ' Fixture
        mainForm.DGV.Columns(2).Width = 65                ' Q-ty
        mainForm.DGV.Columns(2).DefaultCellStyle.Font = New Font("Tahoma", 10)
        mainForm.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(3).Width = 62                ' BelImlight
        mainForm.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(4).Width = 62                ' PRLightigTouring
        mainForm.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(5).Width = 62                ' BlackOut
        mainForm.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(6).Width = 62                ' Vision
        mainForm.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(7).Width = 62                ' Stage
        mainForm.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        mainForm.DGV.Columns(8).Width = 48                ' Weight
        mainForm.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(9).Width = 48                ' Power/length
        mainForm.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(10).Width = 48                ' Price
        mainForm.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'mainForm.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'mainForm.DGV.Columns(11).Width = 65
        'mainForm.DGV.Columns(11).DefaultCellStyle.Font = New Font("Tahoma", 10, FontStyle.Bold)

        mainForm.DGV.Columns(11).Visible = False

        mainForm.DGV.Columns(3).DefaultCellStyle.BackColor = col(0)
        mainForm.DGV.Columns(4).DefaultCellStyle.BackColor = col(1)
        mainForm.DGV.Columns(5).DefaultCellStyle.BackColor = col(2)
        mainForm.DGV.Columns(6).DefaultCellStyle.BackColor = col(3)
        mainForm.DGV.Columns(7).DefaultCellStyle.BackColor = col(4)

        For i = 0 To mainForm.DGV.Rows.Count - 2

            mainForm.DGV.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(237, 237, 250)

        Next i
    End Sub

    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Sub dgv_clickCell(_sender As Object, _e As DataGridViewCellEventArgs)

        Dim index As Integer
        index = _e.RowIndex
        mainForm.selIndex = index
        'Console.WriteLine(_e)
        Dim selectedRow As DataGridViewRow
        selectedRow = _sender.Rows(index)

        mainForm.DGV.Rows(index).Selected = True
        mainForm.btn_store_qty.Text = mainForm.DGV.Rows(index).Cells(2).Value
        'Console.WriteLine(mainForm.DGV.Rows(index).Cells(2).Value)

    End Sub
    '===================================================================================
    '             === Create Smeta dataTable ===
    '===================================================================================

    Sub createSmeta_dt()

        Dim dtExist As Boolean = False

        Dim dt As DataTable
        dt = New DataTable
        'add three colums to the datatable
        dt.Columns.Add("#")
        dt.Columns.Add("Fixture")
        dt.Columns.Add("Qty")
        dt.Columns.Add("BelImlight")
        dt.Columns.Add("PRLightigTouring")
        dt.Columns.Add("BlackOut")
        dt.Columns.Add("Vision")
        dt.Columns.Add("Stage")
        dt.Columns.Add("Weight")
        dt.Columns.Add("Power")
        dt.Columns.Add("Price")
        dt.Columns.Add("Result")


        dt.Columns(0).DataType = System.Type.GetType("System.Int32")               ' #
        dt.Columns(1).DataType = System.Type.GetType("System.String")              ' Fixture
        dt.Columns(2).DataType = System.Type.GetType("System.Int32")               ' Q-ty
        dt.Columns(3).DataType = System.Type.GetType("System.Int32")               ' BelImlight
        dt.Columns(4).DataType = System.Type.GetType("System.Int32")               ' PRLightigTouring
        dt.Columns(5).DataType = System.Type.GetType("System.Int32")               ' BlackOut
        dt.Columns(6).DataType = System.Type.GetType("System.Int32")               ' Vision
        dt.Columns(7).DataType = System.Type.GetType("System.Int32")               ' Stage
        dt.Columns(8).DataType = System.Type.GetType("System.Int32")               ' Weight

        If mainForm.iDepartment = 3 Then
            dt.Columns(9).DataType = System.Type.GetType("System.String")          ' Power/length
        Else
            dt.Columns(9).DataType = System.Type.GetType("System.Int32")           ' Power/length
        End If

        dt.Columns(10).DataType = System.Type.GetType("System.Int32")              ' Price
        dt.Columns.Add()
        dt.Columns(11).DataType = System.Type.GetType("System.Int32")              ' Result
        dt.Columns(11).ColumnName = "Result"

        dt.TableName = mainForm.dts.Tables(0).TableName

        For Each tbl As DataTable In mainForm.smeta_dts.Tables
            If tbl.TableName = dt.TableName Then
                dtExist = True
            End If
        Next tbl

        If Not dtExist Then
            mainForm.smeta_dts.Tables.Add(dt)
        End If
        dtExist = False

        addItem()
        addRow(dt.TableName)

    End Sub

    '===================================================================================
    '             === Add row to Smeta dataTable ===
    '===================================================================================
    Sub addRow(_tName As String)

        mainForm.smeta_dts.Tables(_tName).ImportRow(mainForm.dts.Tables(0).Rows(mainForm.selIndex))

        'smetaForm.DGV_smeta_1.DataSource = mainForm.smeta_dts.Tables(_tName)

    End Sub

    '===================================================================================
    '             === Add items to Smeta menu ===
    '===================================================================================
    Sub addItem()

        Dim itemExist, subItemExist As Boolean
        itemExist = False
        subItemExist = False

        Dim s_item, s_subitem As String
        s_item = mainForm.sDepartment(mainForm.iDepartment)
        s_subitem = mainForm.i_pivot_wsDict(mainForm.iDepartment)(mainForm.iCategory).Name

        Dim item As New ToolStripMenuItem(s_item)
        Dim subItem As New ToolStripMenuItem(s_subitem)

        Dim index As Integer

        If smetaForm.menuStripCat.DropDownItems.Count = 0 Then
            smetaForm.menuStripCat.DropDownItems.Add(item)
        End If

        For Each itm As ToolStripMenuItem In smetaForm.menuStripCat.DropDownItems
            If itm.ToString = s_item Then
                itemExist = True
            End If
        Next itm

        If Not itemExist Then
            smetaForm.menuStripCat.DropDownItems.Add(item)
        End If


        For Each itm As ToolStripMenuItem In smetaForm.menuStripCat.DropDownItems

            If itm.ToString = s_item Then
                index = smetaForm.menuStripCat.DropDownItems.IndexOf(itm)
                item = smetaForm.menuStripCat.DropDownItems.Item(index)
                Console.WriteLine(item)
            End If

        Next itm

        If item.DropDownItems.Count = 0 Then
            item.DropDownItems.Add(subItem)
        End If

        For Each itm As ToolStripMenuItem In item.DropDownItems

            If itm.ToString = s_subitem Then
                subItemExist = True
            End If

        Next itm

        If Not subItemExist Then
            item.DropDownItems.Add(subItem)
        End If

        subItemExist = False
        itemExist = False

    End Sub
End Module
