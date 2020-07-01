Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Public Class mainForm

    Public sDir As String
    Public sFilePath As String
    Public sFileName As String


    Public fileNames As Collection
    Public filePath As Collection


    ' Dictionaries with Integer key
    Public i_wsDict As Dictionary(Of Integer, ExcelWorksheet)
    Public i_xlTableDict As Dictionary(Of Integer, ExcelTable)
    Public i_pivot_wsDict As Dictionary(Of Integer, Dictionary(Of Integer, ExcelWorksheet))
    Public i_pivotTableDict As Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable))

    '   Allow to get any excel table by number of iDepartment, iCategory, iCompany
    '   for Console.WriteLine(example i_superPivotDict(0)(0)(0).name) - movHeads_tbl
    Public i_superPivotDict As Dictionary(Of Integer, Dictionary(Of Integer, Dictionary(Of Integer, ExcelTable)))

    Public iDepartment, iCategory, iCompany As Integer

    Public dts, smeta_dts As DataSet
    Public selIndex As Integer      ' Selercted index

    Public sCompany() As String = {"belimlight", "PRLighting", "blackout", "vision", "stage"}
    Public sDepartment() As String = {"Lighting", "Screen", "Commutation", "Trusses and motors", "Construction", "Sound"}
    Public cancelFlag As Boolean = False

    Public smetaDepartment As New Collection

    '===================================================================================
    '             === mainForm_Load ===
    '===================================================================================
    Private Sub mainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        checkExpirationDate()
        smeta_dts = New DataSet
        chk_detail.Text = "Кратко"
        txt_qty.Text = String.Empty
    End Sub

    '===================================================================================
    '             === File => Load database ===
    '===================================================================================
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        load_db()

        iDepartment = 0
        iCategory = 0
        iCompany = 1

        menuItem_department.Enabled = True
        menuItem_company.Enabled = True

        cancelFlag = False
    End Sub


#Region "select Lighting"
    Private Sub item_movHeads_Click(sender As Object, e As EventArgs) Handles item_movHeads.Click

        iDepartment = 0
        iCategory = 0
        fillDGV("Lighting", sender)

    End Sub

    Private Sub item_strobes_Click(sender As Object, e As EventArgs) Handles item_strobes.Click

        iDepartment = 0
        iCategory = 1
        fillDGV("Lighting", sender)

    End Sub

    Private Sub item_blinders_Click(sender As Object, e As EventArgs) Handles item_blinders.Click

        iDepartment = 0
        iCategory = 2
        fillDGV("Lighting", sender)

    End Sub

    Private Sub item_arch_Click(sender As Object, e As EventArgs) Handles item_arch.Click

        iDepartment = 0
        iCategory = 3
        fillDGV("Lighting", sender)

    End Sub

    Private Sub item_LED_Click(sender As Object, e As EventArgs) Handles item_LED.Click

        iDepartment = 0
        iCategory = 4
        fillDGV("Lighting", sender)

    End Sub

    Private Sub item_smoke_Click(sender As Object, e As EventArgs) Handles item_smoke.Click

        iDepartment = 0
        iCategory = 5
        fillDGV("Lighting", sender)

    End Sub

    Private Sub item_consoles_Click(sender As Object, e As EventArgs) Handles item_consoles.Click

        iDepartment = 0
        iCategory = 6
        fillDGV("Lighting", sender)

    End Sub

    Private Sub item_intercom_Click(sender As Object, e As EventArgs) Handles item_intercom.Click

        iDepartment = 0
        iCategory = 7
        fillDGV("Lighting", sender)

    End Sub
#End Region

#Region "select Screen"
    Private Sub item_modules_Click(sender As Object, e As EventArgs) Handles item_modules.Click
        iDepartment = 1
        iCategory = 0
        fillDGV("Screen", sender)
    End Sub

    Private Sub item_servers_Click(sender As Object, e As EventArgs) Handles item_servers.Click
        iDepartment = 1
        iCategory = 1
        fillDGV("Screen", sender)
    End Sub

    Private Sub item_controllers1_Click(sender As Object, e As EventArgs) Handles item_controllers1.Click
        iDepartment = 1
        iCategory = 2
        fillDGV("Screen", sender)
    End Sub

    Private Sub item_controllers2_Click(sender As Object, e As EventArgs) Handles item_controllers2.Click
        iDepartment = 1
        iCategory = 3
        fillDGV("Screen", sender)
    End Sub

    Private Sub item_projectors_Click(sender As Object, e As EventArgs) Handles item_projectors.Click
        iDepartment = 1
        iCategory = 4
        fillDGV("Screen", sender)
    End Sub

    Private Sub item_scr_construction_Click(sender As Object, e As EventArgs) Handles item_scr_construction.Click
        iDepartment = 1
        iCategory = 5
        fillDGV("Screen", sender)
    End Sub

    Private Sub item_lightDesks_Click(sender As Object, e As EventArgs) Handles item_lightDesks.Click
        iDepartment = 1
        iCategory = 6
        fillDGV("Screen", sender)
    End Sub

    Private Sub item_cameras_Click(sender As Object, e As EventArgs) Handles item_cameras.Click
        iDepartment = 1
        iCategory = 7
        fillDGV("Screen", sender)
    End Sub
#End Region

#Region "Select Commutation"
    Private Sub item_pwrdistr_Click(sender As Object, e As EventArgs) Handles item_pwrdistr.Click
        iDepartment = 2
        iCategory = 0
        fillDGV("Commutation", sender)
    End Sub

    Private Sub item_comm_Click(sender As Object, e As EventArgs) Handles item_comm.Click
        iDepartment = 2
        iCategory = 1
        fillDGV("Commutation", sender)
    End Sub

    Private Sub item_pwrcomm_Click(sender As Object, e As EventArgs) Handles item_pwrcomm.Click
        iDepartment = 2
        iCategory = 2
        fillDGV("Commutation", sender)
    End Sub

    Private Sub item_rest_Click(sender As Object, e As EventArgs) Handles item_rest.Click
        iDepartment = 2
        iCategory = 3
        fillDGV("Commutation", sender)
    End Sub
#End Region

#Region "Select Truss and motors"
    Private Sub item_truss30x30_Click(sender As Object, e As EventArgs) Handles item_truss30x30.Click
        iDepartment = 3
        iCategory = 0
        fillDGV("Trusses and motors", sender)
    End Sub

    Private Sub item_truss40x40_Click(sender As Object, e As EventArgs) Handles item_truss40x40.Click
        iDepartment = 3
        iCategory = 1
        fillDGV("Trusses and motors", sender)
    End Sub

    Private Sub item_truss50x60_Click(sender As Object, e As EventArgs) Handles item_truss50x60.Click
        iDepartment = 3
        iCategory = 2
        fillDGV("Trusses and motors", sender)
    End Sub

    Private Sub item_motors_Click(sender As Object, e As EventArgs) Handles item_motors.Click
        iDepartment = 3
        iCategory = 3
        fillDGV("Trusses and motors", sender)
    End Sub

    Private Sub item_rigging_Click(sender As Object, e As EventArgs) Handles item_rigging.Click
        iDepartment = 3
        iCategory = 4
        fillDGV("Trusses and motors", sender)
    End Sub

    Private Sub item_diff_Click(sender As Object, e As EventArgs) Handles item_diff.Click
        iDepartment = 3
        iCategory = 5
        fillDGV("Trusses and motors", sender)
    End Sub

    Private Sub item_completeConstr_Click(sender As Object, e As EventArgs) Handles item_completeConstr.Click
        iDepartment = 3
        iCategory = 6
        fillDGV("Trusses and motors", sender)
    End Sub

    Private Sub item_stagelifts_Click(sender As Object, e As EventArgs) Handles item_stagelifts.Click
        iDepartment = 3
        iCategory = 7
        fillDGV("Trusses and motors", sender)
    End Sub
#End Region

#Region "Select Construction"
    Private Sub item_stageModules_Click(sender As Object, e As EventArgs) Handles item_stageModules.Click
        iDepartment = 4
        iCategory = 0
        fillDGV("Construction", sender)
    End Sub

    Private Sub item_scaffold_J001_Click(sender As Object, e As EventArgs) Handles item_scaffold_J001.Click
        iDepartment = 4
        iCategory = 1
        fillDGV("Construction", sender)
    End Sub

    Private Sub item_scaffold_J004_Click(sender As Object, e As EventArgs) Handles item_scaffold_J004.Click
        iDepartment = 4
        iCategory = 2
        fillDGV("Construction", sender)
    End Sub

    Private Sub item_scaffold_steps_Click(sender As Object, e As EventArgs) Handles item_scaffold_steps.Click
        iDepartment = 4
        iCategory = 3
        fillDGV("Construction", sender)
    End Sub

    Private Sub item_barricades_Click(sender As Object, e As EventArgs) Handles item_barricades.Click
        iDepartment = 4
        iCategory = 4
        fillDGV("Construction", sender)
    End Sub

    Private Sub item_details_Click(sender As Object, e As EventArgs) Handles item_details.Click
        iDepartment = 4
        iCategory = 5
        fillDGV("Construction", sender)
    End Sub
#End Region

#Region "Select Sound"
    Private Sub item_speakers_Click(sender As Object, e As EventArgs) Handles item_speakers.Click
        iDepartment = 5
        iCategory = 0
        fillDGV("Sound", sender)
    End Sub

    Private Sub item_ampracks_Click(sender As Object, e As EventArgs) Handles item_ampracks.Click
        iDepartment = 5
        iCategory = 1
        fillDGV("Sound", sender)
    End Sub

    Private Sub item_monitors_Click(sender As Object, e As EventArgs) Handles item_monitors.Click
        iDepartment = 5
        iCategory = 2
        fillDGV("Sound", sender)
    End Sub

    Private Sub item_mixdesks_Click(sender As Object, e As EventArgs) Handles item_mixdesks.Click
        iDepartment = 5
        iCategory = 3
        fillDGV("Sound", sender)
    End Sub

    Private Sub item_dj_stage_Click(sender As Object, e As EventArgs) Handles item_dj_stage.Click
        iDepartment = 5
        iCategory = 4
        fillDGV("Sound", sender)
    End Sub
#End Region

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub


    Private Sub dgv_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellClick
        dgv_clickCell(sender, e)
        txt_qty.Text = ""
        txt_qty.BackColor = Color.Yellow
        txt_qty.Select()
    End Sub
    Private Sub txt_qty_TextChanged(sender As Object, e As EventArgs) Handles txt_qty.TextChanged
        If txt_qty.Text <> "" Then
            txt_qty.BackColor = Color.White
        End If
    End Sub

    '===================================================================================
    '             === Add tables to Smeta dataTable ===
    '===================================================================================
    Private Sub btn_add_to_smeta_Click(sender As Object, e As EventArgs) Handles btn_add_to_smeta.Click

        If txt_qty.Text = "" Then
            MsgBox("Требуется заполнить количество приборов!")
            Exit Sub
        End If
        createSmeta_dt()

        'Console.Write(smeta_dts.Tables.Count & vbTab)
        'Console.WriteLine(selIndex)

        smetaForm.Show()

    End Sub
    Private Sub chk_detail_CheckStateChanged(sender As Object, e As EventArgs) Handles chk_detail.CheckStateChanged
        If chk_detail.Checked = True Then
            chk_detail.Text = "Подробно"
        Else
            chk_detail.Text = "Кратко"
        End If
        format_DGV(DGV)
        If smetaForm.DGV_smeta.Rows.Count > 0 Then
            format_DGV(smetaForm.DGV_smeta)
        End If
    End Sub

End Class
