Public Class datasetForm
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        mainForm.btn_loadDB.PerformClick()
        createLightingDataset()
    End Sub
    Private Sub item_movHeads_Click(sender As Object, e As EventArgs) Handles item_movHeads.Click

        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)



    End Sub

    Private Sub item_strobes_Click(sender As Object, e As EventArgs) Handles item_strobes.Click
        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_blinders_Click(sender As Object, e As EventArgs) Handles item_blinders.Click
        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_arch_Click(sender As Object, e As EventArgs) Handles item_arch.Click
        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_LED_Click(sender As Object, e As EventArgs) Handles item_LED.Click
        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_smoke_Click(sender As Object, e As EventArgs) Handles item_smoke.Click
        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_consoles_Click(sender As Object, e As EventArgs) Handles item_consoles.Click
        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_intercom_Click(sender As Object, e As EventArgs) Handles item_intercom.Click
        If mainForm.lightDataset Is Nothing Then
            createLightingDataset()
        End If

        writeToLabel("Lighting", sender)

    End Sub

    Private Sub item_belimlight_Click(sender As Object, e As EventArgs) Handles item_belimlight.Click
        writeToLabelCompany(sender)
        Dim c As Color = Color.FromArgb(252, 228, 214)
        dgv_dataset.DataSource = mainForm.lightDataset.Tables(0)
        format_dgv_dataset(mainForm.tbl_Lighting_tables(0, 0).Name, c)

    End Sub

    Private Sub item_PRLighting_Click(sender As Object, e As EventArgs) Handles item_PRLighting.Click
        writeToLabelCompany(sender)
    End Sub

    Private Sub item_blackout_Click(sender As Object, e As EventArgs) Handles item_blackout.Click
        writeToLabelCompany(sender)
    End Sub

    Private Sub item_vision_Click(sender As Object, e As EventArgs) Handles item_vision.Click
        writeToLabelCompany(sender)
    End Sub

    Private Sub item_stage_Click(sender As Object, e As EventArgs) Handles item_stage.Click
        writeToLabelCompany(sender)
    End Sub

    '===================================================================================
    '             === MY FUNCTIONS ===
    '===================================================================================

    Sub writeToLabel(_department As String, _sender As Object)
        Me.GroupBox1.Visible = True
        Me.GroupBox2.Visible = True
        Me.lbl_dpartmentValue.Text = _department
        Me.lbl_subsectionValue.Text = _sender.text
    End Sub

    Sub writeToLabelCompany(_sender As Object)
        Me.GroupBox3.Visible = True
        Me.lbl_companyValue.Text = _sender.text
    End Sub

    '===================================================================================      
    '                === Format DataGridView ===
    '===================================================================================
    Sub format_dgv_dataset(_dtName As String, _color As Color)

        dgv_dataset.Columns(0).Width = 40                ' #
        dgv_dataset.Columns(1).Width = 175               ' Fixture
        dgv_dataset.Columns(2).Width = 40                ' Q-ty
        dgv_dataset.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv_dataset.Columns(3).Width = 220               ' BelImlight_1  (PRLightigTouring, BlackOut, Vision, Stage)
        dgv_dataset.Columns(4).Width = 40                ' Q-ty_1
        dgv_dataset.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv_dataset.Columns(5).Width = 220               ' BelImlight_2  (PRLightigTouring, BlackOut, Vision, Stage)
        dgv_dataset.Columns(6).Width = 40                ' Q-ty_2
        dgv_dataset.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv_dataset.Columns(7).Width = 180               ' BelImlight_3  (PRLightigTouring, BlackOut, Vision, Stage)
        dgv_dataset.Columns(8).Width = 40                ' Q-ty_3
        dgv_dataset.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For i = 0 To dgv_dataset.Rows.Count - 2

            'mainForm.DGV_in.Rows(i).Cells(1).Value = Date.FromOADate(mainForm.DGV_in.Rows(i).Cells(1).Value)
            dgv_dataset.RowsDefaultCellStyle.BackColor = _color
            dgv_dataset.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)

        Next i
    End Sub
End Class