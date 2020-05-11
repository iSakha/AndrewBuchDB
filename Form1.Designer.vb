<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tabControl = New System.Windows.Forms.TabControl()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.btn_loadDB = New System.Windows.Forms.Button()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lbl_qtyTotal = New System.Windows.Forms.Label()
        Me.lbl_smeta_qty = New System.Windows.Forms.Label()
        Me.lbl_qty_vision = New System.Windows.Forms.Label()
        Me.lbl_qty_blackout = New System.Windows.Forms.Label()
        Me.lbl_qty_PRLighting = New System.Windows.Forms.Label()
        Me.lbl_qty_belimlight = New System.Windows.Forms.Label()
        Me.txt_qty3 = New System.Windows.Forms.TextBox()
        Me.txt_qty2 = New System.Windows.Forms.TextBox()
        Me.txt_qty1 = New System.Windows.Forms.TextBox()
        Me.txt_qty = New System.Windows.Forms.TextBox()
        Me.rtb_fixtureName = New System.Windows.Forms.RichTextBox()
        Me.rtb_ThirdName = New System.Windows.Forms.RichTextBox()
        Me.rtb_SecondName = New System.Windows.Forms.RichTextBox()
        Me.rtb_FirstName = New System.Windows.Forms.RichTextBox()
        Me.DGV_light = New System.Windows.Forms.DataGridView()
        Me.btn_vision = New System.Windows.Forms.Button()
        Me.btn_blackOut = New System.Windows.Forms.Button()
        Me.btn_prLight = New System.Windows.Forms.Button()
        Me.btn_belIm = New System.Windows.Forms.Button()
        Me.lbl_category = New System.Windows.Forms.Label()
        Me.cmb_category = New System.Windows.Forms.ComboBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.TabPage7 = New System.Windows.Forms.TabPage()
        Me.btn_save = New System.Windows.Forms.Button()
        Me.btn_update = New System.Windows.Forms.Button()
        Me.btn_del = New System.Windows.Forms.Button()
        Me.btn_add = New System.Windows.Forms.Button()
        Me.btn_next = New System.Windows.Forms.Button()
        Me.btn_prev = New System.Windows.Forms.Button()
        Me.OFD = New System.Windows.Forms.OpenFileDialog()
        Me.tabControl.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DGV_light, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tabControl
        '
        Me.tabControl.Controls.Add(Me.TabPage3)
        Me.tabControl.Controls.Add(Me.TabPage1)
        Me.tabControl.Controls.Add(Me.TabPage2)
        Me.tabControl.Controls.Add(Me.TabPage4)
        Me.tabControl.Controls.Add(Me.TabPage5)
        Me.tabControl.Controls.Add(Me.TabPage6)
        Me.tabControl.Controls.Add(Me.TabPage7)
        Me.tabControl.Location = New System.Drawing.Point(0, 0)
        Me.tabControl.Name = "tabControl"
        Me.tabControl.SelectedIndex = 0
        Me.tabControl.Size = New System.Drawing.Size(1067, 659)
        Me.tabControl.TabIndex = 0
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btn_loadDB)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(1059, 633)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Menu"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btn_loadDB
        '
        Me.btn_loadDB.Location = New System.Drawing.Point(72, 114)
        Me.btn_loadDB.Name = "btn_loadDB"
        Me.btn_loadDB.Size = New System.Drawing.Size(143, 63)
        Me.btn_loadDB.TabIndex = 0
        Me.btn_loadDB.Text = "Load DB"
        Me.btn_loadDB.UseVisualStyleBackColor = True
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lbl_qtyTotal)
        Me.TabPage1.Controls.Add(Me.lbl_smeta_qty)
        Me.TabPage1.Controls.Add(Me.lbl_qty_vision)
        Me.TabPage1.Controls.Add(Me.lbl_qty_blackout)
        Me.TabPage1.Controls.Add(Me.lbl_qty_PRLighting)
        Me.TabPage1.Controls.Add(Me.lbl_qty_belimlight)
        Me.TabPage1.Controls.Add(Me.txt_qty3)
        Me.TabPage1.Controls.Add(Me.txt_qty2)
        Me.TabPage1.Controls.Add(Me.txt_qty1)
        Me.TabPage1.Controls.Add(Me.txt_qty)
        Me.TabPage1.Controls.Add(Me.rtb_fixtureName)
        Me.TabPage1.Controls.Add(Me.rtb_ThirdName)
        Me.TabPage1.Controls.Add(Me.rtb_SecondName)
        Me.TabPage1.Controls.Add(Me.rtb_FirstName)
        Me.TabPage1.Controls.Add(Me.DGV_light)
        Me.TabPage1.Controls.Add(Me.btn_vision)
        Me.TabPage1.Controls.Add(Me.btn_blackOut)
        Me.TabPage1.Controls.Add(Me.btn_prLight)
        Me.TabPage1.Controls.Add(Me.btn_belIm)
        Me.TabPage1.Controls.Add(Me.lbl_category)
        Me.TabPage1.Controls.Add(Me.cmb_category)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1059, 633)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Свет"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lbl_qtyTotal
        '
        Me.lbl_qtyTotal.AutoSize = True
        Me.lbl_qtyTotal.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lbl_qtyTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_qtyTotal.ForeColor = System.Drawing.Color.DarkBlue
        Me.lbl_qtyTotal.Location = New System.Drawing.Point(226, 68)
        Me.lbl_qtyTotal.Name = "lbl_qtyTotal"
        Me.lbl_qtyTotal.Size = New System.Drawing.Size(0, 25)
        Me.lbl_qtyTotal.TabIndex = 16
        '
        'lbl_smeta_qty
        '
        Me.lbl_smeta_qty.AutoSize = True
        Me.lbl_smeta_qty.BackColor = System.Drawing.Color.SeaShell
        Me.lbl_smeta_qty.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_smeta_qty.ForeColor = System.Drawing.Color.DarkBlue
        Me.lbl_smeta_qty.Location = New System.Drawing.Point(36, 68)
        Me.lbl_smeta_qty.Name = "lbl_smeta_qty"
        Me.lbl_smeta_qty.Size = New System.Drawing.Size(167, 25)
        Me.lbl_smeta_qty.TabIndex = 15
        Me.lbl_smeta_qty.Text = "Всего из сметы"
        '
        'lbl_qty_vision
        '
        Me.lbl_qty_vision.AutoSize = True
        Me.lbl_qty_vision.BackColor = System.Drawing.Color.SeaShell
        Me.lbl_qty_vision.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_qty_vision.ForeColor = System.Drawing.Color.DarkBlue
        Me.lbl_qty_vision.Location = New System.Drawing.Point(810, 68)
        Me.lbl_qty_vision.Name = "lbl_qty_vision"
        Me.lbl_qty_vision.Size = New System.Drawing.Size(0, 25)
        Me.lbl_qty_vision.TabIndex = 14
        '
        'lbl_qty_blackout
        '
        Me.lbl_qty_blackout.AutoSize = True
        Me.lbl_qty_blackout.BackColor = System.Drawing.Color.SeaShell
        Me.lbl_qty_blackout.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_qty_blackout.ForeColor = System.Drawing.Color.DarkBlue
        Me.lbl_qty_blackout.Location = New System.Drawing.Point(668, 68)
        Me.lbl_qty_blackout.Name = "lbl_qty_blackout"
        Me.lbl_qty_blackout.Size = New System.Drawing.Size(0, 25)
        Me.lbl_qty_blackout.TabIndex = 14
        '
        'lbl_qty_PRLighting
        '
        Me.lbl_qty_PRLighting.AutoSize = True
        Me.lbl_qty_PRLighting.BackColor = System.Drawing.Color.SeaShell
        Me.lbl_qty_PRLighting.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_qty_PRLighting.ForeColor = System.Drawing.Color.DarkBlue
        Me.lbl_qty_PRLighting.Location = New System.Drawing.Point(528, 68)
        Me.lbl_qty_PRLighting.Name = "lbl_qty_PRLighting"
        Me.lbl_qty_PRLighting.Size = New System.Drawing.Size(0, 25)
        Me.lbl_qty_PRLighting.TabIndex = 14
        '
        'lbl_qty_belimlight
        '
        Me.lbl_qty_belimlight.AutoSize = True
        Me.lbl_qty_belimlight.BackColor = System.Drawing.Color.SeaShell
        Me.lbl_qty_belimlight.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_qty_belimlight.ForeColor = System.Drawing.Color.DarkBlue
        Me.lbl_qty_belimlight.Location = New System.Drawing.Point(377, 68)
        Me.lbl_qty_belimlight.Name = "lbl_qty_belimlight"
        Me.lbl_qty_belimlight.Size = New System.Drawing.Size(0, 25)
        Me.lbl_qty_belimlight.TabIndex = 14
        '
        'txt_qty3
        '
        Me.txt_qty3.Location = New System.Drawing.Point(1002, 546)
        Me.txt_qty3.Name = "txt_qty3"
        Me.txt_qty3.Size = New System.Drawing.Size(55, 20)
        Me.txt_qty3.TabIndex = 13
        '
        'txt_qty2
        '
        Me.txt_qty2.Location = New System.Drawing.Point(739, 546)
        Me.txt_qty2.Name = "txt_qty2"
        Me.txt_qty2.Size = New System.Drawing.Size(55, 20)
        Me.txt_qty2.TabIndex = 12
        '
        'txt_qty1
        '
        Me.txt_qty1.Location = New System.Drawing.Point(473, 546)
        Me.txt_qty1.Name = "txt_qty1"
        Me.txt_qty1.Size = New System.Drawing.Size(55, 20)
        Me.txt_qty1.TabIndex = 11
        '
        'txt_qty
        '
        Me.txt_qty.Location = New System.Drawing.Point(208, 546)
        Me.txt_qty.Name = "txt_qty"
        Me.txt_qty.Size = New System.Drawing.Size(55, 20)
        Me.txt_qty.TabIndex = 10
        '
        'rtb_fixtureName
        '
        Me.rtb_fixtureName.Location = New System.Drawing.Point(4, 526)
        Me.rtb_fixtureName.Name = "rtb_fixtureName"
        Me.rtb_fixtureName.Size = New System.Drawing.Size(199, 65)
        Me.rtb_fixtureName.TabIndex = 9
        Me.rtb_fixtureName.Text = ""
        '
        'rtb_ThirdName
        '
        Me.rtb_ThirdName.Location = New System.Drawing.Point(799, 526)
        Me.rtb_ThirdName.Name = "rtb_ThirdName"
        Me.rtb_ThirdName.Size = New System.Drawing.Size(199, 65)
        Me.rtb_ThirdName.TabIndex = 7
        Me.rtb_ThirdName.Text = ""
        '
        'rtb_SecondName
        '
        Me.rtb_SecondName.Location = New System.Drawing.Point(534, 526)
        Me.rtb_SecondName.Name = "rtb_SecondName"
        Me.rtb_SecondName.Size = New System.Drawing.Size(199, 65)
        Me.rtb_SecondName.TabIndex = 6
        Me.rtb_SecondName.Text = ""
        '
        'rtb_FirstName
        '
        Me.rtb_FirstName.Location = New System.Drawing.Point(269, 526)
        Me.rtb_FirstName.Name = "rtb_FirstName"
        Me.rtb_FirstName.Size = New System.Drawing.Size(199, 65)
        Me.rtb_FirstName.TabIndex = 5
        Me.rtb_FirstName.Text = ""
        '
        'DGV_light
        '
        Me.DGV_light.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_light.Location = New System.Drawing.Point(0, 110)
        Me.DGV_light.Name = "DGV_light"
        Me.DGV_light.Size = New System.Drawing.Size(1056, 394)
        Me.DGV_light.TabIndex = 3
        '
        'btn_vision
        '
        Me.btn_vision.BackColor = System.Drawing.Color.FromArgb(CType(CType(226, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.btn_vision.Location = New System.Drawing.Point(766, 18)
        Me.btn_vision.Name = "btn_vision"
        Me.btn_vision.Size = New System.Drawing.Size(134, 33)
        Me.btn_vision.TabIndex = 2
        Me.btn_vision.Text = "Multivision"
        Me.btn_vision.UseVisualStyleBackColor = False
        '
        'btn_blackOut
        '
        Me.btn_blackOut.BackColor = System.Drawing.Color.FromArgb(CType(CType(237, Byte), Integer), CType(CType(237, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.btn_blackOut.Location = New System.Drawing.Point(623, 18)
        Me.btn_blackOut.Name = "btn_blackOut"
        Me.btn_blackOut.Size = New System.Drawing.Size(134, 33)
        Me.btn_blackOut.TabIndex = 2
        Me.btn_blackOut.Text = "Blackout"
        Me.btn_blackOut.UseVisualStyleBackColor = False
        '
        'btn_prLight
        '
        Me.btn_prLight.BackColor = System.Drawing.Color.FromArgb(CType(CType(221, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.btn_prLight.Location = New System.Drawing.Point(480, 18)
        Me.btn_prLight.Name = "btn_prLight"
        Me.btn_prLight.Size = New System.Drawing.Size(134, 33)
        Me.btn_prLight.TabIndex = 2
        Me.btn_prLight.Text = "PRLighting"
        Me.btn_prLight.UseVisualStyleBackColor = False
        '
        'btn_belIm
        '
        Me.btn_belIm.BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(228, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.btn_belIm.Location = New System.Drawing.Point(337, 18)
        Me.btn_belIm.Name = "btn_belIm"
        Me.btn_belIm.Size = New System.Drawing.Size(134, 33)
        Me.btn_belIm.TabIndex = 2
        Me.btn_belIm.Text = "Belimlight"
        Me.btn_belIm.UseVisualStyleBackColor = False
        '
        'lbl_category
        '
        Me.lbl_category.AutoSize = True
        Me.lbl_category.Location = New System.Drawing.Point(24, 4)
        Me.lbl_category.Name = "lbl_category"
        Me.lbl_category.Size = New System.Drawing.Size(60, 13)
        Me.lbl_category.TabIndex = 1
        Me.lbl_category.Text = "Категория"
        '
        'cmb_category
        '
        Me.cmb_category.FormattingEnabled = True
        Me.cmb_category.Items.AddRange(New Object() {"Головы/Moving heads", "Стробоскопы/strobes, Прожектора следящего света/followspots", "Пары, блайндера/PAR's, blinders", "Архитектурный свет/Architecture fixtures", "Светодиодные приборы/LED fixtures", "Дым, туман, вентиляторы, прочее/Fog, Haze, fans, rest", "Пульты/lighting desks", "Системы связи/Intercoms and radios"})
        Me.cmb_category.Location = New System.Drawing.Point(24, 30)
        Me.cmb_category.Name = "cmb_category"
        Me.cmb_category.Size = New System.Drawing.Size(282, 21)
        Me.cmb_category.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1059, 633)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Экран"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(1059, 633)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Коммутация"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(1059, 633)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Фермы,моторы"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'TabPage6
        '
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(1059, 633)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "Конструктив"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'TabPage7
        '
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Size = New System.Drawing.Size(1059, 633)
        Me.TabPage7.TabIndex = 6
        Me.TabPage7.Text = "Звук"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'btn_save
        '
        Me.btn_save.Location = New System.Drawing.Point(902, 688)
        Me.btn_save.Name = "btn_save"
        Me.btn_save.Size = New System.Drawing.Size(75, 23)
        Me.btn_save.TabIndex = 14
        Me.btn_save.Text = "Save"
        Me.btn_save.UseVisualStyleBackColor = True
        '
        'btn_update
        '
        Me.btn_update.Location = New System.Drawing.Point(820, 688)
        Me.btn_update.Name = "btn_update"
        Me.btn_update.Size = New System.Drawing.Size(75, 23)
        Me.btn_update.TabIndex = 14
        Me.btn_update.Text = "Update"
        Me.btn_update.UseVisualStyleBackColor = True
        '
        'btn_del
        '
        Me.btn_del.Location = New System.Drawing.Point(738, 688)
        Me.btn_del.Name = "btn_del"
        Me.btn_del.Size = New System.Drawing.Size(75, 23)
        Me.btn_del.TabIndex = 14
        Me.btn_del.Text = "Delete"
        Me.btn_del.UseVisualStyleBackColor = True
        '
        'btn_add
        '
        Me.btn_add.Location = New System.Drawing.Point(656, 688)
        Me.btn_add.Name = "btn_add"
        Me.btn_add.Size = New System.Drawing.Size(75, 23)
        Me.btn_add.TabIndex = 14
        Me.btn_add.Text = "Add"
        Me.btn_add.UseVisualStyleBackColor = True
        '
        'btn_next
        '
        Me.btn_next.Location = New System.Drawing.Point(113, 688)
        Me.btn_next.Name = "btn_next"
        Me.btn_next.Size = New System.Drawing.Size(75, 23)
        Me.btn_next.TabIndex = 14
        Me.btn_next.Text = ">>>"
        Me.btn_next.UseVisualStyleBackColor = True
        '
        'btn_prev
        '
        Me.btn_prev.Location = New System.Drawing.Point(32, 688)
        Me.btn_prev.Name = "btn_prev"
        Me.btn_prev.Size = New System.Drawing.Size(75, 23)
        Me.btn_prev.TabIndex = 14
        Me.btn_prev.Text = "<<<"
        Me.btn_prev.UseVisualStyleBackColor = True
        '
        'OFD
        '
        Me.OFD.FileName = "OpenFileDialog1"
        '
        'mainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1068, 736)
        Me.Controls.Add(Me.btn_save)
        Me.Controls.Add(Me.btn_update)
        Me.Controls.Add(Me.tabControl)
        Me.Controls.Add(Me.btn_del)
        Me.Controls.Add(Me.btn_prev)
        Me.Controls.Add(Me.btn_add)
        Me.Controls.Add(Me.btn_next)
        Me.Name = "mainForm"
        Me.Text = "mainForm"
        Me.tabControl.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.DGV_light, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tabControl As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents lbl_category As Label
    Friend WithEvents cmb_category As ComboBox
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents DGV_light As DataGridView
    Friend WithEvents btn_vision As Button
    Friend WithEvents btn_blackOut As Button
    Friend WithEvents btn_prLight As Button
    Friend WithEvents btn_belIm As Button
    Friend WithEvents rtb_SecondName As RichTextBox
    Friend WithEvents rtb_FirstName As RichTextBox
    Friend WithEvents rtb_ThirdName As RichTextBox
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents btn_loadDB As Button
    Friend WithEvents OFD As OpenFileDialog
    Friend WithEvents rtb_fixtureName As RichTextBox
    Friend WithEvents txt_qty3 As TextBox
    Friend WithEvents txt_qty2 As TextBox
    Friend WithEvents txt_qty1 As TextBox
    Friend WithEvents txt_qty As TextBox
    Friend WithEvents btn_save As Button
    Friend WithEvents btn_update As Button
    Friend WithEvents btn_del As Button
    Friend WithEvents btn_add As Button
    Friend WithEvents btn_next As Button
    Friend WithEvents btn_prev As Button
    Friend WithEvents lbl_qty_vision As Label
    Friend WithEvents lbl_qty_blackout As Label
    Friend WithEvents lbl_qty_PRLighting As Label
    Friend WithEvents lbl_qty_belimlight As Label
    Friend WithEvents lbl_qtyTotal As Label
    Friend WithEvents lbl_smeta_qty As Label
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents TabPage6 As TabPage
    Friend WithEvents TabPage7 As TabPage
End Class
