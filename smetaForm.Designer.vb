<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class smetaForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DGV_smeta = New System.Windows.Forms.DataGridView()
        Me.menu_smeta = New System.Windows.Forms.MenuStrip()
        Me.menuStripCat = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmb_item = New System.Windows.Forms.ComboBox()
        Me.txt_subitem = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.DGV_smeta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menu_smeta.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(363, 580)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DGV_smeta
        '
        Me.DGV_smeta.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_smeta.Dock = System.Windows.Forms.DockStyle.Top
        Me.DGV_smeta.Location = New System.Drawing.Point(0, 24)
        Me.DGV_smeta.Name = "DGV_smeta"
        Me.DGV_smeta.Size = New System.Drawing.Size(800, 489)
        Me.DGV_smeta.TabIndex = 2
        '
        'menu_smeta
        '
        Me.menu_smeta.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuStripCat})
        Me.menu_smeta.Location = New System.Drawing.Point(0, 0)
        Me.menu_smeta.Name = "menu_smeta"
        Me.menu_smeta.Size = New System.Drawing.Size(800, 24)
        Me.menu_smeta.TabIndex = 3
        Me.menu_smeta.Text = "MenuStrip1"
        '
        'menuStripCat
        '
        Me.menuStripCat.Name = "menuStripCat"
        Me.menuStripCat.Size = New System.Drawing.Size(75, 20)
        Me.menuStripCat.Text = "Categories"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(444, 580)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'cmb_item
        '
        Me.cmb_item.FormattingEnabled = True
        Me.cmb_item.Items.AddRange(New Object() {"Item1", "Item2", "Item3", "Item4"})
        Me.cmb_item.Location = New System.Drawing.Point(118, 561)
        Me.cmb_item.Name = "cmb_item"
        Me.cmb_item.Size = New System.Drawing.Size(121, 21)
        Me.cmb_item.TabIndex = 4
        '
        'txt_subitem
        '
        Me.txt_subitem.Location = New System.Drawing.Point(592, 561)
        Me.txt_subitem.Name = "txt_subitem"
        Me.txt_subitem.Size = New System.Drawing.Size(100, 20)
        Me.txt_subitem.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(118, 545)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Items"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(589, 545)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Subitems"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(282, 530)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'smetaForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 615)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_subitem)
        Me.Controls.Add(Me.cmb_item)
        Me.Controls.Add(Me.DGV_smeta)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.menu_smeta)
        Me.MainMenuStrip = Me.menu_smeta
        Me.Name = "smetaForm"
        Me.Text = "smetaForm"
        CType(Me.DGV_smeta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menu_smeta.ResumeLayout(False)
        Me.menu_smeta.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents DGV_smeta As DataGridView
    Friend WithEvents menu_smeta As MenuStrip
    Friend WithEvents menuStripCat As ToolStripMenuItem
    Friend WithEvents Button2 As Button
    Friend WithEvents cmb_item As ComboBox
    Friend WithEvents txt_subitem As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button3 As Button
End Class
