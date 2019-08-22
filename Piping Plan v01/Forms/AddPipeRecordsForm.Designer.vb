<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AddPipeRecordsForm
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
        Me.btnSavePipe = New System.Windows.Forms.Button()
        Me.ckbPipeSkontrolovane = New System.Windows.Forms.CheckBox()
        Me.cboPipeOhybRovna = New System.Windows.Forms.ComboBox()
        Me.cboPipeHrubka = New System.Windows.Forms.ComboBox()
        Me.cboPipePriemer = New System.Windows.Forms.ComboBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtPipePN = New System.Windows.Forms.TextBox()
        Me.txtPipeDlzka = New System.Windows.Forms.TextBox()
        Me.txtPipeSubPN = New System.Windows.Forms.TextBox()
        Me.gridAddPipeRecord = New System.Windows.Forms.DataGridView()
        Me.btnDelPipe = New System.Windows.Forms.Button()
        Me.txtPipeID = New System.Windows.Forms.TextBox()
        Me.btnEditPipe = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.gridAddPipeRecord, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnSavePipe
        '
        Me.btnSavePipe.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSavePipe.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSavePipe.Font = New System.Drawing.Font("Century Gothic", 11.0!)
        Me.btnSavePipe.ForeColor = System.Drawing.Color.Black
        Me.btnSavePipe.Location = New System.Drawing.Point(854, 15)
        Me.btnSavePipe.Margin = New System.Windows.Forms.Padding(0)
        Me.btnSavePipe.Name = "btnSavePipe"
        Me.btnSavePipe.Size = New System.Drawing.Size(120, 51)
        Me.btnSavePipe.TabIndex = 7
        Me.btnSavePipe.Text = "Uložiť novú"
        Me.btnSavePipe.UseVisualStyleBackColor = False
        '
        'ckbPipeSkontrolovane
        '
        Me.ckbPipeSkontrolovane.AutoSize = True
        Me.ckbPipeSkontrolovane.Enabled = False
        Me.ckbPipeSkontrolovane.Font = New System.Drawing.Font("Century Gothic", 12.0!)
        Me.ckbPipeSkontrolovane.Location = New System.Drawing.Point(614, 44)
        Me.ckbPipeSkontrolovane.Name = "ckbPipeSkontrolovane"
        Me.ckbPipeSkontrolovane.Size = New System.Drawing.Size(141, 25)
        Me.ckbPipeSkontrolovane.TabIndex = 6
        Me.ckbPipeSkontrolovane.Text = "Skontrolované"
        Me.ckbPipeSkontrolovane.UseVisualStyleBackColor = True
        '
        'cboPipeOhybRovna
        '
        Me.cboPipeOhybRovna.Font = New System.Drawing.Font("Century Gothic", 14.0!)
        Me.cboPipeOhybRovna.FormattingEnabled = True
        Me.cboPipeOhybRovna.Items.AddRange(New Object() {"Ohyb", "Rovna"})
        Me.cboPipeOhybRovna.Location = New System.Drawing.Point(479, 36)
        Me.cboPipeOhybRovna.Name = "cboPipeOhybRovna"
        Me.cboPipeOhybRovna.Size = New System.Drawing.Size(129, 30)
        Me.cboPipeOhybRovna.TabIndex = 5
        '
        'cboPipeHrubka
        '
        Me.cboPipeHrubka.Font = New System.Drawing.Font("Century Gothic", 14.0!)
        Me.cboPipeHrubka.FormattingEnabled = True
        Me.cboPipeHrubka.Items.AddRange(New Object() {"1", "1,5"})
        Me.cboPipeHrubka.Location = New System.Drawing.Point(326, 36)
        Me.cboPipeHrubka.Name = "cboPipeHrubka"
        Me.cboPipeHrubka.Size = New System.Drawing.Size(61, 30)
        Me.cboPipeHrubka.TabIndex = 3
        '
        'cboPipePriemer
        '
        Me.cboPipePriemer.Font = New System.Drawing.Font("Century Gothic", 14.0!)
        Me.cboPipePriemer.FormattingEnabled = True
        Me.cboPipePriemer.Items.AddRange(New Object() {"12", "16", "18", "22", "28", "35", "42", "54", "64"})
        Me.cboPipePriemer.Location = New System.Drawing.Point(259, 36)
        Me.cboPipePriemer.Name = "cboPipePriemer"
        Me.cboPipePriemer.Size = New System.Drawing.Size(61, 30)
        Me.cboPipePriemer.TabIndex = 2
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label18.Location = New System.Drawing.Point(478, 18)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(111, 17)
        Me.Label18.TabIndex = 111
        Me.Label18.Text = "Ohýbaná / Rovná"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label17.Location = New System.Drawing.Point(393, 18)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(39, 17)
        Me.Label17.TabIndex = 110
        Me.Label17.Text = "Dĺžka"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label16.Location = New System.Drawing.Point(325, 18)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 17)
        Me.Label16.TabIndex = 109
        Me.Label16.Text = "Hrúbka"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label15.Location = New System.Drawing.Point(257, 18)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(53, 17)
        Me.Label15.TabIndex = 108
        Me.Label15.Text = "Priemer"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label13.Location = New System.Drawing.Point(180, 17)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(73, 17)
        Me.Label13.TabIndex = 107
        Me.Label13.Text = "Pod-Trubka"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label14.Location = New System.Drawing.Point(7, 17)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 17)
        Me.Label14.TabIndex = 114
        Me.Label14.Text = "Piping (PN)"
        '
        'txtPipePN
        '
        Me.txtPipePN.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPipePN.Font = New System.Drawing.Font("Century Gothic", 14.0!)
        Me.txtPipePN.ForeColor = System.Drawing.Color.FromArgb(CType(CType(76, Byte), Integer), CType(CType(77, Byte), Integer), CType(CType(80, Byte), Integer))
        Me.txtPipePN.Location = New System.Drawing.Point(8, 35)
        Me.txtPipePN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtPipePN.MaxLength = 23
        Me.txtPipePN.Name = "txtPipePN"
        Me.txtPipePN.Size = New System.Drawing.Size(167, 30)
        Me.txtPipePN.TabIndex = 0
        Me.txtPipePN.WordWrap = False
        '
        'txtPipeDlzka
        '
        Me.txtPipeDlzka.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPipeDlzka.Font = New System.Drawing.Font("Century Gothic", 14.0!)
        Me.txtPipeDlzka.ForeColor = System.Drawing.Color.FromArgb(CType(CType(76, Byte), Integer), CType(CType(77, Byte), Integer), CType(CType(80, Byte), Integer))
        Me.txtPipeDlzka.Location = New System.Drawing.Point(394, 36)
        Me.txtPipeDlzka.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtPipeDlzka.MaxLength = 23
        Me.txtPipeDlzka.Name = "txtPipeDlzka"
        Me.txtPipeDlzka.Size = New System.Drawing.Size(79, 30)
        Me.txtPipeDlzka.TabIndex = 4
        Me.txtPipeDlzka.WordWrap = False
        '
        'txtPipeSubPN
        '
        Me.txtPipeSubPN.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPipeSubPN.Font = New System.Drawing.Font("Century Gothic", 14.0!)
        Me.txtPipeSubPN.ForeColor = System.Drawing.Color.FromArgb(CType(CType(76, Byte), Integer), CType(CType(77, Byte), Integer), CType(CType(80, Byte), Integer))
        Me.txtPipeSubPN.Location = New System.Drawing.Point(181, 35)
        Me.txtPipeSubPN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtPipeSubPN.MaxLength = 23
        Me.txtPipeSubPN.Name = "txtPipeSubPN"
        Me.txtPipeSubPN.Size = New System.Drawing.Size(72, 30)
        Me.txtPipeSubPN.TabIndex = 1
        Me.txtPipeSubPN.WordWrap = False
        '
        'gridAddPipeRecord
        '
        Me.gridAddPipeRecord.AllowUserToAddRows = False
        Me.gridAddPipeRecord.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridAddPipeRecord.Location = New System.Drawing.Point(10, 75)
        Me.gridAddPipeRecord.Name = "gridAddPipeRecord"
        Me.gridAddPipeRecord.Size = New System.Drawing.Size(817, 256)
        Me.gridAddPipeRecord.TabIndex = 10
        '
        'btnDelPipe
        '
        Me.btnDelPipe.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnDelPipe.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDelPipe.Font = New System.Drawing.Font("Century Gothic", 11.0!)
        Me.btnDelPipe.ForeColor = System.Drawing.Color.Black
        Me.btnDelPipe.Location = New System.Drawing.Point(1113, 15)
        Me.btnDelPipe.Margin = New System.Windows.Forms.Padding(0)
        Me.btnDelPipe.Name = "btnDelPipe"
        Me.btnDelPipe.Size = New System.Drawing.Size(120, 51)
        Me.btnDelPipe.TabIndex = 8
        Me.btnDelPipe.Text = "Odstrániť označenú"
        Me.btnDelPipe.UseVisualStyleBackColor = False
        '
        'txtPipeID
        '
        Me.txtPipeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPipeID.Font = New System.Drawing.Font("Century Gothic", 14.0!)
        Me.txtPipeID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(76, Byte), Integer), CType(CType(77, Byte), Integer), CType(CType(80, Byte), Integer))
        Me.txtPipeID.Location = New System.Drawing.Point(451, 4)
        Me.txtPipeID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtPipeID.MaxLength = 23
        Me.txtPipeID.Name = "txtPipeID"
        Me.txtPipeID.Size = New System.Drawing.Size(21, 30)
        Me.txtPipeID.TabIndex = 104
        Me.txtPipeID.Visible = False
        Me.txtPipeID.WordWrap = False
        '
        'btnEditPipe
        '
        Me.btnEditPipe.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnEditPipe.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEditPipe.Font = New System.Drawing.Font("Century Gothic", 11.0!)
        Me.btnEditPipe.ForeColor = System.Drawing.Color.Black
        Me.btnEditPipe.Location = New System.Drawing.Point(983, 15)
        Me.btnEditPipe.Margin = New System.Windows.Forms.Padding(0)
        Me.btnEditPipe.Name = "btnEditPipe"
        Me.btnEditPipe.Size = New System.Drawing.Size(120, 51)
        Me.btnEditPipe.TabIndex = 7
        Me.btnEditPipe.Text = "Upraviť existujúcu"
        Me.btnEditPipe.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label1.Location = New System.Drawing.Point(611, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(172, 17)
        Me.Label1.TabIndex = 111
        Me.Label1.Text = "Skontrolované voči výkresu"
        '
        'AddPipeRecordsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1242, 343)
        Me.Controls.Add(Me.gridAddPipeRecord)
        Me.Controls.Add(Me.btnDelPipe)
        Me.Controls.Add(Me.btnEditPipe)
        Me.Controls.Add(Me.btnSavePipe)
        Me.Controls.Add(Me.ckbPipeSkontrolovane)
        Me.Controls.Add(Me.cboPipeOhybRovna)
        Me.Controls.Add(Me.cboPipeHrubka)
        Me.Controls.Add(Me.cboPipePriemer)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.txtPipeID)
        Me.Controls.Add(Me.txtPipePN)
        Me.Controls.Add(Me.txtPipeDlzka)
        Me.Controls.Add(Me.txtPipeSubPN)
        Me.Name = "AddPipeRecordsForm"
        Me.Text = "AddPipeRecordsForm"
        CType(Me.gridAddPipeRecord, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnSavePipe As Button
    Friend WithEvents ckbPipeSkontrolovane As CheckBox
    Friend WithEvents cboPipeOhybRovna As ComboBox
    Friend WithEvents cboPipeHrubka As ComboBox
    Friend WithEvents cboPipePriemer As ComboBox
    Friend WithEvents Label18 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents txtPipePN As TextBox
    Friend WithEvents txtPipeDlzka As TextBox
    Friend WithEvents txtPipeSubPN As TextBox
    Friend WithEvents gridAddPipeRecord As DataGridView
    Friend WithEvents btnDelPipe As Button
    Friend WithEvents txtPipeID As TextBox
    Friend WithEvents btnEditPipe As Button
    Friend WithEvents Label1 As Label
End Class
