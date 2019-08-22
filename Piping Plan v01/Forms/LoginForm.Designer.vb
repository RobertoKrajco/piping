<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoginForm1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LoginForm1))
        Me.cmbProfil = New System.Windows.Forms.ComboBox()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.OK = New System.Windows.Forms.Button()
        Me.PasswordTextBox = New System.Windows.Forms.TextBox()
        Me.UsernameTextBox = New System.Windows.Forms.TextBox()
        Me.LogoPictureBox = New System.Windows.Forms.PictureBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbProfil
        '
        Me.cmbProfil.Font = New System.Drawing.Font("Century Gothic", 12.0!)
        Me.cmbProfil.FormattingEnabled = True
        Me.cmbProfil.Items.AddRange(New Object() {"operator", "majster", "planovac", "admin", "inzinier"})
        Me.cmbProfil.Location = New System.Drawing.Point(224, 21)
        Me.cmbProfil.Name = "cmbProfil"
        Me.cmbProfil.Size = New System.Drawing.Size(163, 29)
        Me.cmbProfil.TabIndex = 15
        Me.cmbProfil.Text = "admin"
        '
        'Cancel
        '
        Me.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel.Font = New System.Drawing.Font("Century Gothic", 11.0!)
        Me.Cancel.Location = New System.Drawing.Point(311, 155)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(76, 30)
        Me.Cancel.TabIndex = 14
        Me.Cancel.Text = "&Cancel"
        '
        'OK
        '
        Me.OK.Font = New System.Drawing.Font("Century Gothic", 11.0!)
        Me.OK.Location = New System.Drawing.Point(224, 155)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(76, 30)
        Me.OK.TabIndex = 13
        Me.OK.Text = "&OK"
        '
        'PasswordTextBox
        '
        Me.PasswordTextBox.Font = New System.Drawing.Font("Century Gothic", 12.0!)
        Me.PasswordTextBox.Location = New System.Drawing.Point(224, 121)
        Me.PasswordTextBox.Name = "PasswordTextBox"
        Me.PasswordTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.PasswordTextBox.Size = New System.Drawing.Size(163, 27)
        Me.PasswordTextBox.TabIndex = 12
        Me.PasswordTextBox.Text = "Admin@"
        '
        'UsernameTextBox
        '
        Me.UsernameTextBox.Font = New System.Drawing.Font("Century Gothic", 12.0!)
        Me.UsernameTextBox.Location = New System.Drawing.Point(224, 72)
        Me.UsernameTextBox.Name = "UsernameTextBox"
        Me.UsernameTextBox.Size = New System.Drawing.Size(163, 27)
        Me.UsernameTextBox.TabIndex = 10
        Me.UsernameTextBox.Text = "admin"
        '
        'LogoPictureBox
        '
        Me.LogoPictureBox.Image = CType(resources.GetObject("LogoPictureBox.Image"), System.Drawing.Image)
        Me.LogoPictureBox.Location = New System.Drawing.Point(-2, -2)
        Me.LogoPictureBox.Name = "LogoPictureBox"
        Me.LogoPictureBox.Size = New System.Drawing.Size(202, 193)
        Me.LogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.LogoPictureBox.TabIndex = 9
        Me.LogoPictureBox.TabStop = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label17.Location = New System.Drawing.Point(222, 3)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(105, 17)
        Me.Label17.TabIndex = 113
        Me.Label17.Text = "Užívateľský profil"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label1.Location = New System.Drawing.Point(222, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 17)
        Me.Label1.TabIndex = 113
        Me.Label1.Text = "Osobné číslo"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Century Gothic", 9.0!)
        Me.Label2.Location = New System.Drawing.Point(222, 101)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 17)
        Me.Label2.TabIndex = 113
        Me.Label2.Text = "Heslo"
        '
        'LoginForm1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(408, 191)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.cmbProfil)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.PasswordTextBox)
        Me.Controls.Add(Me.UsernameTextBox)
        Me.Controls.Add(Me.LogoPictureBox)
        Me.Name = "LoginForm1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "LoginForm"
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbProfil As ComboBox
    Friend WithEvents Cancel As Button
    Friend WithEvents OK As Button
    Friend WithEvents PasswordTextBox As TextBox
    Friend WithEvents UsernameTextBox As TextBox
    Friend WithEvents LogoPictureBox As PictureBox
    Friend WithEvents Label17 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
End Class
