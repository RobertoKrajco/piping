<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmReportSklad
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
        Me.rvSklad = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.SuspendLayout()
        '
        'rvSklad
        '
        Me.rvSklad.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rvSklad.LocalReport.ReportEmbeddedResource = "Piping_Plan_v01.reportSklad.rdlc"
        Me.rvSklad.Location = New System.Drawing.Point(0, 0)
        Me.rvSklad.Name = "rvSklad"
        Me.rvSklad.ServerReport.BearerToken = Nothing
        Me.rvSklad.Size = New System.Drawing.Size(1159, 585)
        Me.rvSklad.TabIndex = 0
        '
        'frmReportSklad
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1159, 585)
        Me.Controls.Add(Me.rvSklad)
        Me.Name = "frmReportSklad"
        Me.Text = "frmReportSklad"
        Me.ResumeLayout(False)

    End Sub

    Public WithEvents rvSklad As Microsoft.Reporting.WinForms.ReportViewer
End Class
