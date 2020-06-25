<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
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
        Me.txtServerNameIP = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbxDetails = New System.Windows.Forms.ListBox()
        Me.lblServerNameIP = New System.Windows.Forms.Label()
        Me.cmdScan = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtServerNameIP
        '
        Me.txtServerNameIP.Location = New System.Drawing.Point(129, 6)
        Me.txtServerNameIP.Name = "txtServerNameIP"
        Me.txtServerNameIP.Size = New System.Drawing.Size(205, 22)
        Me.txtServerNameIP.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Server Name/IP:"
        '
        'lbxDetails
        '
        Me.lbxDetails.Font = New System.Drawing.Font("DejaVu Sans Mono", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbxDetails.FormattingEnabled = True
        Me.lbxDetails.ItemHeight = 15
        Me.lbxDetails.Location = New System.Drawing.Point(12, 77)
        Me.lbxDetails.Name = "lbxDetails"
        Me.lbxDetails.Size = New System.Drawing.Size(776, 364)
        Me.lbxDetails.TabIndex = 2
        '
        'lblServerNameIP
        '
        Me.lblServerNameIP.AutoSize = True
        Me.lblServerNameIP.Location = New System.Drawing.Point(126, 31)
        Me.lblServerNameIP.Name = "lblServerNameIP"
        Me.lblServerNameIP.Size = New System.Drawing.Size(113, 17)
        Me.lblServerNameIP.TabIndex = 3
        Me.lblServerNameIP.Text = "lblServerNameIP"
        '
        'cmdScan
        '
        Me.cmdScan.Location = New System.Drawing.Point(340, 6)
        Me.cmdScan.Name = "cmdScan"
        Me.cmdScan.Size = New System.Drawing.Size(75, 23)
        Me.cmdScan.TabIndex = 4
        Me.cmdScan.Text = "Scan"
        Me.cmdScan.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cmdScan)
        Me.Controls.Add(Me.lblServerNameIP)
        Me.Controls.Add(Me.lbxDetails)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtServerNameIP)
        Me.Name = "frmMain"
        Me.Text = "Simple System Info"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtServerNameIP As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lbxDetails As ListBox
    Friend WithEvents lblServerNameIP As Label
    Friend WithEvents cmdScan As Button
End Class
