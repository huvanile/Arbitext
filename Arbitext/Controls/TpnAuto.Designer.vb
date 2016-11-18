<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TpnAuto
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnProceed = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblNumber = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblStatus.Location = New System.Drawing.Point(53, 142)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(191, 93)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Press ""Proceed"" to begin"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnProceed
        '
        Me.btnProceed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProceed.Location = New System.Drawing.Point(53, 60)
        Me.btnProceed.Name = "btnProceed"
        Me.btnProceed.Size = New System.Drawing.Size(75, 23)
        Me.btnProceed.TabIndex = 1
        Me.btnProceed.Text = "Proceed"
        Me.btnProceed.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(169, 60)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblNumber
        '
        Me.lblNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNumber.ForeColor = System.Drawing.Color.Maroon
        Me.lblNumber.Location = New System.Drawing.Point(53, 103)
        Me.lblNumber.Name = "lblNumber"
        Me.lblNumber.Size = New System.Drawing.Size(191, 23)
        Me.lblNumber.TabIndex = 3
        Me.lblNumber.Text = "Search Result #:  "
        Me.lblNumber.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblNumber.Visible = False
        '
        'TpnAuto
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lblNumber)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnProceed)
        Me.Controls.Add(Me.lblStatus)
        Me.Name = "TpnAuto"
        Me.Size = New System.Drawing.Size(300, 575)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents btnProceed As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblNumber As System.Windows.Forms.Label
End Class
