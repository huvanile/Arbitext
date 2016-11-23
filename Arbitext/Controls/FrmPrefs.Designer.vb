<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmPrefs
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPrefs))
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.grpPB = New System.Windows.Forms.GroupBox()
        Me.btnFindDevices = New System.Windows.Forms.Button()
        Me.btnTestPB = New System.Windows.Forms.Button()
        Me.txtPBDeviceID = New System.Windows.Forms.TextBox()
        Me.lblPB2 = New System.Windows.Forms.Label()
        Me.txtPBAPIKey = New System.Windows.Forms.TextBox()
        Me.lblPB1 = New System.Windows.Forms.Label()
        Me.chkNotifyViaPB = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtMinProfit = New System.Windows.Forms.TextBox()
        Me.chkOnMaybes = New System.Windows.Forms.CheckBox()
        Me.chkOnWinners = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.grpGmail = New System.Windows.Forms.GroupBox()
        Me.btnTestGmail = New System.Windows.Forms.Button()
        Me.txtEmailPassword = New System.Windows.Forms.TextBox()
        Me.lblEmail2 = New System.Windows.Forms.Label()
        Me.txtEmailAddress = New System.Windows.Forms.TextBox()
        Me.lblEmail1 = New System.Windows.Forms.Label()
        Me.chkNotifyViaGmail = New System.Windows.Forms.CheckBox()
        Me.grpSorting = New System.Windows.Forms.GroupBox()
        Me.txtSFTPDirectory = New System.Windows.Forms.TextBox()
        Me.txtSFTPPass = New System.Windows.Forms.TextBox()
        Me.txtSFTPUser = New System.Windows.Forms.TextBox()
        Me.txtSFTPURL = New System.Windows.Forms.TextBox()
        Me.grpOnlyShowPosts = New System.Windows.Forms.GroupBox()
        Me.optShowAll = New System.Windows.Forms.RadioButton()
        Me.optUpdated14Days = New System.Windows.Forms.RadioButton()
        Me.optUpdated7Days = New System.Windows.Forms.RadioButton()
        Me.optUpdatedToday = New System.Windows.Forms.RadioButton()
        Me.optPostedToday = New System.Windows.Forms.RadioButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.grpPB.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.grpGmail.SuspendLayout()
        Me.grpSorting.SuspendLayout()
        Me.grpOnlyShowPosts.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(619, 414)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(132, 66)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(760, 414)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(134, 66)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'grpPB
        '
        Me.grpPB.Controls.Add(Me.btnFindDevices)
        Me.grpPB.Controls.Add(Me.btnTestPB)
        Me.grpPB.Controls.Add(Me.txtPBDeviceID)
        Me.grpPB.Controls.Add(Me.lblPB2)
        Me.grpPB.Controls.Add(Me.txtPBAPIKey)
        Me.grpPB.Controls.Add(Me.lblPB1)
        Me.grpPB.Controls.Add(Me.chkNotifyViaPB)
        Me.grpPB.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpPB.Location = New System.Drawing.Point(303, 18)
        Me.grpPB.Name = "grpPB"
        Me.grpPB.Size = New System.Drawing.Size(304, 152)
        Me.grpPB.TabIndex = 2
        Me.grpPB.TabStop = False
        Me.grpPB.Text = "Pushbullet Notification Settings"
        '
        'btnFindDevices
        '
        Me.btnFindDevices.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFindDevices.Location = New System.Drawing.Point(242, 97)
        Me.btnFindDevices.Name = "btnFindDevices"
        Me.btnFindDevices.Size = New System.Drawing.Size(47, 21)
        Me.btnFindDevices.TabIndex = 14
        Me.btnFindDevices.Text = "Find"
        Me.btnFindDevices.UseVisualStyleBackColor = True
        '
        'btnTestPB
        '
        Me.btnTestPB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestPB.Location = New System.Drawing.Point(242, 19)
        Me.btnTestPB.Name = "btnTestPB"
        Me.btnTestPB.Size = New System.Drawing.Size(47, 21)
        Me.btnTestPB.TabIndex = 13
        Me.btnTestPB.Text = "Test"
        Me.btnTestPB.UseVisualStyleBackColor = True
        '
        'txtPBDeviceID
        '
        Me.txtPBDeviceID.Location = New System.Drawing.Point(16, 119)
        Me.txtPBDeviceID.Name = "txtPBDeviceID"
        Me.txtPBDeviceID.Size = New System.Drawing.Size(273, 21)
        Me.txtPBDeviceID.TabIndex = 12
        '
        'lblPB2
        '
        Me.lblPB2.AutoSize = True
        Me.lblPB2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPB2.Location = New System.Drawing.Point(13, 99)
        Me.lblPB2.Name = "lblPB2"
        Me.lblPB2.Size = New System.Drawing.Size(232, 15)
        Me.lblPB2.TabIndex = 11
        Me.lblPB2.Text = "Pushbullet device index, as found with -->"
        '
        'txtPBAPIKey
        '
        Me.txtPBAPIKey.Location = New System.Drawing.Point(16, 69)
        Me.txtPBAPIKey.Name = "txtPBAPIKey"
        Me.txtPBAPIKey.Size = New System.Drawing.Size(273, 21)
        Me.txtPBAPIKey.TabIndex = 10
        '
        'lblPB1
        '
        Me.lblPB1.AutoSize = True
        Me.lblPB1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPB1.Location = New System.Drawing.Point(13, 49)
        Me.lblPB1.Name = "lblPB1"
        Me.lblPB1.Size = New System.Drawing.Size(110, 15)
        Me.lblPB1.TabIndex = 9
        Me.lblPB1.Text = "Pushbullet API key:"
        '
        'chkNotifyViaPB
        '
        Me.chkNotifyViaPB.AutoSize = True
        Me.chkNotifyViaPB.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNotifyViaPB.Location = New System.Drawing.Point(16, 21)
        Me.chkNotifyViaPB.Name = "chkNotifyViaPB"
        Me.chkNotifyViaPB.Size = New System.Drawing.Size(195, 19)
        Me.chkNotifyViaPB.TabIndex = 4
        Me.chkNotifyViaPB.Text = "Enable PushBullet notifications"
        Me.chkNotifyViaPB.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtCity)
        Me.GroupBox2.Controls.Add(Me.txtMinProfit)
        Me.GroupBox2.Controls.Add(Me.chkOnMaybes)
        Me.GroupBox2.Controls.Add(Me.chkOnWinners)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(619, 18)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(276, 173)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "General settings"
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(18, 83)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(237, 21)
        Me.txtCity.TabIndex = 8
        '
        'txtMinProfit
        '
        Me.txtMinProfit.Location = New System.Drawing.Point(193, 22)
        Me.txtMinProfit.Name = "txtMinProfit"
        Me.txtMinProfit.Size = New System.Drawing.Size(62, 21)
        Me.txtMinProfit.TabIndex = 7
        '
        'chkOnMaybes
        '
        Me.chkOnMaybes.AutoSize = True
        Me.chkOnMaybes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOnMaybes.Location = New System.Drawing.Point(18, 142)
        Me.chkOnMaybes.Name = "chkOnMaybes"
        Me.chkOnMaybes.Size = New System.Drawing.Size(119, 19)
        Me.chkOnMaybes.TabIndex = 4
        Me.chkOnMaybes.Text = "Notify on maybes"
        Me.chkOnMaybes.UseVisualStyleBackColor = True
        '
        'chkOnWinners
        '
        Me.chkOnWinners.AutoSize = True
        Me.chkOnWinners.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOnWinners.Location = New System.Drawing.Point(18, 117)
        Me.chkOnWinners.Name = "chkOnWinners"
        Me.chkOnWinners.Size = New System.Drawing.Size(119, 19)
        Me.chkOnWinners.TabIndex = 3
        Me.chkOnWinners.Text = "Notify on winners"
        Me.chkOnWinners.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(18, 63)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(98, 15)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "City (e.g., Dallas)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(18, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 30)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Minimum tolerable profit:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Enter 0 to disable)"
        '
        'grpGmail
        '
        Me.grpGmail.Controls.Add(Me.btnTestGmail)
        Me.grpGmail.Controls.Add(Me.txtEmailPassword)
        Me.grpGmail.Controls.Add(Me.lblEmail2)
        Me.grpGmail.Controls.Add(Me.txtEmailAddress)
        Me.grpGmail.Controls.Add(Me.lblEmail1)
        Me.grpGmail.Controls.Add(Me.chkNotifyViaGmail)
        Me.grpGmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpGmail.Location = New System.Drawing.Point(303, 180)
        Me.grpGmail.Name = "grpGmail"
        Me.grpGmail.Size = New System.Drawing.Size(304, 155)
        Me.grpGmail.TabIndex = 4
        Me.grpGmail.TabStop = False
        Me.grpGmail.Text = "Gmail Notification Settings"
        '
        'btnTestGmail
        '
        Me.btnTestGmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestGmail.Location = New System.Drawing.Point(243, 17)
        Me.btnTestGmail.Name = "btnTestGmail"
        Me.btnTestGmail.Size = New System.Drawing.Size(47, 21)
        Me.btnTestGmail.TabIndex = 19
        Me.btnTestGmail.Text = "Test"
        Me.btnTestGmail.UseVisualStyleBackColor = True
        '
        'txtEmailPassword
        '
        Me.txtEmailPassword.Location = New System.Drawing.Point(17, 117)
        Me.txtEmailPassword.Name = "txtEmailPassword"
        Me.txtEmailPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtEmailPassword.Size = New System.Drawing.Size(273, 21)
        Me.txtEmailPassword.TabIndex = 18
        '
        'lblEmail2
        '
        Me.lblEmail2.AutoSize = True
        Me.lblEmail2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmail2.Location = New System.Drawing.Point(14, 97)
        Me.lblEmail2.Name = "lblEmail2"
        Me.lblEmail2.Size = New System.Drawing.Size(153, 15)
        Me.lblEmail2.TabIndex = 17
        Me.lblEmail2.Text = "Google Account Password:"
        '
        'txtEmailAddress
        '
        Me.txtEmailAddress.Location = New System.Drawing.Point(17, 67)
        Me.txtEmailAddress.Name = "txtEmailAddress"
        Me.txtEmailAddress.Size = New System.Drawing.Size(273, 21)
        Me.txtEmailAddress.TabIndex = 16
        '
        'lblEmail1
        '
        Me.lblEmail1.AutoSize = True
        Me.lblEmail1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmail1.Location = New System.Drawing.Point(14, 47)
        Me.lblEmail1.Name = "lblEmail1"
        Me.lblEmail1.Size = New System.Drawing.Size(132, 15)
        Me.lblEmail1.TabIndex = 15
        Me.lblEmail1.Text = "Google Email Address:"
        '
        'chkNotifyViaGmail
        '
        Me.chkNotifyViaGmail.AutoSize = True
        Me.chkNotifyViaGmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNotifyViaGmail.Location = New System.Drawing.Point(17, 19)
        Me.chkNotifyViaGmail.Name = "chkNotifyViaGmail"
        Me.chkNotifyViaGmail.Size = New System.Drawing.Size(169, 19)
        Me.chkNotifyViaGmail.TabIndex = 14
        Me.chkNotifyViaGmail.Text = "Enable Gmail notifications"
        Me.chkNotifyViaGmail.UseVisualStyleBackColor = True
        '
        'grpSorting
        '
        Me.grpSorting.Controls.Add(Me.txtSFTPDirectory)
        Me.grpSorting.Controls.Add(Me.txtSFTPPass)
        Me.grpSorting.Controls.Add(Me.txtSFTPUser)
        Me.grpSorting.Controls.Add(Me.txtSFTPURL)
        Me.grpSorting.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpSorting.Location = New System.Drawing.Point(303, 347)
        Me.grpSorting.Name = "grpSorting"
        Me.grpSorting.Size = New System.Drawing.Size(304, 133)
        Me.grpSorting.TabIndex = 5
        Me.grpSorting.TabStop = False
        Me.grpSorting.Text = "SFTP Config"
        '
        'txtSFTPDirectory
        '
        Me.txtSFTPDirectory.Location = New System.Drawing.Point(17, 105)
        Me.txtSFTPDirectory.Name = "txtSFTPDirectory"
        Me.txtSFTPDirectory.Size = New System.Drawing.Size(273, 21)
        Me.txtSFTPDirectory.TabIndex = 3
        '
        'txtSFTPPass
        '
        Me.txtSFTPPass.Location = New System.Drawing.Point(16, 78)
        Me.txtSFTPPass.Name = "txtSFTPPass"
        Me.txtSFTPPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSFTPPass.Size = New System.Drawing.Size(273, 21)
        Me.txtSFTPPass.TabIndex = 2
        '
        'txtSFTPUser
        '
        Me.txtSFTPUser.Location = New System.Drawing.Point(16, 51)
        Me.txtSFTPUser.Name = "txtSFTPUser"
        Me.txtSFTPUser.Size = New System.Drawing.Size(273, 21)
        Me.txtSFTPUser.TabIndex = 1
        '
        'txtSFTPURL
        '
        Me.txtSFTPURL.Location = New System.Drawing.Point(16, 24)
        Me.txtSFTPURL.Name = "txtSFTPURL"
        Me.txtSFTPURL.Size = New System.Drawing.Size(273, 21)
        Me.txtSFTPURL.TabIndex = 0
        '
        'grpOnlyShowPosts
        '
        Me.grpOnlyShowPosts.Controls.Add(Me.optShowAll)
        Me.grpOnlyShowPosts.Controls.Add(Me.optUpdated14Days)
        Me.grpOnlyShowPosts.Controls.Add(Me.optUpdated7Days)
        Me.grpOnlyShowPosts.Controls.Add(Me.optUpdatedToday)
        Me.grpOnlyShowPosts.Controls.Add(Me.optPostedToday)
        Me.grpOnlyShowPosts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpOnlyShowPosts.Location = New System.Drawing.Point(619, 199)
        Me.grpOnlyShowPosts.Name = "grpOnlyShowPosts"
        Me.grpOnlyShowPosts.Size = New System.Drawing.Size(276, 154)
        Me.grpOnlyShowPosts.TabIndex = 6
        Me.grpOnlyShowPosts.TabStop = False
        Me.grpOnlyShowPosts.Text = "Only show posts"
        '
        'optShowAll
        '
        Me.optShowAll.AutoSize = True
        Me.optShowAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optShowAll.Location = New System.Drawing.Point(21, 121)
        Me.optShowAll.Name = "optShowAll"
        Me.optShowAll.Size = New System.Drawing.Size(185, 19)
        Me.optShowAll.TabIndex = 4
        Me.optShowAll.TabStop = True
        Me.optShowAll.Text = "(Show all posts, no time filter)"
        Me.optShowAll.UseVisualStyleBackColor = True
        '
        'optUpdated14Days
        '
        Me.optUpdated14Days.AutoSize = True
        Me.optUpdated14Days.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUpdated14Days.Location = New System.Drawing.Point(21, 96)
        Me.optUpdated14Days.Name = "optUpdated14Days"
        Me.optUpdated14Days.Size = New System.Drawing.Size(152, 19)
        Me.optUpdated14Days.TabIndex = 3
        Me.optUpdated14Days.TabStop = True
        Me.optUpdated14Days.Text = "Updated in last 14 days"
        Me.optUpdated14Days.UseVisualStyleBackColor = True
        '
        'optUpdated7Days
        '
        Me.optUpdated7Days.AutoSize = True
        Me.optUpdated7Days.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUpdated7Days.Location = New System.Drawing.Point(21, 71)
        Me.optUpdated7Days.Name = "optUpdated7Days"
        Me.optUpdated7Days.Size = New System.Drawing.Size(145, 19)
        Me.optUpdated7Days.TabIndex = 2
        Me.optUpdated7Days.TabStop = True
        Me.optUpdated7Days.Text = "Updated in last 7 days"
        Me.optUpdated7Days.UseVisualStyleBackColor = True
        '
        'optUpdatedToday
        '
        Me.optUpdatedToday.AutoSize = True
        Me.optUpdatedToday.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUpdatedToday.Location = New System.Drawing.Point(21, 46)
        Me.optUpdatedToday.Name = "optUpdatedToday"
        Me.optUpdatedToday.Size = New System.Drawing.Size(104, 19)
        Me.optUpdatedToday.TabIndex = 1
        Me.optUpdatedToday.TabStop = True
        Me.optUpdatedToday.Text = "Updated today"
        Me.optUpdatedToday.UseVisualStyleBackColor = True
        '
        'optPostedToday
        '
        Me.optPostedToday.AutoSize = True
        Me.optPostedToday.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optPostedToday.Location = New System.Drawing.Point(21, 21)
        Me.optPostedToday.Name = "optPostedToday"
        Me.optPostedToday.Size = New System.Drawing.Size(95, 19)
        Me.optPostedToday.TabIndex = 0
        Me.optPostedToday.TabStop = True
        Me.optPostedToday.Text = "Posted today"
        Me.optPostedToday.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(22, 57)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(263, 398)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(261, 28)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Set Those Preferences"
        '
        'FrmPrefs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(909, 492)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.grpOnlyShowPosts)
        Me.Controls.Add(Me.grpSorting)
        Me.Controls.Add(Me.grpGmail)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.grpPB)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Name = "FrmPrefs"
        Me.Text = "FrmPrefs"
        Me.grpPB.ResumeLayout(False)
        Me.grpPB.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.grpGmail.ResumeLayout(False)
        Me.grpGmail.PerformLayout()
        Me.grpSorting.ResumeLayout(False)
        Me.grpSorting.PerformLayout()
        Me.grpOnlyShowPosts.ResumeLayout(False)
        Me.grpOnlyShowPosts.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents grpPB As System.Windows.Forms.GroupBox
    Friend WithEvents btnFindDevices As System.Windows.Forms.Button
    Friend WithEvents btnTestPB As System.Windows.Forms.Button
    Friend WithEvents txtPBDeviceID As System.Windows.Forms.TextBox
    Friend WithEvents lblPB2 As System.Windows.Forms.Label
    Friend WithEvents txtPBAPIKey As System.Windows.Forms.TextBox
    Friend WithEvents lblPB1 As System.Windows.Forms.Label
    Friend WithEvents chkNotifyViaPB As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtMinProfit As System.Windows.Forms.TextBox
    Friend WithEvents chkOnMaybes As System.Windows.Forms.CheckBox
    Friend WithEvents chkOnWinners As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents grpGmail As System.Windows.Forms.GroupBox
    Friend WithEvents grpSorting As System.Windows.Forms.GroupBox
    Friend WithEvents grpOnlyShowPosts As System.Windows.Forms.GroupBox
    Friend WithEvents optShowAll As System.Windows.Forms.RadioButton
    Friend WithEvents optUpdated14Days As System.Windows.Forms.RadioButton
    Friend WithEvents optUpdated7Days As System.Windows.Forms.RadioButton
    Friend WithEvents optUpdatedToday As System.Windows.Forms.RadioButton
    Friend WithEvents optPostedToday As System.Windows.Forms.RadioButton
    Friend WithEvents btnTestGmail As System.Windows.Forms.Button
    Friend WithEvents txtEmailPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail2 As System.Windows.Forms.Label
    Friend WithEvents txtEmailAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail1 As System.Windows.Forms.Label
    Friend WithEvents chkNotifyViaGmail As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSFTPPass As System.Windows.Forms.TextBox
    Friend WithEvents txtSFTPUser As System.Windows.Forms.TextBox
    Friend WithEvents txtSFTPURL As System.Windows.Forms.TextBox
    Friend WithEvents txtSFTPDirectory As System.Windows.Forms.TextBox
End Class
