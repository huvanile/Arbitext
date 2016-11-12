Imports System.Net
Imports Arbitext.RegistryHelpers
Imports Arbitext.EmailHelpers

Public Class FrmPrefs

    Private Sub toggleEmailVisibility()
        With chkNotifyViaGmail
            If .Checked = True Then
                txtEmailAddress.Enabled = True
                txtEmailPassword.Enabled = True
                lblEmail1.Enabled = True
                lblEmail2.Enabled = True
                btnTestGmail.Enabled = True
            Else
                txtEmailAddress.Enabled = False
                txtEmailPassword.Enabled = False
                lblEmail1.Enabled = False
                lblEmail2.Enabled = False
                btnTestGmail.Enabled = False
            End If
        End With
    End Sub

    Private Sub togglePBVisibility()
        With chkNotifyViaPB
            If .Checked = True Then
                txtPBAPIKey.Enabled = True
                txtPBDeviceID.Enabled = True
                lblPB1.Enabled = True
                lblPB2.Enabled = True
                btnTestPB.Enabled = True
                btnFindDevices.Enabled = True
            Else
                txtPBAPIKey.Enabled = False
                txtPBDeviceID.Enabled = False
                lblPB1.Enabled = False
                lblPB2.Enabled = False
                btnTestPB.Enabled = False
                btnFindDevices.Enabled = False
            End If
        End With
    End Sub

    Private Sub FrmPrefs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        addressNullPrefs()
        loadVariablesFromRegistry()
        loadFormPrefsFromVariables()
        If txtPBDeviceID.Text = "your device index" Then txtPBDeviceID.Text = ""
        Me.Text = ThisAddIn.Title
        toggleEmailVisibility()
        togglePBVisibility()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        saveFormPrefsToRegistry()
        loadVariablesFromRegistry()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
        ThisAddIn.frmPrefs = Nothing
    End Sub

    Private Sub btnTestPB_Click(sender As Object, e As EventArgs) Handles btnTestPB.Click
        'sendPushbulletNote "CL-BS Connection", "This is a test note from the Craigslist BookScouter connection script.  It works.  Woo.  Now go make some money."
        'btnTestPB.Text = "Sent!"
    End Sub

    Private Sub btnFindDevices_Click(sender As Object, e As EventArgs) Handles btnFindDevices.Click
        'Dim resp As String
        'Dim d As Variant
        'Dim splitholder
        'resp = ShellRun("pythonw -c ""from pushbullet import Pushbullet ; pb=Pushbullet('" & thisaddin.pbapikey & "') ; print(pb.devices) """)
        'If resp Like "*Device*" Then
        '    resp = Replace(resp, "Device", "")
        '    splitholder = Split(resp, ",")
        '    resp = "Enter the index / number of your devices, as listed below:" & vbCrLf & vbCrLf
        '    For d = LBound(splitholder) To UBound(splitholder)
        '        resp = resp & "[" & d & "] " & clean(splitholder(d), False, True, False, False, True, False, False) & vbCrLf
        '    Next d
        '    MsgBox resp, vbInformation, title
        'Else
        '    MsgBox resp, vbInformation, title
        'End If
    End Sub

    Private Sub chkNotifyViaPB_CheckedChanged(sender As Object, e As EventArgs) Handles chkNotifyViaPB.CheckedChanged
        togglePBVisibility()
    End Sub

    Private Sub btnGrabCoverArt_Click(sender As Object, e As EventArgs) Handles btnGrabCoverArt.Click
        Dim webClient As New WebClient
        webClient.DownloadFile("http://datadump.thelightningpath.com/images/bookimages/PlaceholderBook.png", Environ("Temp") & "\cover.jpg")
        If System.IO.File.Exists(Environ("Temp") & "\cover.jpg") Then
            btnGrabCoverArt.Text = "Done!"
        Else
            btnGrabCoverArt.Text = "(try again)"
        End If
    End Sub

    Private Sub btnTestGmail_Click(sender As Object, e As EventArgs) Handles btnTestGmail.Click
        sendSilentNotification("<p>This is a test note from the Craigslist BookScouter connection script.  It works.  Woo.  Now go make some money.</p><img src='http://replygif.net/i/159.gif'/>", "Test email")
        btnTestGmail.Text = "Sent!"
    End Sub

    Private Sub chkNotifyViaGmail_CheckedChanged(sender As Object, e As EventArgs) Handles chkNotifyViaGmail.CheckedChanged
        toggleEmailVisibility()
    End Sub
End Class