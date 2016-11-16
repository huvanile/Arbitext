Imports Microsoft.Win32

Public Class RegistryHelpers
    Public Const RegistryFolder As String = "HKEY_CURRENT_USER\SOFTWARE\ARBITEXT\"

    Public Shared Function GetSetting(subFolder As String, keyName As String)
        Dim myKey As RegistryKey = Registry.LocalMachine.OpenSubKey(RegistryFolder + subFolder + "\", True)
        If Not IsNothing(myKey) Then
            Dim keyObject As Object
            keyObject = myKey.GetValue(keyName)
            myKey.Close()
            If IsNothing(keyObject) Then
                Return ""
            Else
                Return keyObject.ToString()
            End If
        Else
            Return ""
        End If
    End Function

    Public Shared Sub SaveSetting(subFolder As String, keyName As String, keyValue As String)
        Registry.LocalMachine.CreateSubKey(RegistryFolder + subFolder + "\")
        Dim myKey As RegistryKey = Registry.LocalMachine.OpenSubKey(RegistryFolder + subFolder + "\", True)
        If Not IsNothing(myKey) Then
            myKey.SetValue(keyName, keyValue, RegistryValueKind.String)
            myKey.Close()
        End If
    End Sub

    ''' <summary>
    ''' Used to populate the registry with default preference values
    ''' Prevents runtime errors which could occur if a null pref value is used
    ''' </summary>
    Public Shared Sub addressNullPrefs()
        Try

            'booleans
            If My.Computer.Registry.GetValue(RegistryFolder, "EmailsOK", Nothing) Is Nothing Then My.Computer.Registry.SetValue(RegistryFolder, "EmailsOK", False)
            If My.Computer.Registry.GetValue(RegistryFolder, "NotifyViaPBOK", Nothing) Is Nothing Then My.Computer.Registry.SetValue(RegistryFolder, "NotifyViaPBOK", False)
            If My.Computer.Registry.GetValue(RegistryFolder, "KeepIEVisibleOK", Nothing) Is Nothing Then My.Computer.Registry.SetValue(RegistryFolder, "KeepIEVisibleOK", False)
            If My.Computer.Registry.GetValue(RegistryFolder, "AutoCategorizeOK", Nothing) Is Nothing Then My.Computer.Registry.SetValue(RegistryFolder, "AutoCategorizeOK", True)
            If My.Computer.Registry.GetValue(RegistryFolder, "OnWinnersOK", Nothing) Is Nothing Then My.Computer.Registry.SetValue(RegistryFolder, "OnWinnersOK", False)
            If My.Computer.Registry.GetValue(RegistryFolder, "OnMaybesOK", Nothing) Is Nothing Then My.Computer.Registry.SetValue(RegistryFolder, "OnMaybesOK", False)

            'numbers
            If My.Computer.Registry.GetValue(RegistryFolder, "MaxResults", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "MaxResults", Nothing) = 0 Then _
                My.Computer.Registry.SetValue(RegistryFolder, "MaxResults", 500)
            If My.Computer.Registry.GetValue(RegistryFolder, "MinTolerableProfit", Nothing) Is Nothing Then _
                My.Computer.Registry.SetValue(RegistryFolder, "MinTolerableProfit", 45)

            'strings
            If My.Computer.Registry.GetValue(RegistryFolder, "PbAPIKey", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "PbAPIKey", Nothing) = "" _
                Then My.Computer.Registry.SetValue(RegistryFolder, "PbAPIKey", "<enter your pushbullet API key>")
            If My.Computer.Registry.GetValue(RegistryFolder, "PbDeviceID", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "PbDeviceID", Nothing) = "" _
                Then My.Computer.Registry.SetValue(RegistryFolder, "PbDeviceID", "<enter your pushbullet device ID>")
            If My.Computer.Registry.GetValue(RegistryFolder, "EmailAddress", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "EmailAddress", Nothing) = "" _
                Then My.Computer.Registry.SetValue(RegistryFolder, "EmailAddress", "<enter your gmail address>")
            If My.Computer.Registry.GetValue(RegistryFolder, "EmailPassword", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "EmailPassword", Nothing) = "" _
                Then My.Computer.Registry.SetValue(RegistryFolder, "EmailPassword", "<enter your gmail password>")
            If My.Computer.Registry.GetValue(RegistryFolder, "City", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "City", Nothing) = "" _
                Then My.Computer.Registry.SetValue(RegistryFolder, "City", "Dallas")
            If My.Computer.Registry.GetValue(RegistryFolder, "PostTimingPref", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "PostTimingPref", Nothing) = "" _
                Then My.Computer.Registry.SetValue(RegistryFolder, "PostTimingPref", "timingUpdated14Days")

        Catch exAll As Exception
            MsgBox("An error has occurred when pulling in the Arbitext add-in's default preferences.  Error message:" & vbCrLf & vbCrLf & exAll.Message, vbCritical, ThisAddIn.Title)
        End Try
    End Sub

    ''' <summary>
    ''' Used to populate the global / shared variables from their stored registry values
    ''' </summary>
    Public Shared Sub loadVariablesFromRegistry()
        Try

            'booleans
            ThisAddIn.EmailsOK = CBool(My.Computer.Registry.GetValue(RegistryFolder, "EmailsOK", Nothing))
            ThisAddIn.NotifyViaPBOK = CBool(My.Computer.Registry.GetValue(RegistryFolder, "NotifyViaPBOK", Nothing))
            ThisAddIn.KeepIEVisibleOK = CBool(My.Computer.Registry.GetValue(RegistryFolder, "KeepIEVisibleOK", Nothing))
            ThisAddIn.AutoCategorizeOK = CBool(My.Computer.Registry.GetValue(RegistryFolder, "AutoCategorizeOK", Nothing))
            ThisAddIn.OnWinnersOK = CBool(My.Computer.Registry.GetValue(RegistryFolder, "OnWinnersOK", Nothing))
            ThisAddIn.OnMaybesOK = CBool(My.Computer.Registry.GetValue(RegistryFolder, "OnMaybesOK", Nothing))

            'other
            ThisAddIn.PbAPIKey = My.Computer.Registry.GetValue(RegistryFolder, "PbAPIKey", Nothing)
            ThisAddIn.PbDeviceID = My.Computer.Registry.GetValue(RegistryFolder, "PbDeviceID", Nothing)
            ThisAddIn.EmailAddress = My.Computer.Registry.GetValue(RegistryFolder, "EmailAddress", Nothing)
            ThisAddIn.EmailPassword = My.Computer.Registry.GetValue(RegistryFolder, "EmailPassword", Nothing)
            ThisAddIn.MaxResults = My.Computer.Registry.GetValue(RegistryFolder, "MaxResults", Nothing)
            ThisAddIn.MinTolerableProfit = My.Computer.Registry.GetValue(RegistryFolder, "MinTolerableProfit", Nothing)
            ThisAddIn.PostTimingPref = My.Computer.Registry.GetValue(RegistryFolder, "PostTimingPref", Nothing)
            ThisAddIn.SaveWBFileName = My.Computer.Registry.GetValue(RegistryFolder, "SaveWBFileName", Nothing)
            ThisAddIn.SaveWBFilePath = My.Computer.Registry.GetValue(RegistryFolder, "SaveWBFilePath", Nothing)
            ThisAddIn.City = My.Computer.Registry.GetValue(RegistryFolder, "City", Nothing)

        Catch exAll As Exception
            MsgBox("An error has occurred when getting the bot's preferences.  Please close and reopen this bot and try again.", vbCritical, ThisAddIn.Title)
        End Try
    End Sub

    ''' <summary>
    ''' Used to retrieve user preferences from their varaibles upon prefs form initialize
    ''' </summary>
    Public Shared Sub loadFormPrefsFromVariables()
        Try

            'checkboxes
            ThisAddIn.frmPrefs.chkNotifyViaPB.Checked = ThisAddIn.NotifyViaPBOK
            ThisAddIn.frmPrefs.chkNotifyViaGmail.Checked = ThisAddIn.EmailsOK
            ThisAddIn.frmPrefs.chkAutoCategorize.Checked = ThisAddIn.AutoCategorizeOK
            ThisAddIn.frmPrefs.chkOnMaybes.Checked = ThisAddIn.OnMaybesOK
            ThisAddIn.frmPrefs.chkOnWinners.Checked = ThisAddIn.OnWinnersOK

            'textboxes
            ThisAddIn.frmPrefs.txtPBAPIKey.Text = ThisAddIn.PbAPIKey
            ThisAddIn.frmPrefs.txtPBDeviceID.Text = ThisAddIn.PbDeviceID
            ThisAddIn.frmPrefs.txtEmailAddress.Text = ThisAddIn.EmailAddress
            ThisAddIn.frmPrefs.txtEmailPassword.Text = ThisAddIn.EmailPassword
            ThisAddIn.frmPrefs.txtResultAmt.Text = ThisAddIn.MaxResults
            ThisAddIn.frmPrefs.txtMinProfit.Text = ThisAddIn.MinTolerableProfit
            ThisAddIn.frmPrefs.txtCity.Text = ThisAddIn.City

            'radio buttons
            Select Case ThisAddIn.PostTimingPref
                Case "timingPostedToday"
                    ThisAddIn.frmPrefs.optPostedToday.Checked = True
                Case "timingUpdatedToday"
                    ThisAddIn.frmPrefs.optUpdatedToday.Checked = True
                Case "timingUpdated7Days"
                    ThisAddIn.frmPrefs.optUpdated7Days.Checked = True
                Case "timingUpdated14Days"
                    ThisAddIn.frmPrefs.optUpdated14Days.Checked = True
                Case "timingShowAll"
                    ThisAddIn.frmPrefs.optShowAll.Checked = True
            End Select

        Catch exAll As Exception
            MsgBox("An error has occurred when loading the bot's preferences.  Please close and reopen this bot and try again.", vbCritical, ThisAddIn.Title)
        End Try
    End Sub

    ''' <summary>
    ''' Used to save the preferences as set in frmPrefs into the registry
    ''' </summary>
    Public Shared Sub saveFormPrefsToRegistry()
        Try

            With My.Computer.Registry

                'checkboxes
                .SetValue(RegistryFolder, "EmailsOK", ThisAddIn.frmPrefs.chkNotifyViaGmail.Checked)
                .SetValue(RegistryFolder, "NotifyViaPBOK", ThisAddIn.frmPrefs.chkNotifyViaPB.Checked)
                .SetValue(RegistryFolder, "AutoCategorizeOK", ThisAddIn.frmPrefs.chkAutoCategorize.Checked)
                .SetValue(RegistryFolder, "OnWinnersOK", ThisAddIn.frmPrefs.chkOnWinners.Checked)
                .SetValue(RegistryFolder, "OnMaybesOK", ThisAddIn.frmPrefs.chkOnMaybes.Checked)

                'other
                .SetValue(RegistryFolder, "PbAPIKey", ThisAddIn.frmPrefs.txtPBAPIKey.Text)
                .SetValue(RegistryFolder, "PbDeviceID", ThisAddIn.frmPrefs.txtPBDeviceID.Text)
                .SetValue(RegistryFolder, "EmailAddress", ThisAddIn.frmPrefs.txtEmailAddress.Text)
                .SetValue(RegistryFolder, "EmailPassword", ThisAddIn.frmPrefs.txtEmailPassword.Text)
                .SetValue(RegistryFolder, "MaxResults", ThisAddIn.frmPrefs.txtResultAmt.Text)
                .SetValue(RegistryFolder, "MinTolerableProfit", ThisAddIn.frmPrefs.txtMinProfit.Text)
                .SetValue(RegistryFolder, "City", ThisAddIn.frmPrefs.txtCity.Text)

                'radio buttons
                If ThisAddIn.frmPrefs.optPostedToday.Checked Then
                    .SetValue(RegistryFolder, "PostTimingPref", "timingPostedToday")
                ElseIf ThisAddIn.frmPrefs.optUpdatedToday.Checked Then
                    .SetValue(RegistryFolder, "PostTimingPref", "timingUpdatedToday")
                ElseIf ThisAddIn.frmPrefs.optUpdated7Days.Checked Then
                    .SetValue(RegistryFolder, "PostTimingPref", "timingUpdated7Days")
                ElseIf ThisAddIn.frmPrefs.optUpdated14Days.Checked Then
                    .SetValue(RegistryFolder, "PostTimingPref", "timingUpdated14Days")
                Else
                    .SetValue(RegistryFolder, "PostTimingPref", "timingShowAll")
                End If

            End With

        Catch exAll As Exception
            MsgBox("An error has occurred when saving the bot's preferences.  Please close and reopen this bot and try again.", vbCritical, ThisAddIn.Title)
        End Try

    End Sub
End Class