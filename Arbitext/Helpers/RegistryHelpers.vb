Imports Microsoft.Win32
Imports ArbitextClassLibrary.Globals

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

            If My.Computer.Registry.GetValue(RegistryFolder, "MinTolerableProfit", Nothing) Is Nothing Then _
                My.Computer.Registry.SetValue(RegistryFolder, "MinTolerableProfit", 15)
            If My.Computer.Registry.GetValue(RegistryFolder, "City", Nothing) Is Nothing _
                Or My.Computer.Registry.GetValue(RegistryFolder, "City", Nothing) = "" _
                Then My.Computer.Registry.SetValue(RegistryFolder, "City", "Dallas")

        Catch exAll As Exception
            MsgBox("An error has occurred when pulling in the Arbitext add-in's default preferences.  Error message:" & vbCrLf & vbCrLf & exAll.Message, vbCritical, Title)
        End Try
    End Sub

    ''' <summary>
    ''' Used to populate the global / shared variables from their stored registry values
    ''' </summary>
    Public Shared Sub loadVariablesFromRegistry()
        Try

            City = My.Computer.Registry.GetValue(RegistryFolder, "City", Nothing)

        Catch exAll As Exception
            MsgBox("An error has occurred when getting the bot's preferences.  Please close and reopen this bot and try again.", vbCritical, Title)
        End Try
    End Sub

    ''' <summary>
    ''' Used to retrieve user preferences from their varaibles upon prefs form initialize
    ''' </summary>
    Public Shared Sub loadFormPrefsFromVariables()
        Try

            'textboxes
            ThisAddIn.frmPrefs.txtCity.Text = City

        Catch exAll As Exception
            MsgBox("An error has occurred when loading the bot's preferences.  Please close and reopen this bot and try again.", vbCritical, Title)
        End Try
    End Sub

    ''' <summary>
    ''' Used to save the preferences as set in frmPrefs into the registry
    ''' </summary>
    Public Shared Sub saveFormPrefsToRegistry()
        Try

            With My.Computer.Registry

                .SetValue(RegistryFolder, "MinTolerableProfit", ThisAddIn.frmPrefs.txtMinProfit.Text)
                .SetValue(RegistryFolder, "City", ThisAddIn.frmPrefs.txtCity.Text)

            End With

        Catch exAll As Exception
            MsgBox("An error has occurred when saving the bot's preferences.  Please close and reopen this bot and try again.", vbCritical, Title)
        End Try

    End Sub
End Class