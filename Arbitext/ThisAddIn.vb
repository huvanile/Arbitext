Imports System.Threading

Public Class ThisAddIn
    Public Shared AppExcel As Excel.Application
    Public Shared Proceed As Boolean
    Public Shared AppIE As Object
    Public Shared frmPrefs As FrmPrefs
    Public Shared t1 As Thread
    Public Shared SaveAsFolder As String

    'preferences
    Public Shared SaveWBFilePath As String                 'saveFileAs: output filepath
    Public Shared SaveWBFileName As String

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        AppExcel = Me.Application
        RegistryHelpers.addressNullPrefs()
        ThisAddIn.Proceed = True
        RegistryHelpers.loadVariablesFromRegistry()
    End Sub

End Class
