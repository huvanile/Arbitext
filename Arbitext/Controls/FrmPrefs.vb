Imports System.Net
Imports ArbitextClassLibrary.Globals
Imports Arbitext.RegistryHelpers

Public Class FrmPrefs

    Private Sub FrmPrefs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        addressNullPrefs()
        loadVariablesFromRegistry()
        loadFormPrefsFromVariables()
        Me.Text = Title
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        saveFormPrefsToRegistry()
        loadVariablesFromRegistry()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
        ThisAddIn.frmPrefs = Nothing
    End Sub


End Class