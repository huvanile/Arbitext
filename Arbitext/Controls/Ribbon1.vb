Imports Microsoft.Office.Tools.Ribbon
Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers

Public Class Ribbon1

#Region "Find Deals"

    Private Sub btnSearch_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSearch.Click
        ThisAddIn.AppExcel.StatusBar = False
        Dim multiplePostsAnalysis As New MultiplePostsAnalysis
        multiplePostsAnalysis = Nothing
    End Sub

    Private Sub btnAnalyze_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAnalyze.Click
        ThisAddIn.AppExcel.StatusBar = False
        Dim singlePostAnalysis As New SinglePostAnalysis
        singlePostAnalysis = Nothing
    End Sub

#End Region

#Region "Trash Stuff"
    Private Sub btnEmptyTrash_Click(sender As Object, e As RibbonControlEventArgs) Handles btnEmptyTrash.Click
        ThisAddIn.AppExcel.StatusBar = False
        ThisAddIn.Proceed = True
        trashWSCheck()
        If ThisAddIn.Proceed Then
            If MsgBox("Empty the trash: Are you sure?", vbYesNoCancel, ThisAddIn.Title) = vbYes Then ThisAddIn.AppExcel.Sheets("Trash").Range("A2:zz" & lastUsedRow()).ClearContents
        End If
    End Sub

    Private Sub btnTrashBadDeals_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTrashBadDeals.Click
        ThisAddIn.AppExcel.StatusBar = False
        ThisAddIn.Proceed = True
        automatedChecksPageCheck()
        If ThisAddIn.Proceed Then
            ThisAddIn.AppExcel.ScreenUpdating = False
            Dim r As Integer
            For r = lastUsedRow("Automated Checks") To 4 Step -1
                With ThisAddIn.AppExcel.Sheets("Automated Checks")
                    If IsNumeric(.Range("i" & r).Value) Then
                        If .Range("i" & r).Value <= 0 Then
                            categorizeRow(r)
                        End If
                    End If
                End With
            Next r
            ThisAddIn.AppExcel.ScreenUpdating = True
            ThisAddIn.AppExcel.Goto(ThisAddIn.AppExcel.Range("A1"), True)
        End If
    End Sub
#End Region

#Region "Deal Categorization"
    Private Sub btnKeeper_Click(sender As Object, e As RibbonControlEventArgs) Handles btnKeeper.Click
        ThisAddIn.AppExcel.StatusBar = False
        With ThisAddIn.AppExcel
            Dim callingWS As String = .ActiveSheet.Name
            Dim x As Integer
            If Not doesWSExist("Keepers") Then BuildWSResults.buildResultWS("Keepers")
            .Sheets(callingWS).Activate
            Select Case .ActiveSheet.Name
                Case "Automated Checks"
                    For x = .Selection.Rows.Count To 1 Step -1
                        If .Selection(x).Row > 3 Then
                            categorizeRow(.Selection(x).Row, "Keepers")
                        End If
                    Next x
                Case "Maybes", "Trash"
                    ThisAddIn.Proceed = True
                    verifyOneRowSelected()
                    If ThisAddIn.Proceed Then reCategorizeRow(.Selection.Row, "Keepers")
                Case Else
                    MsgBox("This tool can't be run from the '" & .ActiveSheet.Name & "' sheet.", vbCritical, ThisAddIn.Title)
            End Select
        End With
    End Sub

    Private Sub btnMaybe_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMaybe.Click
        ThisAddIn.AppExcel.StatusBar = False
        With ThisAddIn.AppExcel
            Dim callingWS As String = .ActiveSheet.Name
            Dim x As Integer
            If Not doesWSExist("Maybes") Then BuildWSResults.buildResultWS("Maybes")
            .Sheets(callingWS).Activate
            Select Case .ActiveSheet.Name
                Case "Automated Checks"
                    For x = .Selection.Rows.Count To 1 Step -1
                        If .Selection(x).Row > 3 Then
                            categorizeRow(.Selection(x).Row, "Maybes")
                        End If
                    Next x
                Case "Keepers", "Trash"
                    ThisAddIn.Proceed = True
                    If ThisAddIn.Proceed Then verifyOneRowSelected()
                    reCategorizeRow(.Selection.Row, "Maybes")
                Case Else
                    MsgBox("This tool can't be run from the '" & .ActiveSheet.Name & "' sheet.", vbCritical, ThisAddIn.Title)
            End Select
        End With
    End Sub

    Private Sub btnTrash_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTrash.Click
        ThisAddIn.AppExcel.StatusBar = False
        With ThisAddIn.AppExcel
            Dim callingWS As String = .ActiveSheet.Name
            Dim x As Integer
            If Not doesWSExist("Trash") Then BuildWSResults.buildResultWS("Trash")
            .Sheets(callingWS).Activate
            Select Case .ActiveSheet.Name
                Case "Automated Checks"
                    For x = .Selection.Rows.Count To 1 Step -1
                        If .Selection(x).Row > 3 Then
                            categorizeRow(.Selection(x).Row)
                        End If
                    Next x
                Case "Keepers", "Maybes"
                    ThisAddIn.Proceed = True
                    verifyOneRowSelected()
                    If ThisAddIn.Proceed Then reCategorizeRow(.Selection.Row, "Trash")
                Case Else
                    MsgBox("This tool can't be run from the '" & .ActiveSheet.Name & "' sheet.", vbCritical, ThisAddIn.Title)
            End Select
        End With
    End Sub

#End Region

#Region "Sheet Builders"

    Private Sub btnActivityLog_Click(sender As Object, e As RibbonControlEventArgs) Handles btnActivityLog.Click
        ThisAddIn.AppExcel.StatusBar = False
        BuildWSActivityLog.buildWSActivityLog()
    End Sub

#End Region

    Private Sub btnSetPrefs_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSetPrefs.Click
        ThisAddIn.AppExcel.StatusBar = False
        ThisAddIn.frmPrefs = New FrmPrefs
        ThisAddIn.frmPrefs.Show()
    End Sub

End Class
