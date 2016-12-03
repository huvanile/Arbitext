Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Public Shared tpnAuto As TpnAuto : Public Shared ctpAuto As Microsoft.Office.Tools.CustomTaskPane

#Region "Find Deals"

    Private Sub btnSearch_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSearch.Click
        tpnAuto = New TpnAuto
        ctpAuto = Globals.ThisAddIn.CustomTaskPanes.Add(tpnAuto, "Mass Search")
        ctpAuto.Width = 300
        ctpAuto.Control.Width = 300
        ctpAuto.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpAuto.Visible = True
    End Sub

    Private Sub btnAnalyze_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAnalyze.Click
        ThisAddIn.AppExcel.StatusBar = False
        Dim singlePostAnalysis As New SinglePostAnalysis
        singlePostAnalysis = Nothing
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
