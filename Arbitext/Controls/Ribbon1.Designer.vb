﻿Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.tabBooks = Me.Factory.CreateRibbonTab
        Me.grpFinders = Me.Factory.CreateRibbonGroup
        Me.btnSearch = Me.Factory.CreateRibbonButton
        Me.btnAnalyze = Me.Factory.CreateRibbonButton
        Me.grpOther = Me.Factory.CreateRibbonGroup
        Me.mnuBuildSheets = Me.Factory.CreateRibbonMenu
        Me.btnActivityLog = Me.Factory.CreateRibbonButton
        Me.btnSetPrefs = Me.Factory.CreateRibbonButton
        Me.tabBooks.SuspendLayout()
        Me.grpFinders.SuspendLayout()
        Me.grpOther.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabBooks
        '
        Me.tabBooks.Groups.Add(Me.grpFinders)
        Me.tabBooks.Groups.Add(Me.grpOther)
        Me.tabBooks.Label = "ARBITEXT"
        Me.tabBooks.Name = "tabBooks"
        '
        'grpFinders
        '
        Me.grpFinders.Items.Add(Me.btnSearch)
        Me.grpFinders.Items.Add(Me.btnAnalyze)
        Me.grpFinders.Label = "Find Deals"
        Me.grpFinders.Name = "grpFinders"
        '
        'btnSearch
        '
        Me.btnSearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSearch.Image = CType(resources.GetObject("btnSearch.Image"), System.Drawing.Image)
        Me.btnSearch.Label = "Search for Deals"
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.ShowImage = True
        '
        'btnAnalyze
        '
        Me.btnAnalyze.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAnalyze.Image = CType(resources.GetObject("btnAnalyze.Image"), System.Drawing.Image)
        Me.btnAnalyze.Label = "Analyze Single Post"
        Me.btnAnalyze.Name = "btnAnalyze"
        Me.btnAnalyze.ShowImage = True
        '
        'grpOther
        '
        Me.grpOther.Items.Add(Me.mnuBuildSheets)
        Me.grpOther.Items.Add(Me.btnSetPrefs)
        Me.grpOther.Label = "Other Stuff"
        Me.grpOther.Name = "grpOther"
        '
        'mnuBuildSheets
        '
        Me.mnuBuildSheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.mnuBuildSheets.Image = CType(resources.GetObject("mnuBuildSheets.Image"), System.Drawing.Image)
        Me.mnuBuildSheets.Items.Add(Me.btnActivityLog)
        Me.mnuBuildSheets.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.mnuBuildSheets.Label = "Build Sheets"
        Me.mnuBuildSheets.Name = "mnuBuildSheets"
        Me.mnuBuildSheets.ShowImage = True
        '
        'btnActivityLog
        '
        Me.btnActivityLog.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnActivityLog.Description = "Build the 'Activity Log' worksheet"
        Me.btnActivityLog.Image = CType(resources.GetObject("btnActivityLog.Image"), System.Drawing.Image)
        Me.btnActivityLog.Label = "Activity Log"
        Me.btnActivityLog.Name = "btnActivityLog"
        Me.btnActivityLog.ShowImage = True
        '
        'btnSetPrefs
        '
        Me.btnSetPrefs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSetPrefs.Image = CType(resources.GetObject("btnSetPrefs.Image"), System.Drawing.Image)
        Me.btnSetPrefs.Label = "Set Prefs"
        Me.btnSetPrefs.Name = "btnSetPrefs"
        Me.btnSetPrefs.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.tabBooks)
        Me.tabBooks.ResumeLayout(False)
        Me.tabBooks.PerformLayout()
        Me.grpFinders.ResumeLayout(False)
        Me.grpFinders.PerformLayout()
        Me.grpOther.ResumeLayout(False)
        Me.grpOther.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tabBooks As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents btnSearch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAnalyze As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOther As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents mnuBuildSheets As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnActivityLog As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSetPrefs As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFinders As Microsoft.Office.Tools.Ribbon.RibbonGroup
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
