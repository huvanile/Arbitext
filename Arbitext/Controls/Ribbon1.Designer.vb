Partial Class Ribbon1
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
        Me.grpFind = Me.Factory.CreateRibbonGroup
        Me.btnKeeper = Me.Factory.CreateRibbonButton
        Me.btnMaybe = Me.Factory.CreateRibbonButton
        Me.btnTrash = Me.Factory.CreateRibbonButton
        Me.grpTrash = Me.Factory.CreateRibbonGroup
        Me.btnTrashBadDeals = Me.Factory.CreateRibbonButton
        Me.btnEmptyTrash = Me.Factory.CreateRibbonButton
        Me.grpOther = Me.Factory.CreateRibbonGroup
        Me.mnuBuildSheets = Me.Factory.CreateRibbonMenu
        Me.btnAutomatedChecks = Me.Factory.CreateRibbonButton
        Me.btnSingleCheck = Me.Factory.CreateRibbonButton
        Me.btnMultipostManual = Me.Factory.CreateRibbonButton
        Me.btnKeepers = Me.Factory.CreateRibbonButton
        Me.btnColorLegend = Me.Factory.CreateRibbonButton
        Me.btnActivityLog = Me.Factory.CreateRibbonButton
        Me.btnBuildTrash = Me.Factory.CreateRibbonButton
        Me.btnSetPrefs = Me.Factory.CreateRibbonButton
        Me.btnBuildWSMaybes = Me.Factory.CreateRibbonButton
        Me.tabBooks.SuspendLayout()
        Me.grpFinders.SuspendLayout()
        Me.grpFind.SuspendLayout()
        Me.grpTrash.SuspendLayout()
        Me.grpOther.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabBooks
        '
        Me.tabBooks.Groups.Add(Me.grpFinders)
        Me.tabBooks.Groups.Add(Me.grpFind)
        Me.tabBooks.Groups.Add(Me.grpTrash)
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
        Me.btnAnalyze.Label = "Analyze Deal"
        Me.btnAnalyze.Name = "btnAnalyze"
        Me.btnAnalyze.ShowImage = True
        '
        'grpFind
        '
        Me.grpFind.Items.Add(Me.btnKeeper)
        Me.grpFind.Items.Add(Me.btnMaybe)
        Me.grpFind.Items.Add(Me.btnTrash)
        Me.grpFind.Label = "Categorize Deals"
        Me.grpFind.Name = "grpFind"
        '
        'btnKeeper
        '
        Me.btnKeeper.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnKeeper.Image = CType(resources.GetObject("btnKeeper.Image"), System.Drawing.Image)
        Me.btnKeeper.Label = "It's a Keeper"
        Me.btnKeeper.Name = "btnKeeper"
        Me.btnKeeper.ShowImage = True
        '
        'btnMaybe
        '
        Me.btnMaybe.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnMaybe.Image = CType(resources.GetObject("btnMaybe.Image"), System.Drawing.Image)
        Me.btnMaybe.Label = "It's a Maybe"
        Me.btnMaybe.Name = "btnMaybe"
        Me.btnMaybe.ShowImage = True
        '
        'btnTrash
        '
        Me.btnTrash.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnTrash.Image = CType(resources.GetObject("btnTrash.Image"), System.Drawing.Image)
        Me.btnTrash.Label = "It's Trash"
        Me.btnTrash.Name = "btnTrash"
        Me.btnTrash.ShowImage = True
        '
        'grpTrash
        '
        Me.grpTrash.Items.Add(Me.btnTrashBadDeals)
        Me.grpTrash.Items.Add(Me.btnEmptyTrash)
        Me.grpTrash.Label = "Trash Stuff"
        Me.grpTrash.Name = "grpTrash"
        '
        'btnTrashBadDeals
        '
        Me.btnTrashBadDeals.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnTrashBadDeals.Image = CType(resources.GetObject("btnTrashBadDeals.Image"), System.Drawing.Image)
        Me.btnTrashBadDeals.Label = "Trash Bad Deals"
        Me.btnTrashBadDeals.Name = "btnTrashBadDeals"
        Me.btnTrashBadDeals.ShowImage = True
        '
        'btnEmptyTrash
        '
        Me.btnEmptyTrash.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnEmptyTrash.Image = CType(resources.GetObject("btnEmptyTrash.Image"), System.Drawing.Image)
        Me.btnEmptyTrash.Label = "Empty Trash"
        Me.btnEmptyTrash.Name = "btnEmptyTrash"
        Me.btnEmptyTrash.ShowImage = True
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
        Me.mnuBuildSheets.Items.Add(Me.btnAutomatedChecks)
        Me.mnuBuildSheets.Items.Add(Me.btnSingleCheck)
        Me.mnuBuildSheets.Items.Add(Me.btnMultipostManual)
        Me.mnuBuildSheets.Items.Add(Me.btnKeepers)
        Me.mnuBuildSheets.Items.Add(Me.btnBuildTrash)
        Me.mnuBuildSheets.Items.Add(Me.btnBuildWSMaybes)
        Me.mnuBuildSheets.Items.Add(Me.btnColorLegend)
        Me.mnuBuildSheets.Items.Add(Me.btnActivityLog)
        Me.mnuBuildSheets.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.mnuBuildSheets.Label = "Build Sheets"
        Me.mnuBuildSheets.Name = "mnuBuildSheets"
        Me.mnuBuildSheets.ShowImage = True
        '
        'btnAutomatedChecks
        '
        Me.btnAutomatedChecks.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAutomatedChecks.Description = "Build the 'Automated Checks' worksheet"
        Me.btnAutomatedChecks.Image = CType(resources.GetObject("btnAutomatedChecks.Image"), System.Drawing.Image)
        Me.btnAutomatedChecks.Label = "Automated Checks"
        Me.btnAutomatedChecks.Name = "btnAutomatedChecks"
        Me.btnAutomatedChecks.ShowImage = True
        '
        'btnSingleCheck
        '
        Me.btnSingleCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSingleCheck.Description = "Build the 'Single Check' worksheet"
        Me.btnSingleCheck.Image = CType(resources.GetObject("btnSingleCheck.Image"), System.Drawing.Image)
        Me.btnSingleCheck.Label = "Single Checks"
        Me.btnSingleCheck.Name = "btnSingleCheck"
        Me.btnSingleCheck.ShowImage = True
        '
        'btnMultipostManual
        '
        Me.btnMultipostManual.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnMultipostManual.Description = "Build the 'Multipost Manual Check' worksheet"
        Me.btnMultipostManual.Image = CType(resources.GetObject("btnMultipostManual.Image"), System.Drawing.Image)
        Me.btnMultipostManual.Label = "Multipost Manual Check"
        Me.btnMultipostManual.Name = "btnMultipostManual"
        Me.btnMultipostManual.ShowImage = True
        '
        'btnKeepers
        '
        Me.btnKeepers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnKeepers.Description = "Build the 'Keepers' worksheet"
        Me.btnKeepers.Image = CType(resources.GetObject("btnKeepers.Image"), System.Drawing.Image)
        Me.btnKeepers.Label = "Keepers"
        Me.btnKeepers.Name = "btnKeepers"
        Me.btnKeepers.ShowImage = True
        '
        'btnColorLegend
        '
        Me.btnColorLegend.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnColorLegend.Description = "Build the 'Color Legend' worksheet"
        Me.btnColorLegend.Image = CType(resources.GetObject("btnColorLegend.Image"), System.Drawing.Image)
        Me.btnColorLegend.Label = "Color Legend"
        Me.btnColorLegend.Name = "btnColorLegend"
        Me.btnColorLegend.ShowImage = True
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
        'btnBuildTrash
        '
        Me.btnBuildTrash.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnBuildTrash.Description = "Build the 'Trash' worksheet"
        Me.btnBuildTrash.Image = CType(resources.GetObject("btnBuildTrash.Image"), System.Drawing.Image)
        Me.btnBuildTrash.Label = "Trash"
        Me.btnBuildTrash.Name = "btnBuildTrash"
        Me.btnBuildTrash.ShowImage = True
        '
        'btnSetPrefs
        '
        Me.btnSetPrefs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSetPrefs.Image = CType(resources.GetObject("btnSetPrefs.Image"), System.Drawing.Image)
        Me.btnSetPrefs.Label = "Set Prefs"
        Me.btnSetPrefs.Name = "btnSetPrefs"
        Me.btnSetPrefs.ShowImage = True
        '
        'btnBuildWSMaybes
        '
        Me.btnBuildWSMaybes.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnBuildWSMaybes.Description = "Build the 'Maybes' worksheet"
        Me.btnBuildWSMaybes.Image = CType(resources.GetObject("btnBuildWSMaybes.Image"), System.Drawing.Image)
        Me.btnBuildWSMaybes.Label = "Maybes"
        Me.btnBuildWSMaybes.Name = "btnBuildWSMaybes"
        Me.btnBuildWSMaybes.ShowImage = True
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
        Me.grpFind.ResumeLayout(False)
        Me.grpFind.PerformLayout()
        Me.grpTrash.ResumeLayout(False)
        Me.grpTrash.PerformLayout()
        Me.grpOther.ResumeLayout(False)
        Me.grpOther.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tabBooks As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpFind As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnSearch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAnalyze As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnKeeper As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnMaybe As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnTrash As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTrash As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnTrashBadDeals As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnEmptyTrash As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOther As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents mnuBuildSheets As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnAutomatedChecks As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSingleCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnMultipostManual As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnKeepers As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnColorLegend As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnActivityLog As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnBuildTrash As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSetPrefs As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFinders As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnBuildWSMaybes As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
