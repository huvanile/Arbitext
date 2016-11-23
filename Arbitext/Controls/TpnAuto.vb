Public Class TpnAuto
    Delegate Sub setLabelSafeCallback([theText] As String)
    Delegate Sub visibleLblSafeCallback([theText] As String)

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If MsgBox("Quit current search?", vbYesNoCancel, ThisAddIn.Title) = vbYes Then
            If Not IsNothing(ThisAddIn.t1) Then
                ThisAddIn.t1.Abort()
                ThisAddIn.t1 = Nothing
            End If
            ThisAddIn.Proceed = False
            Ribbon1.ctpAuto.Visible = False
            Ribbon1.ctpAuto.Dispose()
        End If
    End Sub

    Private Sub btnProceed_Click(sender As Object, e As EventArgs) Handles btnProceed.Click
        Dim multiplePostsAnalysis = New MultiplePostsAnalysis
        multiplePostsAnalysis = Nothing
    End Sub

#Region "Thread-safe GUI Updaters"

    Public Sub showLblRecordSafe(theText As String)
        Try
            With lblNumber
                If .InvokeRequired Then
                    Dim d As New visibleLblSafeCallback(AddressOf showLblRecordSafe)
                    Me.Invoke(d, New Object() {[theText]})
                    d = Nothing
                Else
                    .Visible = True
                    .Refresh()
                    .Invalidate()
                End If
            End With
        Catch exAll As Exception : End Try
    End Sub

    Public Sub showLblCountsSafe(theText As String)
        Try
            With lblCounts
                If .InvokeRequired Then
                    Dim d As New visibleLblSafeCallback(AddressOf showLblCountsSafe)
                    Me.Invoke(d, New Object() {[theText]})
                    d = Nothing
                Else
                    .Visible = True
                    .Refresh()
                    .Invalidate()
                End If
            End With
        Catch exAll As Exception : End Try
    End Sub

    Public Sub hideLblCountsSafe(theText As String)
        Try
            With lblCounts
                If .InvokeRequired Then
                    Dim d As New visibleLblSafeCallback(AddressOf hideLblCountsSafe)
                    Me.Invoke(d, New Object() {[theText]})
                    d = Nothing
                Else
                    .Visible = False
                    .Refresh()
                    .Invalidate()
                End If
            End With
        Catch exAll As Exception : End Try
    End Sub

    Public Sub hideLblRecordSafe(theText As String)
        Try
            With lblNumber
                If .InvokeRequired Then
                    Dim d As New visibleLblSafeCallback(AddressOf hideLblRecordSafe)
                    Me.Invoke(d, New Object() {[theText]})
                    d = Nothing
                Else
                    .Visible = False
                    .Refresh()
                    .Invalidate()
                End If
            End With
        Catch exAll As Exception : End Try
    End Sub

    Public Sub UpdateLblCountSafe(theText As String)
        Try
            With lblCounts
                If .InvokeRequired Then
                    Dim d As New setLabelSafeCallback(AddressOf UpdateLblCountSafe)
                    Me.Invoke(d, New Object() {[theText]})
                    d = Nothing
                Else
                    .Text = theText
                    .Refresh()
                    .Invalidate()
                End If
            End With
        Catch exAll As Exception : End Try
    End Sub

    Public Sub UpdateLblNumberSafe(theText As String)
        Try
            With lblNumber
                If .InvokeRequired Then
                    Dim d As New setLabelSafeCallback(AddressOf UpdateLblNumberSafe)
                    Me.Invoke(d, New Object() {[theText]})
                    d = Nothing
                Else
                    .Text = theText
                    .Refresh()
                    .Invalidate()
                End If
            End With
        Catch exAll As Exception : End Try
    End Sub

    Public Sub UpdateLblStatusSafe(theText As String)
        Try
            With lblStatus
                If .InvokeRequired Then
                    Dim d As New setLabelSafeCallback(AddressOf UpdateLblStatusSafe)
                    Me.Invoke(d, New Object() {[theText]})
                    d = Nothing
                Else
                    .Text = theText
                    .Refresh()
                    .Invalidate()
                End If
            End With
        Catch exAll As Exception : End Try
    End Sub

    Private Sub TpnAuto_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtCity.Text = ThisAddIn.City
        txtTimingPref.Text = ThisAddIn.PostTimingPref
    End Sub

#End Region

End Class
