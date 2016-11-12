Imports System.Drawing
Imports System.Windows.Forms

Public Class DataValidationHelpers

    Public Shared Sub checkForNumber(theControl As Control)
        If theControl.Text = "" Or Not IsNumeric(theControl.Text) Then
            ThisAddIn.Proceed = False
            theControl.ForeColor = Color.White
            theControl.BackColor = Color.Red
        Else
            theControl.ForeColor = Color.Black
            theControl.BackColor = Color.White
        End If
        theControl.Refresh()
    End Sub

    Public Shared Sub checkForValue(theControl As Control)
        If theControl.Text = "" Then
            ThisAddIn.Proceed = False
            theControl.ForeColor = Color.White
            theControl.BackColor = Color.Red
        Else
            theControl.ForeColor = Color.Black
            theControl.BackColor = Color.White
        End If
        theControl.Refresh()
    End Sub
End Class
