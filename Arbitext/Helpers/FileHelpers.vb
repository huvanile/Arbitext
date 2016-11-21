Imports System.IO

Public Class FileHelpers
    Public Shared Sub WriteToFile(filePath As String, theValue As String)
        If File.Exists(filePath) = False Then
            File.WriteAllText(filePath, theValue)
        End If
    End Sub
End Class

