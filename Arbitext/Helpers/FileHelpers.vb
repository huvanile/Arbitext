Imports System.IO

Public Class FileHelpers
    Public Shared Sub WriteToFile(filePath As String, theValue As String)
        If File.Exists(filePath) Then Kill(filePath)
        File.WriteAllText(filePath, theValue)
    End Sub
End Class

