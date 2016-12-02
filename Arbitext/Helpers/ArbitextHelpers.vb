Imports Arbitext.ExcelHelpers
Imports Microsoft.Office.Interop.Excel

Public Class ArbitextHelpers

    Public Shared Function bookCountFromString(theStr As String) As Integer
        Dim tmpCount As Integer = 0
        Dim tmp1 As Integer
        Dim tmp2 As Integer
        Dim tmp3 As Integer
        tmp1 = UBound(Split(LCase(theStr), "isbn"))
        tmp2 = UBound(Split(LCase(theStr), "978"))
        tmp3 = UBound(Split(LCase(theStr), "$"))
        tmpCount = tmp1
        If tmp2 > tmpCount Then tmpCount = tmp2
        If tmp3 > tmpCount Then tmpCount = tmp3
        Return tmpCount
    End Function

    Public Shared Function randomUID() As String
        randomUID = ""
        Dim rgch As String
        rgch = "abcdefgh"
        rgch = rgch & "0123456789"
        Do Until Len(randomUID) = 16
            randomUID = randomUID & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Loop
    End Function

    Public Shared Sub grabPHPicIfNeeded()
        If Not System.IO.File.Exists(Environ("Temp") & "\placeholder.jpg") Then
            Dim wc As New System.Net.WebClient
            wc.DownloadFile("http://datadump.thelightningpath.com/images/bookimages/PlaceholderBook.png", Environ("Temp") & "\placeholder.jpg")
            wc = Nothing
        End If
    End Sub

    Public Shared Sub rowTitles(theRange As Excel.Range)
        With theRange
            .HorizontalAlignment = XlHAlign.xlHAlignRight
            .Interior.ColorIndex = 15
        End With
        thinInnerBorder(theRange)
    End Sub

    Public Shared Sub rowValues(theRange As Excel.Range)
        theRange.Interior.ColorIndex = 6
        thinInnerBorder(theRange)
    End Sub

    Public Shared Sub unFilterTrash()
        Try
            ThisAddIn.AppExcel.Worksheets("Trash").AutoFilter.Sort.SortFields.Clear
            ThisAddIn.AppExcel.Worksheets("Trash").ShowAllData
        Catch ex As Exception : End Try
    End Sub

End Class
