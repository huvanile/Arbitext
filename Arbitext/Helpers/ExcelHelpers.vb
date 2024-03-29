﻿Imports Microsoft.Office.Interop.Excel
Imports ArbitextClassLibrary.Globals
Public Class ExcelHelpers

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

    Public Shared Sub verifyOneRowSelected()
        If ThisAddIn.AppExcel.Selection.Rows.Count > 1 Then
            MsgBox("Only 1 row must be selected", vbInformation, Title)
            ThisAddIn.Proceed = False
        End If
    End Sub

    Public Shared Sub DeleteWS(theWS As String)
        Try
            ThisAddIn.AppExcel.DisplayAlerts = False
            ThisAddIn.AppExcel.Worksheets(theWS).delete
        Catch ex As Exception

        Finally
            ThisAddIn.AppExcel.DisplayAlerts = True
        End Try
    End Sub

    Public Shared Sub deleteAllPics(Optional theWS As String = "")
        If theWS = "" Then theWS = ThisAddIn.AppExcel.ActiveSheet.Name
        Dim shp As Excel.Shape
        For Each shp In ThisAddIn.AppExcel.Sheets(theWS).Shapes
            If (shp.Type = 13) Then shp.Delete()
        Next shp
    End Sub

    Public Shared Sub thinOuterBorders(theRange As Excel.Range)
        theRange.Borders(XlBordersIndex.xlDiagonalDown).LineStyle = XlLineStyle.xlLineStyleNone
        theRange.Borders(XlBordersIndex.xlDiagonalUp).LineStyle = XlLineStyle.xlLineStyleNone
        With theRange.Borders(XlBordersIndex.xlEdgeLeft)
            .LineStyle = XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = XlBorderWeight.xlThin
        End With
        With theRange.Borders(XlBordersIndex.xlEdgeTop)
            .LineStyle = XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = XlBorderWeight.xlThin
        End With
        With theRange.Borders(XlBordersIndex.xlEdgeBottom)
            .LineStyle = XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = XlBorderWeight.xlThin
        End With
        With theRange.Borders(XlBordersIndex.xlEdgeRight)
            .LineStyle = XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = XlBorderWeight.xlThin
        End With
        theRange.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlLineStyleNone
        theRange.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlLineStyleNone
    End Sub


    Public Shared Sub setNoBorders(theRange As Range)
        With theRange
            .Borders(XlBordersIndex.xlDiagonalDown).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlDiagonalUp).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlLineStyleNone
        End With
    End Sub

    Public Shared Sub standardPageTitle(theTitle As String)
        With ThisAddIn.AppExcel.Range("A1")
            .WrapText = False
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Cambria"
            .Value = theTitle
        End With
    End Sub

    Public Shared Sub standardColumnTitles(theRange As Range)
        thinInnerBorder(theRange)
        With theRange
            .Font.Bold = True
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .Interior.Color = 11851260
            .WrapText = True
        End With
    End Sub

    Public Shared Sub thinInnerBorder(theRange As Range)
        With theRange
            With .Borders(XlBordersIndex.xlEdgeLeft)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            With .Borders(XlBordersIndex.xlEdgeTop)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            With .Borders(XlBordersIndex.xlEdgeBottom)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            With .Borders(XlBordersIndex.xlEdgeRight)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            With .Borders(XlBordersIndex.xlInsideVertical)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            With .Borders(XlBordersIndex.xlInsideHorizontal)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
        End With
    End Sub

    'check to see if any workbook is open.  returns boolean variable.
    Public Shared Function isAnyWBOpen() As Boolean
        Try
            Dim check As String = ThisAddIn.AppExcel.ActiveSheet.Name
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Sub createWS(theName As String)
        Try
            If Not isAnyWBOpen() Then
                ThisAddIn.AppExcel.Workbooks.Add()
            End If
            ThisAddIn.AppExcel.DisplayAlerts = False
            ThisAddIn.AppExcel.Sheets.Add.Name = theName
        Catch ex As Exception

        Finally
            ThisAddIn.AppExcel.DisplayAlerts = True
        End Try
    End Sub

    Public Shared Function doesWSExist(ws As String)
        doesWSExist = False
        Try
            Diagnostics.Debug.Print(ThisAddIn.AppExcel.Sheets(ws).name)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function columnLetter(ColumnNumber As Integer) As String
        'converts column number to it's letter equivalent
        If ColumnNumber <= 0 Then 'negative column number
            columnLetter = ""
        ElseIf ColumnNumber > 16384 Then 'column not supported (too big) in Excel 2007
            columnLetter = ""
        ElseIf ColumnNumber > 702 Then ' triple letter columns
            columnLetter =
            Chr((Int((ColumnNumber - 1 - 26 - 676) / 676)) Mod 676 + 65) &
            Chr((Int((ColumnNumber - 1 - 26) / 26) Mod 26) + 65) &
            Chr(((ColumnNumber - 1) Mod 26) + 65)
        ElseIf ColumnNumber > 26 Then
            columnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & Chr(((ColumnNumber - 1) Mod 26) + 65)  ' double letter columns
        Else
            columnLetter = Chr(ColumnNumber + 64) ' single letter columns
        End If
    End Function

    Public Shared Function canFind(theValue As String, Optional theWS As String = "", Optional theWB As String = "", Optional returnAddress As Boolean = False, Optional searchPart As Boolean = False) As String
        Try
            Dim tmp As String
            If theWS = "" Then theWS = ThisAddIn.AppExcel.ActiveSheet.Name
            If theWB = "" Then theWB = ThisAddIn.AppExcel.ActiveWorkbook.Name
            If searchPart Then
                tmp = ThisAddIn.AppExcel.Workbooks(theWB).Sheets(theWS).Cells.Find(What:=theValue, LookAt:=XlLookAt.xlPart).Address
            Else
                tmp = ThisAddIn.AppExcel.Workbooks(theWB).Sheets(theWS).Cells.Find(What:=theValue, LookAt:=XlLookAt.xlWhole).Address
            End If
            canFind = True
            If returnAddress Then canFind = tmp
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function canFindInResultCol(theCol As String, theValue As String, Optional theWS As String = "") As String
        Try
            If theWS = "" Then theWS = ThisAddIn.AppExcel.ActiveSheet.Name
            If ThisAddIn.AppExcel.Range(theCol & "4").Value2 = "" Then
                Return False
            Else
                Dim plur = ThisAddIn.AppExcel.Range(theCol & "3").End(XlDirection.xlDown).Row
                Dim tmp As String = ThisAddIn.AppExcel.Sheets(theWS).range(theCol & "4:" & theCol & plur).Find(What:=theValue, LookAt:=XlLookAt.xlWhole).Address
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function lastUsedRow(Optional sheetName As String = "")
        Dim emptyColCount : emptyColCount = 0
        Dim c As Integer : c = 1
        lastUsedRow = 0
        If sheetName = "" Then sheetName = ThisAddIn.AppExcel.ActiveSheet.Name
        Do Until c = 16384 Or emptyColCount > 10
            If ThisAddIn.AppExcel.WorksheetFunction.CountA(ThisAddIn.AppExcel.Sheets(sheetName).Columns(columnLetter(c))) <> 0 Then
                emptyColCount = 0
                If ThisAddIn.AppExcel.Sheets(sheetName).Range(columnLetter(c) & "564000").End(XlDirection.xlUp).Row > lastUsedRow Then
                    lastUsedRow = ThisAddIn.AppExcel.Sheets(sheetName).Range(columnLetter(c) & "564000").End(XlDirection.xlUp).Row
                End If
            Else
                emptyColCount = emptyColCount + 1
            End If
            c = c + 1
        Loop
    End Function

    ''' <summary>
    ''' Converts a column letter to a number
    ''' </summary>
    ''' <param name="InputLetter">The column letter</param>
    ''' <returns>A column number</returns>
    Public Shared Function ColumnNumber(InputLetter As String) As Integer
        Dim Leng As Integer
        Dim i As Integer
        Dim tmp As Integer
        Leng = Len(InputLetter)
        For i = 1 To Leng
            tmp = (Asc(UCase(Mid(InputLetter, i, 1))) - 64) + tmp * 26
        Next i
        ColumnNumber = tmp
    End Function
End Class
