Imports System.Text
Imports mshtml

Public Class TPGResult
    Private _url As String
    Private _html As String
    Private _median As Decimal = -1
    Private _min As Decimal = -1
    Private _max As Decimal = -1
    Private _mean As Decimal = -1
    Private _stdDev As Decimal = -1

    Sub New(searchTerm As String)
        Dim splitholder = Split(searchTerm, " ")
        Dim tmp As String = "http://www.thepricegeek.com/results/"
        For x = LBound(splitholder) To UBound(splitholder)
            If x = 0 Then tmp = tmp & splitholder(x) Else tmp = tmp & "+" & splitholder(x)
        Next
        _url = tmp & "?country=us"
        tmp = ""
        Dim wc As New Net.WebClient
        Dim bHTML() As Byte = wc.DownloadData(_url)
        _html = New UTF8Encoding().GetString(bHTML)
    End Sub

    Public ReadOnly Property URL As String
        Get
            Return _url
        End Get
    End Property

    Public ReadOnly Property HTML As String
        Get
            Return _html
        End Get
    End Property

    Public ReadOnly Property Median As Decimal
        Get
            Dim z As Long : z = 0
            Dim splitholder
            Dim m As String : m = ""
            z = Strings.InStr(_html, "<em class=""median"">")
            If z = 0 Then Throw New Exception
            m = Right(_html, Len(_html) - (z + Len("<em class=""median"">") - 1))
            splitholder = Split(m, "</em>") 'boundary
            m = splitholder(0).trim
            Return CDec(m)
        End Get
    End Property

    Public ReadOnly Property Min As Decimal
        Get
            Dim z As Long : z = 0
            Dim splitholder
            Dim m As String : m = ""
            z = Strings.InStr(_html, "<span class=""min"">")
            If z = 0 Then Throw New Exception
            m = Right(_html, Len(_html) - (z + Len("<span class=""min"">") - 1))
            splitholder = Split(m, "</span>") 'boundary
            m = splitholder(0).trim
            Return CDec(m)
        End Get
    End Property

    Public ReadOnly Property Max As Decimal
        Get
            Dim z As Long : z = 0
            Dim splitholder
            Dim m As String : m = ""
            z = Strings.InStr(_html, "<span class=""max"">")
            If z = 0 Then Throw New Exception
            m = Right(_html, Len(_html) - (z + Len("<span class=""max"">") - 1))
            splitholder = Split(m, "</span>") 'boundary
            m = splitholder(0).trim
            If m Like "* *" Then
                splitholder = Split(m, " ")
                m = splitholder(UBound(splitholder)).trim
            End If
            Return CDec(m)
        End Get
    End Property

    Public ReadOnly Property Mean As Decimal
        Get
            Dim z As Long : z = 0
            Dim splitholder
            Dim m As String : m = ""
            z = Strings.InStr(_html, "Mean:</strong> $")
            If z = 0 Then Throw New Exception
            m = Right(_html, Len(_html) - (z + Len("Mean:</strong> $") - 1))
            splitholder = Split(m, "</div>") 'boundary
            m = splitholder(0).trim
            Return CDec(m)
        End Get
    End Property

    Public ReadOnly Property StdDev As Decimal
        Get
            Dim z As Long : z = 0
            Dim splitholder
            Dim m As String : m = ""
            z = Strings.InStr(1, _html, "Std Dev:</strong> $")
            If z = 0 Then Throw New Exception
            m = Right(_html, Len(_html) - (z + Len("Std Dev:</strong> $") - 1))
            splitholder = Split(m, "</div>") 'boundary
            m = splitholder(0).trim
            Return CDec(m)
        End Get
    End Property
End Class
