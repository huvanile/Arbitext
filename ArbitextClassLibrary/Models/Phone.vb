Imports System.Text

Public Class Phone
    Private _title As String
    Private _imageURL As String
    Private _isParsable As Boolean
    Private _condition As String
    Private _hasSiblings As Boolean         'is the book part of a multipost?
    Private _parentTitle As String          'title of the parent post
    Private _parentBody As String           'html body of the parent post
    Private _parentPostDate As String       'the date of the original / parent post
    Private _median As Decimal
    Private _min As Decimal
    Private _max As Decimal
    Private _mean As Decimal
    Private _stdDev As Decimal
    Private _askingPrice As Decimal
    Private _searchterm As String
    Private _saleDescInPost As String       'the part of the post used to describe this phone. for single-phone posts, this is teh whole post body.

#Region "Contructors"

    Sub New(searchterm As String, askingPrice As Decimal, theSaleDesc As String, queryTPG As Boolean, hasSiblings As Boolean, parentTitle As String, parentBody As String, parentPostDate As String)
        _searchterm = searchterm.Trim
        _askingPrice = askingPrice
        _hasSiblings = hasSiblings
        _parentBody = parentBody
        _parentTitle = parentTitle
        _parentPostDate = parentPostDate
        _saleDescInPost = theSaleDesc.Trim
        _askingPrice = askingPrice
        If GetDataFromThePriceGeek() Then _isParsable = True Else _isParsable = False
    End Sub

#End Region

#Region "Read & Write Properties"

    Property AskingPrice As Decimal
        Get
            Return _askingPrice
        End Get
        Set(value As Decimal)
            _askingPrice = value
        End Set
    End Property

#End Region

#Region "Readonly Properties"

    ReadOnly Property TpgURL As String
        Get
            Dim splitholder = Split(_searchterm, " ")
            Dim tmp As String = "http://www.thepricegeek.com/results/"
            For x = LBound(splitholder) To UBound(splitholder)
                If x = 0 Then tmp = tmp & splitholder(x) Else tmp = tmp & "+" & splitholder(x)
            Next
            Return tmp & "?country=us"
        End Get
    End Property
    ''' <summary>
    ''' Unique phone lead ID
    ''' </summary>
    ''' <returns>Unique phone lead ID based on post date, make, model, and asking price</returns>
    ReadOnly Property ID As String
        Get
            Dim tmp As New StringBuilder
            tmp.Append(Year(_parentPostDate) & Month(_parentPostDate) & Day(_parentPostDate) & Hour(_parentPostDate) & Minute(_parentPostDate))
            tmp.Append(Len(Trim(_searchterm)))
            tmp.Append(Trim(AskingPrice))
            Return tmp.ToString
        End Get
    End Property

    ReadOnly Property AmazonSearchURL As String
        Get
            Return "https://href.li/?http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & StringHelpers.replacePlusWithSpace(_title)
        End Get
    End Property

    ReadOnly Property IsParsable As Boolean
        Get
            If _mean = -1 Or _median = -1 Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    ReadOnly Property IsTrash() As Boolean
        Get
            IsTrash = False
            If IsParsable Then
                If Profit <= 0 Then IsTrash = True
            End If
        End Get
    End Property

    ReadOnly Property IsWinner() As Boolean
        Get
            IsWinner = False
            If IsParsable Then
                If Profit > 15 And Not IsTrash() Then IsWinner = True
            End If
        End Get
    End Property

    ReadOnly Property IsMaybe() As Boolean
        Get
            IsMaybe = False
            If IsParsable Then
                If Not IsWinner() And Not IsTrash() Then IsMaybe = True
            End If
        End Get
    End Property

    ReadOnly Property IsHVOBO As Boolean
        Get
            IsHVOBO = False
            If IsParsable _
            AndAlso Not IsWinner _
            AndAlso Not IsMaybe _
            AndAlso Not IsHVSP _
            AndAlso (isOBO Or _askingPrice = -2) _
            AndAlso _median >= 35 Then
                Return True
            End If
        End Get
    End Property

    ReadOnly Property IsHVSP As Boolean
        Get
            IsHVSP = False
            If IsParsable Then
                If Not IsWinner AndAlso Not IsMaybe Then
                    If _median > 40 AndAlso DateDiff(DateInterval.Day, CDate(_parentPostDate), Now()) >= 14 Then
                        IsHVSP = True
                    End If
                End If
            End If
        End Get
    End Property

    ReadOnly Property SaleDescInPost As String
        Get
            Return _saleDescInPost
        End Get
    End Property

    ReadOnly Property Profit As Decimal
        Get
            Dim tmp As Decimal = 0
            If _median > 0 Then
                tmp = Math.Round(_median - _askingPrice, 2)
                If tmp > 0 Then Return tmp Else Return 0
            Else
                Return 0
            End If
        End Get
    End Property

    ReadOnly Property ProfitPercentage As Decimal
        Get
            If _median > 0 Then
                Return Profit / _askingPrice
            Else
                Return 0
            End If
        End Get
    End Property

    ReadOnly Property ImageURL As String
        Get
            Return "https://encrypted-tbn1.gstatic.com/images?q=tbn:ANd9GcRSMCkoA3IEaxwUHRf4k5GdLK1UafoP92D0V5QEyjyXJWC_SJYW8w"
        End Get
    End Property

    ReadOnly Property Median As Decimal
        Get
            Return _median
        End Get
    End Property

    ReadOnly Property Mean As Decimal
        Get
            Return _mean
        End Get
    End Property

    ReadOnly Property Min As Decimal
        Get
            Return _min
        End Get
    End Property

    ReadOnly Property Max As Decimal
        Get
            Return _max
        End Get
    End Property

    ReadOnly Property StdDev As Decimal
        Get
            Return _stdDev
        End Get
    End Property

    ReadOnly Property SearchTerm As String
        Get
            Return _searchterm
        End Get
    End Property
#End Region

#Region "FLAG Properties"

    'ReadOnly Property isUnlocked() As Boolean

    'End Property

    'ReadOnly Property isNIB() As Boolean

    'End Property

    ReadOnly Property isOBO() As Boolean
        Get
            If _hasSiblings Then
                If CraigslistHelpers.isOBO(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.isOBO(_parentTitle) Or CraigslistHelpers.isOBO(_parentBody) Then Return True Else Return False
            End If
        End Get
    End Property

#End Region

#Region "Methods"

    Public Function GetDataFromThePriceGeek() As Boolean
        Try
            Dim tpgResult As New TPGResult(_searchterm)
            _mean = tpgResult.Mean
            _stdDev = tpgResult.StdDev
            _median = tpgResult.Median
            _max = tpgResult.Max
            _min = tpgResult.Min
            tpgResult = Nothing
            Return True
        Catch ex As Exception
            _mean = -1
            _median = -1
            _max = -1
            _min = -1
            _stdDev = -1
            Return False
        End Try
    End Function

#End Region

End Class

