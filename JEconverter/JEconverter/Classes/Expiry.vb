Public Class Expiry

    Private _ticket As String = String.Empty
    Private _pawner As String = String.Empty
    Private _itemType As String = String.Empty
    Private _description() As String
    Private _phone As String = String.Empty
    Private _isTampered As Integer = -1 '-1 No Verified 0 Not Tampered 1 Tampered
    Private _expiry As Date
    Private _auction As Date
    Private _branch As String = String.Empty

#Region "Properties"
    Public Property TicketNumber() As String
        Get
            Return _ticket
        End Get
        Set(ByVal value As String)
            _ticket = value
        End Set
    End Property

    Public Property PawnerName As String
        Get
            Return _pawner
        End Get
        Set(ByVal value As String)
            _pawner = value
        End Set
    End Property

    Public Property ItemType As String
        Get
            Return _itemType
        End Get
        Set(ByVal value As String)
            _itemType = value
        End Set
    End Property

    Public Property Description() As String()
        Get
            Return _description
        End Get
        Set(ByVal value As String())
            _description = value
        End Set
    End Property

    Public Property PhoneNumber As String
        Get
            Return _phone
        End Get
        Set(ByVal value As String)
            _phone = value
        End Set
    End Property

    Public Property isTampered As Integer
        Get
            Return _isTampered
        End Get
        Set(ByVal value As Integer)
            _isTampered = value
        End Set
    End Property

    Public Property ExpiryDate As Date
        Get
            Return _expiry
        End Get
        Set(ByVal value As Date)
            _expiry = value
        End Set
    End Property

    Public Property AuctionDate As Date
        Get
            Return _auction
        End Get
        Set(ByVal value As Date)
            _auction = value
        End Set
    End Property

    Public Property BranchCode As String
        Get
            Return _branch
        End Get
        Set(ByVal value As String)
            _branch = value
        End Set
    End Property
#End Region
End Class
