Public Class BranchJE

    Private _branchName As String = String.Empty
    Private _branchCode As String = String.Empty
    Private _areaCode As String = String.Empty
    Private _fileURL As String = String.Empty
    Private _JEDate As Date
    Private _RFcode As String = String.Empty
    Private _JEdataSet As DataSet

#Region "Properties"
    Public Property BranchCode As String
        Get
            Return _branchCode
        End Get
        Set(ByVal value As String)
            _branchCode = value
        End Set
    End Property

    Public Property BranchName As String
        Get
            If _branchName = "" Then
                _branchName = getBranchInfo(_branchCode, SheetNum.DistRole, 3)
            End If
            Return _branchName
        End Get
        Set(ByVal value As String)
            _branchName = value
        End Set
    End Property

    Public Property AreaCode As String
        Get
            Return _areaCode
        End Get
        Set(ByVal value As String)
            _areaCode = value
        End Set
    End Property

    Public Property DocDate As Date
        Get
            Return _JEDate
        End Get
        Set(ByVal value As Date)
            _JEDate = value
        End Set
    End Property

    Public Property JournalEntries As DataSet
        Get
            Return _JEdataSet
        End Get
        Set(ByVal value As DataSet)
            _JEdataSet = value
        End Set
    End Property

    Public Property FileURL As String
        Get
            Return _fileURL
        End Get
        Set(ByVal value As String)
            _fileURL = value
        End Set
    End Property

    Public Property RevolvingFund As String
        Get
            Return _RFcode
        End Get
        Set(ByVal value As String)
            _RFcode = value
        End Set
    End Property
#End Region

End Class
