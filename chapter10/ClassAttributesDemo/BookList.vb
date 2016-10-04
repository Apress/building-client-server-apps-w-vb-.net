Public Class BookList
    Private mstrTitle As String
    Private mstrAuthor As String
    Private mdblPrice As Double
    Private mstrPublisher As String

    Public ReadOnly Property Price() As Double
        Get
            Return mdblPrice
        End Get
    End Property

    Public ReadOnly Property Publisher() As String
        Get
            Return mstrPublisher
        End Get
    End Property

    <List("Author", 1)> Public ReadOnly Property Author() As String
        Get
            Return mstrAuthor
        End Get
    End Property

    <List("Book Title", 0)> Public ReadOnly Property Title() As String
        Get
            Return mstrTitle
        End Get
    End Property

    Public Sub New(ByVal sTitle As String, ByVal sAuthor As String, _
    ByVal dPrice As Double, ByVal sPub As String)
        mstrTitle = sTitle
        mstrAuthor = sAuthor
        mdblPrice = dPrice
        mstrPublisher = sPub
    End Sub

End Class

Public Class BookListMgr
    Inherits System.Collections.CollectionBase

    Public Sub Add(ByVal obj As BookList)
        list.Add(obj)
    End Sub

    Public Sub Remove(ByVal Index As Integer)
        list.RemoveAt(Index)
    End Sub

    Public Function Item(ByVal Index As Integer) As BookList
        Return CType(list.Item(Index), BookList)
    End Function
End Class