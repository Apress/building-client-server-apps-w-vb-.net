Public Class ComputerList
    Private mstrName As String
    Private mstrProc As String
    Private mdblSpeed As Double
    Private mdblPrice As Double
    Private mstrManuf As String

    
    <List("Processor", 1)> Public ReadOnly Property Proc() As String
        Get
            Return mstrProc
        End Get
    End Property

    <List("Speed", 2)> Public ReadOnly Property Speed() As Double
        Get
            Return mdblSpeed
        End Get
    End Property

    <List("Computer Name", 0)> Public ReadOnly Property Cname() As String
        Get
            Return mstrName
        End Get
    End Property

    Public ReadOnly Property Price() As Double
        Get
            Return mdblPrice
        End Get
    End Property

    Public ReadOnly Property Manufacturer() As String
        Get
            Return mstrManuf
        End Get
    End Property

    Public Sub New(ByVal Name As String, ByVal Process As String, ByVal Sp As Double, _
    ByVal Pr As Double, ByVal Man As String)
        mstrName = Name
        mstrProc = Process
        mdblSpeed = Sp
        mdblPrice = Pr
        mstrManuf = Man
    End Sub
End Class

Public Class ComputerListMgr
    Inherits System.Collections.CollectionBase

    Public Sub Add(ByVal obj As ComputerList)
        list.Add(obj)
    End Sub

    Public Sub Remove(ByVal Index As Integer)
        list.RemoveAt(Index)
    End Sub

    Public Function Item(ByVal Index As Integer) As ComputerList
        Return CType(list.Item(Index), ComputerList)
    End Function
End Class