Option Strict On
Option Explicit On 

Public Class ListViewColumnSorter
    Implements System.Collections.IComparer

    Private mintSortCol As Integer = 0
    Private mobjOrder As SortOrder = SortOrder.None
    Private mobjCompare As CaseInsensitiveComparer

    Public Sub New()
        mobjCompare = New CaseInsensitiveComparer
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements _
    IComparer.Compare
        Dim intResult As Integer
        Dim lvwItem1 As ListViewItem
        Dim lvwItem2 As ListViewItem

        lvwItem1 = CType(x, ListViewItem)
        lvwItem2 = CType(y, ListViewItem)

        intResult = mobjCompare.Compare(lvwItem1.SubItems(mintSortCol).Text, _
        lvwItem2.SubItems(mintSortCol).Text)

        If (mobjOrder = SortOrder.Ascending) Then
            Return intResult
        Else
            If (mobjOrder = SortOrder.Descending) Then
                Return (-intResult)
            Else
                Return 0
            End If
        End If
    End Function

    Public Property SortColumn() As Integer
        Get
            Return mintSortCol
        End Get
        Set(ByVal Value As Integer)
            mintSortCol = Value
        End Set
    End Property

    Public Property Order() As SortOrder
        Get
            Return mobjOrder
        End Get
        Set(ByVal Value As SortOrder)
            mobjOrder = Value
        End Set
    End Property
End Class

Public Class Win32APIcalls

    Public Declare Function GetKeyState Lib "user32" _
    Alias "GetKeyState" (ByVal nVirtKey As Long) As Integer

End Class
