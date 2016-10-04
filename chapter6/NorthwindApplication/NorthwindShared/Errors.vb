Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Structures

Namespace Errors

    Public Class MaximumLengthException
        Inherits ApplicationException

        Public Sub New(ByVal MaxLength As Integer)
            MyBase.New("The maximum length for this value is " _
            & MaxLength.ToString & " characters.")
        End Sub
    End Class

    Public Class ZeroLengthException
        Inherits System.ApplicationException

        Public Sub New()
            MyBase.New("A value must be entered for this item.")
        End Sub
    End Class

    <Serializable()> Public Class BusinessErrors
        Inherits System.Collections.CollectionBase

        Public Sub Add(ByVal strProperty As String, ByVal strMessage As String)
            Dim obj As New structErrors
            obj.errProperty = strProperty
            obj.errMessage = strMessage
            list.Add(obj)
            obj = Nothing
        End Sub

        Public Function Item(ByVal Index As Integer) As structErrors
            Return CType(list.Item(Index), structErrors)
        End Function

        Public Sub Remove(ByVal Index As Integer)
            list.RemoveAt(Index)
        End Sub
    End Class

End Namespace

