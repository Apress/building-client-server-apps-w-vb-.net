Option Explicit On 
Option Strict On

Namespace Errors
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

    <Serializable()> Public Structure structErrors
        Public errProperty As String
        Public errMessage As String
    End Structure
End Namespace