<AttributeUsage(AttributeTargets.Property)> _
Public Class ListAttribute
    Inherits System.Attribute

    Public Heading As String
    Public Column As Integer

    Public Sub New(ByVal Header As String, ByVal Col As Integer)
        Heading = Header
        Column = Col
    End Sub
End Class