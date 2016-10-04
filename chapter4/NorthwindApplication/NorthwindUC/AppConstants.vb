Option Explicit On 
Option Strict On

Public Class AppConstants
    Public Shared REMOTEOBJECTS As String = "http://Mcsd-gq85jxhy62:80/Northwind/"
End Class

Public Class ChangedEventArgs
    Public Enum eChange
        Added = 0
        Updated = 1
    End Enum

    Private meChange As eChange

    Public Sub New(ByVal ChangeEvent As eChange)
        meChange = ChangeEvent
    End Sub

    Public ReadOnly Property Change() As eChange
        Get
            Return meChange
        End Get
    End Property


End Class
