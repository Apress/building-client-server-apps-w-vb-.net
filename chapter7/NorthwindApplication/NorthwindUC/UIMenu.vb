Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Interfaces

Public Class UIMenu
    Private Shared mUIMenu As UIMenu

    Public Shared Function getInstance() As UIMenu
        If mUIMenu Is Nothing Then
            mUIMenu = New UIMenu
        End If
        Return mUIMenu
    End Function

    Protected Sub New()
        'Do nothing
    End Sub

    Public Function GetMenuStructure() As DataSet
        Dim objIMenu As IMenu
        Dim ds As DataSet

        objIMenu = CType(Activator.GetObject(GetType(IMenu), AppConstants.REMOTEOBJECTS & "MenuDC.rem"), IMenu)
        ds = objIMenu.GetMenuStructure
        objIMenu = Nothing

        Return ds
    End Function
End Class

