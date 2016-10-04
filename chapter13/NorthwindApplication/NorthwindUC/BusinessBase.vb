Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindShared.Interfaces
Imports BusinessRules.Errors
Imports NorthwindTraders.NorthwindShared.Errors

Public MustInherit Class BusinessBase
    Protected mblnDirty As Boolean = False
    Public Loading As Boolean
    Protected WithEvents mobjRules As BrokenRules
    Public Event ObjectChanged(ByVal sender As Object, ByVal e As ChangedEventArgs)
    Public Event BrokenRule(ByVal IsBroken As Boolean)
    Private mstrListener As String
    Protected mobjVal As BusinessRules.Validate.Validation

    Public Sub New(ByVal RemotePath As String)
        mstrListener = RemotePath
        mobjRules = New BrokenRules
    End Sub

    Public ReadOnly Property IsDirty() As Boolean
        Get
            Return mblnDirty
        End Get
    End Property

    Public ReadOnly Property IsValid() As Boolean
        Get
            If mobjRules.Count > 0 Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Private Sub mobjRules_RuleBroken(ByVal IsBroken As Boolean) Handles mobjRules.RuleBroken
        RaiseEvent BrokenRule(IsBroken)
    End Sub

    Protected Sub CallChangedEvent(ByVal sender As Object, ByVal e As ChangedEventArgs)
        RaiseEvent ObjectChanged(sender, e)
    End Sub

    Public Function GetBusinessRules() As BusinessErrors
        Dim objInterface As IBaseInterface
        Dim objBusRules As BusinessErrors

        objInterface = CType(Activator.GetObject(GetType(IBaseInterface), _
        AppConstants.REMOTEOBJECTS & mstrListener), IBaseInterface)
        objBusRules = objInterface.GetBusinessRules
        objInterface = Nothing

        Return objBusRules
    End Function

    Protected Sub Validate(ByVal strProperty As String)
        Try
            mblnDirty = True
            mobjVal.ValidateAndThrow(Me, strProperty)
            mobjRules.BrokenRule(strProperty, False)
        Catch exc As Exception
            mobjRules.BrokenRule(strProperty, True)
            Throw exc
        End Try
    End Sub
End Class