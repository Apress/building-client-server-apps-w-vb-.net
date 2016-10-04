Option Strict On
Option Explicit On 

Public Class BrokenRules
    Inherits System.Collections.DictionaryBase

    Public Event RuleBroken(ByVal IsBroken As Boolean)

    Public Sub BrokenRule(ByVal strProperty As String, ByVal blnBroken As _
Boolean)
        Try
            If blnBroken Then
                dictionary.Add(strProperty, blnBroken)
            Else
                dictionary.Remove(strProperty)
            End If
        Catch
            'Do Nothing
        Finally
            If dictionary.Count > 0 Then
                RaiseEvent RuleBroken(True)
            Else
                RaiseEvent RuleBroken(False)
            End If
        End Try
    End Sub

    Public Sub ClearRules()
        dictionary.Clear()
    End Sub
End Class