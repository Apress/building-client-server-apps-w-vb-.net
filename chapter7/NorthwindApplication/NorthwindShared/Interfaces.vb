Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindShared.Structures
Imports NorthwindTraders.NorthwindShared.Errors

Namespace Interfaces

    Public Interface ITerritory
        Function LoadProxy() As DataSet
        Function LoadRecord(ByVal strID As String) As structTerritory
        Function Save(ByVal sTerritory As structTerritory) _
        As BusinessErrors
        Sub Delete(ByVal strID As String)
        Function GetBusinessRules() As BusinessErrors
    End Interface

    Public Interface IRegion
        Function LoadProxy() As DataSet
        Function LoadRecord(ByVal intID As Integer) As structRegion
        Function Save(ByVal sRegion As structRegion, ByRef intID As Integer) As BusinessErrors
        Sub Delete(ByVal intID As Integer)
        Function GetBusinessRules() As BusinessErrors
    End Interface

    Public Interface ILogError
        Sub Save(ByVal exc As Exception)
    End Interface

    Public Interface IMenu
        Function GetMenuStructure() As DataSet
    End Interface

End Namespace
