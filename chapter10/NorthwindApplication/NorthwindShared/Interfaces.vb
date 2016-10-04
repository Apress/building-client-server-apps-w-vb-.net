Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindShared.Structures
Imports BusinessRules.Errors
Imports NorthwindTraders.NorthwindShared.Errors

Namespace Interfaces

    Public Interface IEmployee
        Inherits IBaseInterface
        Function LoadRecord(ByVal intID As Integer) As structEmployee
        Function Save(ByVal sEmployee As structEmployee, _
             ByRef intID As Integer) As BusinessErrors
        Sub Delete(ByVal intID As Integer)
    End Interface

    Public Interface ITerritory
        Inherits IBaseInterface
        Function LoadRecord(ByVal strID As String) As structTerritory
        Function Save(ByVal sTerritory As structTerritory) _
        As BusinessErrors
        Sub Delete(ByVal strID As String)
    End Interface

    Public Interface IRegion
        Inherits IBaseInterface
        Function LoadRecord(ByVal intID As Integer) As structRegion
        Function Save(ByVal sRegion As structRegion, ByRef intID As Integer) _
        As BusinessErrors
        Sub Delete(ByVal intID As Integer)
    End Interface

    Public Interface IBaseInterface
        Function LoadProxy() As DataSet
        Function GetBusinessRules() As BusinessErrors
    End Interface

    Public Interface ILogError
        Sub Save(ByVal exc As Exception)
    End Interface

    Public Interface IMenu
        Function GetMenuStructure() As DataSet
    End Interface

End Namespace
