Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindShared.Structures

Namespace Interfaces

    Public Interface IRegion
        Function LoadProxy() As DataSet
        Function LoadRecord(ByVal intID As Integer) As structRegion
        Sub Save(ByVal sRegion As structRegion, ByRef intID As Integer)
        Sub Delete(ByVal intID As Integer)
    End Interface

End Namespace
