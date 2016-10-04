Option Explicit On 
Option Strict On

Namespace Structures

<Serializable()> Public Structure structRegion
    Public RegionID As Integer
    Public RegionDescription As String
End Structure

    <Serializable()> Public Structure structErrors
        Public errProperty As String
        Public errMessage As String
    End Structure

End Namespace