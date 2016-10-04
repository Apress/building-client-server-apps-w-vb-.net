Option Explicit On 
Option Strict On

Namespace Structures

    <Serializable()> Public Structure structEmployee
        Public EmployeeID As Integer
        Public LastName As String
        Public FirstName As String
        Public Title As String
        Public TitleOfCourtesy As String
        Public BirthDate As Date
        Public HireDate As Date
        Public Address As String
        Public City As String
        Public Region As String
        Public PostalCode As String
        Public Country As String
        Public HomePhone As String
        Public Extension As String
        Public Photo() As Byte
        Public Notes As String
        Public ReportsTo As Integer
        Public ReportsToFirstName As String
        Public ReportsToLastName As String
        Public PhotoPath As String
        Public Territories() As String
    End Structure

    <Serializable()> Public Structure structTerritory
        Public TerritoryID As String
        Public TerritoryDescription As String
        Public RegionID As Integer
        Public IsNew As Boolean
    End Structure

    <Serializable()> Public Structure structRegion
        Public RegionID As Integer
        Public RegionDescription As String
    End Structure

    <Serializable()> Public Structure structErrors
        Public errProperty As String
        Public errMessage As String
    End Structure

End Namespace