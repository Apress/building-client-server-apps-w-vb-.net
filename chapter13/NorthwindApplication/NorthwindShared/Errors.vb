Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Structures

Namespace Errors

    Public Class AllowedValuesException
        Inherits System.ApplicationException

        Public Sub New(ByVal strValues As String)
            MyBase.New("The value must be one of the following: " & strValues)
        End Sub
    End Class

    Public Class FutureDateException
        Inherits System.ApplicationException

        Public Sub New()
            MyBase.New("This date can not occur in the future.")
        End Sub
    End Class

    Public Class UnderAgeException
        Inherits System.ApplicationException

        Public Sub New(ByVal intAge As Integer)
            MyBase.New("This person must be over " & intAge.ToString & ".")
        End Sub
    End Class

    Public Class OverAgeException
        Inherits System.ApplicationException

        Public Sub New(ByVal intAge As Integer)
            MyBase.New("This person cannot be over the age of " & intAge & ".")
        End Sub
    End Class

    Public Class BeforeCompanyCreatedException
        Inherits System.ApplicationException

        Public Sub New()
            MyBase.New("This date cannot occur before the date the " _
            & "company was created on (1 January 1976).")
        End Sub
    End Class

    Public Class SpecificFutureDateException
        Inherits System.ApplicationException

        Public Enum Unit
            Days = 0
            Weeks = 1
            Months = 2
            Years = 3
        End Enum

        Public Sub New(ByVal intPeriod As Integer, ByVal intUnit As Unit)
            MyBase.New("This date may not occur more than " _
            & intPeriod.ToString & " " & intUnit.ToString & ".")
        End Sub
    End Class

    Public Class MaximumLengthException
        Inherits ApplicationException

        Public Sub New(ByVal MaxLength As Integer)
            MyBase.New("The maximum length for this value is " _
            & MaxLength.ToString & " characters.")
        End Sub
    End Class

    Public Class ZeroLengthException
        Inherits System.ApplicationException

        Public Sub New()
            MyBase.New("A value must be entered for this item.")
        End Sub
    End Class

End Namespace

