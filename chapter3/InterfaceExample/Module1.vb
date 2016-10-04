Module Module1

    Sub Main()
        Dim objAdd As New cAdd
        Dim objTip As New CalcTip
        Console.WriteLine("Processing cAdd Class.")
        ProcessNumbers(objAdd)
        Console.WriteLine("Processing CalcTip Class.")
        ProcessNumbers(objTip)
        Console.ReadLine()
    End Sub

    Private Sub ProcessNumbers(ByVal objI As ITest)
        Console.WriteLine(objI.Add(4, 5))
    End Sub
End Module

Public Interface ITest
    Function Add(ByVal int1 As Integer, ByVal int2 As Integer) As Integer
End Interface

Public Class cAdd
    Implements ITest

    Public Function Add(ByVal int1 As Integer, ByVal int2 As Integer) As Integer Implements ITest.Add
        Return int1 + int2
    End Function
End Class

Public Class CalcTip
    Implements ITest

    Public Function Add(ByVal int1 As Integer, ByVal int2 As Integer) As Integer Implements ITest.Add
        Return (int1 + int2) * 9.25
    End Function
End Class