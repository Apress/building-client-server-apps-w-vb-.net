Option Explicit On 
Option Strict On

Imports BusinessRules.Attributes
Imports BusinessRules.Errors
Imports System.Reflection

Namespace Validate
    Public Class Validation

        Private Shared Function GetDisplayName(ByVal member As MemberInfo) As String
            Dim obj() As Object = member.GetCustomAttributes(True)
            Dim i As Integer
            If obj.Length > 0 Then
                For i = 0 To obj.Length - 1
                    If TypeOf (obj(i)) Is DisplayNameAttribute Then
                        Dim objDisplayNameAttribute As DisplayNameAttribute = CType(obj(i), DisplayNameAttribute)
                        Return objDisplayNameAttribute.Name
                    End If
                Next i
            End If
            Return member.Name
        End Function

        Public Shared Function GetBusinessRules(ByVal cls As Object) As BusinessErrors
            Dim t As Type = cls.GetType
            Dim m As MemberInfo() = t.GetMembers
            Dim i As Integer

            Dim objBusErr As New BusinessErrors

            For i = 0 To m.Length - 1
                Dim obj() As Object = m(i).GetCustomAttributes(True)
                If obj.Length > 0 Then
                    Dim j As Integer
                    For j = 0 To obj.Length - 1
                        If TypeOf obj(j) Is ITest Then
                            Dim objI As ITest = CType(obj(j), ITest)
                            objBusErr.Add(GetDisplayName(m(i)), objI.GetRule)
                        End If
                    Next
                End If
            Next

            Return objBusErr
        End Function

        Public Shared Function validate(ByVal cls As Object) As BusinessErrors
            Dim t As Type = cls.GetType
            Dim i, j As Integer
            Dim bln As Boolean

            Dim objBusErr As New BusinessErrors

            Dim m As MemberInfo() = t.GetMembers

            For i = 0 To m.Length - 1
                Dim objAttrib() As Object = m(i).GetCustomAttributes(True)
                For j = 0 To objAttrib.Length - 1
                    If TypeOf objAttrib(j) Is ITest Then
                        Dim objI As ITest = CType(objAttrib(j), ITest)

                        If TypeOf m(i) Is FieldInfo Then
                            Dim fld As FieldInfo = CType(m(i), FieldInfo)
                            bln = objI.TestCondition(fld.GetValue(cls), cls)
                        End If

                        If TypeOf m(i) Is PropertyInfo Then

                            Dim pro As PropertyInfo = CType(m(i), PropertyInfo)
                            bln = objI.TestCondition(pro.GetValue(cls, Nothing), cls)
                        End If

                        If bln Then
                            objBusErr.Add(m(i).Name, objI.GetRule())
                        End If
                    End If
                Next
            Next

            Return objBusErr
        End Function

        Public Shared Sub ValidateAndThrow(ByVal cls As Object, ByVal field As String)
            Dim t As Type = cls.GetType
            Dim m As MemberInfo() = t.GetMember(field)
            Dim i As Integer
            Dim bln As Boolean

            Dim obj() As Object = m(0).GetCustomAttributes(True)
            If obj.Length > 0 Then
                For i = 0 To obj.Length - 1
                    If TypeOf obj(i) Is ITest Then
                        Dim objI As ITest = CType(obj(i), ITest)

                        If TypeOf m(0) Is FieldInfo Then
                            Dim fld As FieldInfo = CType(m(0), FieldInfo)
                            bln = objI.TestCondition(fld.GetValue(cls), cls)
                        End If

                        If TypeOf m(0) Is PropertyInfo Then
                            Dim pro As PropertyInfo = CType(m(0), PropertyInfo)
                            bln = objI.TestCondition(pro.GetValue(cls, Nothing), cls)
                        End If

                        If bln Then
                            Throw New Exception(objI.GetRule())
                        End If
                    End If
                Next
            End If
        End Sub
    End Class
End Namespace