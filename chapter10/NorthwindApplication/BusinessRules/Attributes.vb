Option Explicit On 
Option Strict On

Imports System.Reflection
Imports System.Text.RegularExpressions

Namespace Attributes

    'Interface for comparison and business rule retrieval
    '--TestCondition takes the value of the property or field and tests 
    'the value against whatever condition the class implements
    '--GetRule returns a string describing the rule in plain english
    'in relationship to the value that was supplied to it
    'For example, if the rule is the MaxLengthAttribute, and the
    'instantiation looked as follows: <MaxLength(10)> the GetRule
    'function would return the following string:
    '"Value cannot be longer than 10 characters."
    Public Interface ITest
        Function TestCondition(ByVal Value As Object, ByRef cls As Object) As Boolean
        Function GetRule() As String
    End Interface

    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class DisplayNameAttribute
        Inherits System.Attribute

        Private _strValue As String

        Public Sub New(ByVal Value As String)
            _strValue = Value
        End Sub

        Public ReadOnly Property Name() As String
            Get
                Return _strValue
            End Get
        End Property
    End Class

    'Usage: String or Numeric Properties or Fields
    'Note, the behavior of this attribute is different if the value is a string
    'or an integer.
    'String Behavior: Tests for a null value
    'Integer Behavior: Converts the number to an int64 (Long). If the value is
    '   zero (0) then it is considered a null. The reason for this is because
    '   all numeric values are defaulted to zero upon creation in .NET
    'Return: True if null
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class NotNullAttribute
        Inherits System.Attribute

        Implements ITest

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            If Value Is Nothing Then
                Return True
            Else
                If IsNumeric(Value) Then
                    If Convert.ToDecimal(Value) = 0 Then
                        Return True
                    End If
                End If
                Return False
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value cannot be null."
        End Function
    End Class

    'Usage: String Properties or Fields
    'Notes: Checks for an empty string value (zero length)
    'Return: True if zero length
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class NotEmptyAttribute
        Inherits System.Attribute

        Implements ITest

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            If Value Is Nothing Then
                Return True
            Else
                Dim str As String = CType(Value, String)
                If str.Trim.Length = 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value cannot be a zero length string."
        End Function
    End Class

    'Usage: String or Char Properties or Fields
    'Checks to determine whether a string is less than the supplied
    'number of characters in length
    'Example:
    '   <(MaxLength(10)> 
    'returns true if the number of characters in the field or 
    'property is over 10 characters in length.
    'Return: True if the value has more characters than allowed
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
Public Class MaxLengthAttribute
        Inherits System.Attribute

        Implements ITest

        Private _intValue As Integer

        Public Sub New(ByVal Value As Integer)
            _intValue = Value
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim strValue As String = Convert.ToString(Value)

            If strValue.Length > _intValue Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value cannot be longer than " & _intValue & " characters."
        End Function
    End Class

    'Usage: Numeric Properties or Fields
    'Used for numeric fields to determine if the value of the property or
    'field is larger than the allowable value.
    'Return: True if the value is larger than the specified value
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class MaxValueAttribute
        Inherits System.Attribute

        Implements ITest

        Private _lngValue As Long

        Public Sub New(ByVal Value As Long)
            _lngValue = Value
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim intValue As Long = Convert.ToInt64(Value)

            If intValue > _lngValue Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value cannot be greater than " & _lngValue & "."
        End Function
    End Class

    'Usage: String or Char Properties or Fields
    'Determines whether a string is less than an allowable value.
    'An example might be a password string which must be longer than
    '6 characters in length.
    'Return: True if the string is less than the specified length
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class MinLengthAttribute
        Inherits System.Attribute

        Implements ITest

        Private _intValue As Integer

        Public Sub New(ByVal Value As Integer)
            _intValue = Value
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim strValue As String = Convert.ToString(Value)

            If strValue.Length < _intValue Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value cannot be less than " & _intValue & " characters in length."
        End Function
    End Class

    'Usage: Date Properties or Fields
    'Return: True if the date occurs before this date.
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class DateNotBeforeAttribute
        Inherits System.Attribute

        Implements ITest

        Private _intValue As Date

        Public Sub New(ByVal Value As Date)
            _intValue = Value
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim dteValue As Date = Convert.ToDateTime(Value)

            If dteValue < _intValue Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value cannot be earlier than " & _intValue.ToShortDateString & "."
        End Function
    End Class

    'Usage: Date Properties or Fields
    'Determines if a date occurs after a specified date
    'return: True if the date occurs after this date
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class DateNotAfterAttribute
        Inherits System.Attribute

        Implements ITest

        Private _intValue As Date

        Public Sub New(ByVal Value As Date)
            _intValue = Value
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim dteValue As Date = Convert.ToDateTime(Value)

            If dteValue > _intValue Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value cannot be later than " & _intValue.ToShortDateString & "."
        End Function
    End Class

    'Usage: String or Char Properties or Fields
    'Determines if the value of the property or field is equal to one of the
    'specified values.
    'At this time the attribute accepts only up to eight(8) values in the
    'format "1", "2", "3"
    'Comparison is case INSENSITIVE. Everything is converted to uppercase
    'for comparison.
    'Return: True if the value of the property or field is not found in the
    'list of allowable values.
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class StringValueInAttribute
        Inherits System.Attribute

        Implements ITest

        Private _strValues As ArrayList = New ArrayList

#Region " Constructors"

        Public Sub New(ByVal Value0 As String)
            _strValues.Add(Value0)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String, _
        ByVal Value2 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
            _strValues.Add(Value2)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String, _
        ByVal Value2 As String, ByVal Value3 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
            _strValues.Add(Value2) : _strValues.Add(Value3)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String, _
        ByVal Value2 As String, ByVal Value3 As String, _
        ByVal Value4 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
            _strValues.Add(Value2) : _strValues.Add(Value3)
            _strValues.Add(Value4)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String, _
        ByVal Value2 As String, ByVal Value3 As String, _
        ByVal Value4 As String, ByVal Value5 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
            _strValues.Add(Value2) : _strValues.Add(Value3)
            _strValues.Add(Value4) : _strValues.Add(Value5)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String, _
        ByVal Value2 As String, ByVal Value3 As String, _
        ByVal Value4 As String, ByVal Value5 As String, _
        ByVal Value6 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
            _strValues.Add(Value2) : _strValues.Add(Value3)
            _strValues.Add(Value4) : _strValues.Add(Value5)
            _strValues.Add(Value6)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String, _
        ByVal Value2 As String, ByVal Value3 As String, _
        ByVal Value4 As String, ByVal Value5 As String, _
        ByVal Value6 As String, ByVal Value7 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
            _strValues.Add(Value2) : _strValues.Add(Value3)
            _strValues.Add(Value4) : _strValues.Add(Value5)
            _strValues.Add(Value6) : _strValues.Add(Value7)
        End Sub

        Public Sub New(ByVal Value0 As String, ByVal Value1 As String, _
        ByVal Value2 As String, ByVal Value3 As String, _
        ByVal Value4 As String, ByVal Value5 As String, _
        ByVal Value6 As String, ByVal Value7 As String, _
        ByVal Value8 As String)
            _strValues.Add(Value0) : _strValues.Add(Value1)
            _strValues.Add(Value2) : _strValues.Add(Value3)
            _strValues.Add(Value4) : _strValues.Add(Value5)
            _strValues.Add(Value6) : _strValues.Add(Value7)
            _strValues.Add(Value8)
        End Sub

#End Region

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim strValue As String = Convert.ToString(Value)
            Dim i As Integer

            For i = 0 To _strValues.Count - 1
                If strValue.ToUpper = Convert.ToString(_strValues.Item(i)).ToUpper Then
                    Return False
                End If
            Next

            Return True
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Dim i As Integer
            Dim strIn As String

            For i = 0 To _strValues.Count - 1
                strIn += Convert.ToString(_strValues(i))
                If i < _strValues.Count - 1 Then
                    strIn += ", "
                End If
            Next

            Return "Value must be one of the following values: " & strIn & "."
        End Function
    End Class

    'Usage: Any property or field
    'Determines if the value of a property or field matches the given
    'regular expression patter.
    'Return: True if their is no match.
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class PatternMatchAttribute
        Inherits System.Attribute

        Implements ITest

        Private _strPattern As String

        Public Sub New(ByVal Pattern As String)
            _strPattern = Pattern
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim strValue As String = CType(Value, String)
            If Not strValue Is Nothing Then
                Dim m As Match = Regex.Match(strValue, _strPattern)

                If m.Value = "" Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value must match the following pattern: " & _strPattern & "."
        End Function
    End Class

    'Usage: Char Properties or Fields
    'Determines if the value of the property or field is between two given
    'character values (note that this does not have to be limited to letters
    'as this is a unicode value comparison). Note that this includes the
    'beginning and ending values to test for.
    'Example <CharacterInRange("A", "Z")> will see if the value of the property
    'matches any character between A and Z inclusive.
    'Return: True if the value of the property or field is not found in the
    'range of values.
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class CharacterInRangeAttribute
        Inherits System.Attribute

        Implements ITest

        Private _chrStart As Char
        Private _chrEnd As Char

        Public Sub New(ByVal chrStart As Char, ByVal chrEnd As Char)
            _chrStart = chrStart
            _chrEnd = chrEnd
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim c As Char = Convert.ToChar(Value)

            If c.CompareTo(_chrStart) >= 0 And c.CompareTo(_chrEnd) <= 0 Then
                Return False
            Else
                Return True
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value must be between " & _chrStart & " and " & _chrEnd & "."
        End Function
    End Class

    'The value must be null if the value of a specific field is not equal to the
    'indicated value.
    'For example: An employee must be a regional directory (R) position in order
    'for there to be a signing name, otherwise, the signing name must be null
    'So, when specifying this business rule, it would look like this:
    '<(NullIfNotValue("Position","R")> Where "Position" is the field to look
    'for the value in and "R" is what the value must be in order for this property
    '(the signing name) to be not null.
    <AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
    Public Class NullIfNotValueAttribute
        Inherits System.Attribute

        Implements ITest

        'Field to compare against
        Private _strField As String
        'Value that the field must be for this field to be not null
        Private _strFieldValue As String

        Public Sub New(ByVal Field As String, ByVal FieldValue As String)
            _strField = Field
            _strFieldValue = FieldValue
        End Sub

        Public Function TestCondition(ByVal Value As Object, ByRef cls As Object) _
        As Boolean Implements ITest.TestCondition
            Dim strValue As String = CType(Value, String)
            Dim m As MemberInfo
            Dim t As Type = cls.GetType
            Dim mi As MemberInfo() = t.GetMembers
            Dim strActualValue As String
            Dim strType As String

            For Each m In mi
                If m.Name = _strField Then
                    If m.MemberType = MemberTypes.Field Then
                        Dim f As FieldInfo = CType(m, FieldInfo)
                        strActualValue = Convert.ToString(f.GetValue(cls))
                        strType = f.FieldType.ToString
                        Exit For
                    End If

                    If m.MemberType = MemberTypes.Property Then
                        Dim p As PropertyInfo = CType(m, PropertyInfo)
                        strActualValue = Convert.ToString(p.GetValue(p, Nothing))
                        strType = p.PropertyType.ToString
                        Exit For
                    End If
                End If
            Next

            If Value.GetType Is GetType(System.String) Then
                If strActualValue <> _strFieldValue Then
                    If strValue Is Nothing Then
                        Return False
                    Else
                        If strValue = "" Then
                            Return False
                        Else
                            Return True
                        End If
                    End If
                Else
                    Return False
                End If
            Else
                If Value.GetType Is GetType(System.Int32) Then
                    If strActualValue <> _strFieldValue Then
                        If Not strValue = "0" Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                End If
            End If
        End Function

        Public Function GetRule() As String Implements ITest.GetRule
            Return "Value must be null if " & _strField & " is not equal to " & _strFieldValue & "."
        End Function
    End Class
End Namespace
