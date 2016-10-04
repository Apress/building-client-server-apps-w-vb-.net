Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Interfaces
Imports NorthwindTraders.NorthwindShared.Structures
Imports BusinessRules.Errors
Imports NorthwindTraders.NorthwindShared.Errors
Imports System.Configuration
Imports System.Data.SqlClient

Public Class EmployeeDC
    Inherits MarshalByRefObject

    Implements IEmployee

    Private mobjBusErr As BusinessErrors

#Region " Private Attributes"

    Private mintEmployeeID As Integer
    Private mstrLastName As String
    Private mstrFirstName As String
    Private mstrTitle As String
    Private mstrTitleOfCourtesy As String
    Private mdteBirthDate As Date
    Private mdteHireDate As Date
    Private mstrAddress As String
    Private mstrCity As String
    Private mstrRegion As String
    Private mstrPostalCode As String
    Private mstrCountry As String
    Private mstrHomePhone As String
    Private mstrExtension As String
    Private mbytPhoto() As Byte
    Private mstrNotes As String
    Private mintReportsTo As Integer
    Private mstrPhotoPath As String
    Private mstrTerritory() As String

#End Region

#Region " Public Attributes"

    Public ReadOnly Property EmployeeID() As Integer
        Get
            Return mintEmployeeID
        End Get
    End Property

    Public Property LastName() As String
        Get
            Return mstrLastName
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Last Name")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 20 Then
                    Throw New MaximumLengthException(20)
                End If

                mstrLastName = Value
            Catch exc As Exception
                mobjBusErr.Add("Last Name", exc.Message)
            End Try
        End Set
    End Property

    Public Property FirstName() As String
        Get
            Return mstrFirstName
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("First Name")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 10 Then
                    Throw New MaximumLengthException(10)
                End If

                mstrFirstName = Value
            Catch exc As Exception
                mobjBusErr.Add("First Name", exc.Message)
            End Try
        End Set
    End Property

    Public Property Title() As String
        Get
            Return mstrTitle
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Title")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 30 Then
                    Throw New MaximumLengthException(30)
                End If

                mstrTitle = Value
            Catch exc As Exception
                mobjBusErr.Add("Title", exc.Message)
            End Try
        End Set
    End Property

    Public Property TitleOfCourtesy() As String
        Get
            Return mstrTitleOfCourtesy
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Exit Property
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    mstrTitleOfCourtesy = Nothing
                    Exit Property
                End If

                'Test for max length
                If Value.Length > 25 Then
                    Throw New MaximumLengthException(25)
                End If

                'Test for specific values
                Select Case Value
                    Case "Mr.", "Ms.", "Dr.", "Mrs."
                        'Do nothing
                    Case Else
                        Throw New AllowedValuesException("Mr., Ms., " _
                        & "Dr., Mrs.")
                End Select

                mstrTitleOfCourtesy = Value
            Catch exc As Exception
                mobjBusErr.Add("Title Of Courtesy", exc.Message)
            End Try
        End Set
    End Property

    Public Property BirthDate() As Date
        Get
            Return mdteBirthDate
        End Get
        Set(ByVal Value As Date)
            Try
                'Check for a future date
                If Value > Now Then
                    Throw New FutureDateException
                End If

                'Check for under 18
                If DateDiff(DateInterval.Day, Value, Now) < (18 * 365) _
                Then
                    Throw New UnderAgeException(18)
                End If

                'Check for over 60
                If DateDiff(DateInterval.Year, Value, Now) > 60 Then
                    Throw New OverAgeException(60)
                End If

                mdteBirthDate = Value
            Catch exc As Exception
                mobjBusErr.Add("Birth Date", exc.Message)
            End Try
        End Set
    End Property

    Public Property HireDate() As Date
        Get
            Return mdteHireDate
        End Get
        Set(ByVal Value As Date)
            Try
                'Check for a future date
                If Value > DateAdd(DateInterval.Day, 14, Now) Then
                    Throw New SpecificFutureDateException(2, _
                    SpecificFutureDateException.Unit.Weeks)
                End If

                'Check for before company creation date
                If Value < #1/1/1976# Then
                    Throw New BeforeCompanyCreatedException
                End If

                mdteHireDate = Value
            Catch exc As Exception
                mobjBusErr.Add("Hire Date", exc.Message)
            End Try
        End Set
    End Property

    Public Property Address() As String
        Get
            Return mstrAddress
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Address")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 60 Then
                    Throw New MaximumLengthException(60)
                End If

                mstrAddress = Value
            Catch exc As Exception
                mobjBusErr.Add("Address", exc.Message)
            End Try
        End Set
    End Property

    Public Property City() As String
        Get
            Return mstrCity
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("City")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 15 Then
                    Throw New MaximumLengthException(15)
                End If

                mstrCity = Value
            Catch exc As Exception
                mobjBusErr.Add("City", exc.Message)
            End Try
        End Set
    End Property

    Public Property Region() As String
        Get
            Return mstrRegion
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then Exit Property

                'Test for empty string
                If Value.Length = 0 Then
                    mstrRegion = Nothing
                    Exit Property
                End If

                'Test for max length
                If Value.Length > 15 Then
                    Throw New MaximumLengthException(15)
                End If

                mstrRegion = Value
            Catch exc As Exception
                mobjBusErr.Add("Region", exc.Message)
            End Try
        End Set
    End Property

    Public Property PostalCode() As String
        Get
            Return mstrPostalCode
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Postal Code")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 10 Then
                    Throw New MaximumLengthException(10)
                End If

                mstrPostalCode = Value
            Catch exc As Exception
                mobjBusErr.Add("Postal Code", exc.Message)
            End Try
        End Set
    End Property

    Public Property Country() As String
        Get
            Return mstrCountry
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Country")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 15 Then
                    Throw New MaximumLengthException(15)
                End If

                mstrCountry = Value
            Catch exc As Exception
                mobjBusErr.Add("Country", exc.Message)
            End Try
        End Set
    End Property

    Public Property HomePhone() As String
        Get
            Return mstrHomePhone
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Home Phone")
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                'Test for max length
                If Value.Length > 24 Then
                    Throw New MaximumLengthException(24)
                End If

                mstrHomePhone = Value
            Catch exc As Exception
                mobjBusErr.Add("Home Phone", exc.Message)
            End Try
        End Set
    End Property

    Public Property Extension() As String
        Get
            Return mstrExtension
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then Exit Property

                'Test for empty string
                If Value.Length = 0 Then
                    mstrExtension = Nothing
                    Exit Property
                End If

                'Test for max length
                If Value.Length > 4 Then
                    Throw New MaximumLengthException(4)
                End If

                mstrExtension = Value
            Catch exc As Exception
                mobjBusErr.Add("Extension", exc.Message)
            End Try
        End Set
    End Property

    Public Property Photo() As Byte()
        Get
            Return mbytPhoto
        End Get
        Set(ByVal Value As Byte())
            If Value Is Nothing Then Exit Property

            mbytPhoto = Value
        End Set
    End Property

    Public Property Notes() As String
        Get
            Return mstrNotes
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then Exit Property

                'Test for empty string
                If Value.Length = 0 Then
                    mstrNotes = Nothing
                    Exit Property
                End If

                mstrNotes = Value
            Catch exc As Exception
                mobjBusErr.Add("Notes", exc.Message)
            End Try
        End Set
    End Property

    Public Property ReportsTo() As Integer
        Get
            Return mintReportsTo
        End Get
        Set(ByVal Value As Integer)
            Try
                'Test for null value
                If Value = 0 Then
                    mintReportsTo = Nothing
                    Exit Property
                End If

                mintReportsTo = Value
            Catch exc As Exception
                mobjBusErr.Add("Reports To", exc.Message)
            End Try
        End Set
    End Property

    Public Property PhotoPath() As String
        Get
            Return mstrPhotoPath
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then Exit Property

                'Test for empty string
                If Value.Length = 0 Then
                    mstrPhotoPath = Nothing
                    Exit Property
                End If

                'Test for max length
                If Value.Length > 255 Then
                    Throw New MaximumLengthException(255)
                End If

                'If FileSystem.FileLen(Value) > 0 Then
                'do nothing, if the file is not found, a
                'FileNotFoundException will be thrown automatically
                'End If

                mstrPhotoPath = Value
            Catch exc As Exception
                mobjBusErr.Add("Photo Path", exc.Message)
            End Try
        End Set
    End Property

    Public Property Territories() As String()
        Get
            Return mstrTerritory
        End Get
        Set(ByVal Value As String())
            Try
                If Value Is Nothing Then
                    Throw New Exception("An employee must be " _
                    & "assigned to at least one territory.")
                Else
                    If Value.Length = 0 Then
                        Throw New Exception("An employee must " _
                        & "be assigned to at least one territory.")
                    End If
                End If
                mstrTerritory = Value
            Catch exc As Exception
                mobjBusErr.Add("Territories", exc.Message)
            End Try
        End Set
    End Property
#End Region

    Public Function GetBusinessRules() As BusinessRules.Errors.BusinessErrors Implements NorthwindShared.Interfaces.IBaseInterface.GetBusinessRules
        Dim objBusRules As New BusinessErrors
        With objBusRules
            .Add("Last Name", "The value cannot be null.")
            .Add("Last Name", "The value cannot be more than 20 characters in length.")
            .Add("First Name", "The value cannot be null.")
            .Add("First Name", "The value cannot be more than 10 characters in length.")
            .Add("Title", "The value cannot be null.")
            .Add("Title", "The value cannot be more than 30 characters in length.")
            .Add("Title Of Courtesy", "The value cannot be more than 25 characters in length.")
            .Add("Title Of Courtesy", "The value must one of the following: Mr., Ms., Dr. or Mrs.")
            .Add("Birth Date", "The value cannot be a date in the future.")
            .Add("Birth Date", "The employee must be 18 years of age or older.")
            .Add("Birth Date", "The employee must be younger than 60 years of age.")
            .Add("Hire Date", "The value may not be more than two weeks in the future.")
            .Add("Hire Date", "The value may not be before the company was created.")
            .Add("Address", "The value cannot be null.")
            .Add("Address", "The value cannot be more than 60 characters in length.")
            .Add("City", "The value cannot be null.")
            .Add("City", "The value cannot be more than 15 characters in length.")
            .Add("Region", "The value cannot be more than 15 characters in length.")
            .Add("Postal Code", "The value cannot be more than 10 characters in length.")
            .Add("Country", "The value cannot be more than 15 characters in length.")
            .Add("Home Phone", "The value cannot be null.")
            .Add("Home Phone", "The value cannot be more than 24 characters in length.")
            .Add("Extension", "The value cannot be more than 4 characters in length.")
            .Add("Photo Path", "The value cannot be more than 255 characters in length.")
        End With

        Return objBusRules
    End Function

    Public Function LoadProxy() As System.Data.DataSet Implements NorthwindShared.Interfaces.IBaseInterface.LoadProxy
        Dim strCN As String = _
             ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_employee_getall"
        End With

        da.Fill(ds)

        cmd = Nothing
        cn.Close()

        Return ds
    End Function

    Public Sub Delete(ByVal intID As Integer) Implements NorthwindShared.Interfaces.IEmployee.Delete
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_employee_delete"
            .Parameters.Add("@id", intID)
            .ExecuteNonQuery()
        End With

        cmd = Nothing
        cn.Close()
    End Sub

    Public Function LoadRecord(ByVal intID As Integer) As NorthwindShared.Structures.structEmployee Implements NorthwindShared.Interfaces.IEmployee.LoadRecord
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        Dim dRow As DataRow
        Dim i As Integer
        Dim sEmployee As structEmployee

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_employee_getone"
            .Parameters.Add("@id", intID)
        End With

        da.Fill(ds)

        cmd = Nothing
        cn.Close()

        With ds.Tables(0).Rows(0)
            sEmployee.EmployeeID = Convert.ToInt32(.Item("EmployeeID"))
            sEmployee.LastName = Convert.ToString(.Item("LastName"))
            sEmployee.FirstName = Convert.ToString(.Item("FirstName"))
            sEmployee.Title = Convert.ToString(.Item("Title"))
            sEmployee.TitleOfCourtesy = _
                  Convert.ToString(.Item("TitleOfCourtesy"))
            sEmployee.BirthDate = Convert.ToDateTime(.Item("BirthDate"))
            sEmployee.HireDate = Convert.ToDateTime(.Item("HireDate"))
            sEmployee.Address = Convert.ToString(.Item("Address"))
            sEmployee.City = Convert.ToString(.Item("City"))
            If Not IsDBNull(.Item("Region")) Then
                sEmployee.Region = Convert.ToString(.Item("Region"))
            End If
            sEmployee.PostalCode = Convert.ToString(.Item("PostalCode"))
            sEmployee.Country = Convert.ToString(.Item("Country"))
            sEmployee.HomePhone = Convert.ToString(.Item("HomePhone"))
            If Not IsDBNull(.Item("Extension")) Then
                sEmployee.Extension = Convert.ToString(.Item("Extension"))
            End If
            If Not IsDBNull(.Item("Photo")) Then
                sEmployee.Photo = CType(.Item("Photo"), Byte())
            End If
            If Not IsDBNull(.Item("Notes")) Then
                sEmployee.Notes = Convert.ToString(.Item("Notes"))
            End If
            If Not IsDBNull(.Item("ReportsTo")) Then
                sEmployee.ReportsTo = Convert.ToInt32(.Item("ReportsTo"))
                sEmployee.ReportsToLastName = _
                      Convert.ToString(.Item("ReportsToLastName"))
                sEmployee.ReportsToFirstName = _
                      Convert.ToString(.Item("ReportsToFirstName"))
            End If
            If Not IsDBNull(.Item("PhotoPath")) Then
                sEmployee.PhotoPath = Convert.ToString(.Item("PhotoPath"))
            End If
        End With

        ReDim sEmployee.Territories(ds.Tables(1).Rows.Count - 1)

        For Each dRow In ds.Tables(1).Rows
            sEmployee.Territories(i) = _
                  Convert.ToString(dRow.Item("TerritoryID"))
            i += 1
        Next

        ds = Nothing

        Return sEmployee
    End Function

    Public Function Save(ByVal sEmployee As NorthwindShared.Structures.structEmployee, ByRef intID As Integer) As BusinessRules.Errors.BusinessErrors Implements NorthwindShared.Interfaces.IEmployee.Save
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim trans As SqlTransaction
        Dim i As Integer
        Dim intTempID As Integer

        intTempID = sEmployee.EmployeeID

        mobjBusErr = New BusinessErrors

        With sEmployee
            Me.mintEmployeeID = .EmployeeID
            Me.LastName = .LastName
            Me.FirstName = .FirstName
            Me.Title = .Title
            Me.TitleOfCourtesy = .TitleOfCourtesy
            Me.BirthDate = .BirthDate
            Me.HireDate = .HireDate
            Me.Address = .Address
            Me.City = .City
            Me.Region = .Region
            Me.PostalCode = .PostalCode
            Me.Country = .Country
            Me.HomePhone = .HomePhone
            Me.Extension = .Extension
            Me.Photo = .Photo
            Me.Notes = .Notes
            Me.ReportsTo = .ReportsTo
            Me.PhotoPath = .PhotoPath
            Me.Territories = .Territories
        End With

        If mobjBusErr.Count = 0 Then

            Try
                Dim prm As SqlParameter
                cn.Open()
                trans = cn.BeginTransaction()

                With cmd
                    .Connection = cn
                    .Transaction = trans
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "usp_employee_save"
                    .Parameters.Add("@id", mintEmployeeID)
                    .Parameters.Add("@lname", mstrLastName)
                    .Parameters.Add("@fname", mstrFirstName)
                    .Parameters.Add("@title", mstrTitle)
                    .Parameters.Add("@courtesy", mstrTitleOfCourtesy)
                    .Parameters.Add("@birth", mdteBirthDate)
                    .Parameters.Add("@hire", mdteHireDate)
                    .Parameters.Add("@address", mstrAddress)
                    .Parameters.Add("@city", mstrCity)
                    .Parameters.Add("@region", mstrRegion)
                    .Parameters.Add("@postal", mstrPostalCode)
                    .Parameters.Add("@country", mstrCountry)
                    .Parameters.Add("@phone", mstrHomePhone)
                    If mstrExtension = "" Then
                        .Parameters.Add("@extension", DBNull.Value)
                    Else
                        .Parameters.Add("@extension", mstrExtension)
                    End If
                    If mbytPhoto Is Nothing Then
                        .Parameters.Add("@photo", DBNull.Value)
                    Else
                        prm = New SqlParameter("@photo", SqlDbType.Image, _
                                                mbytPhoto.Length, ParameterDirection.Input, True, _
                                                0, 0, Nothing, DataRowVersion.Current, _
                                                mbytPhoto)
                        cmd.Parameters.Add(prm)
                    End If

                    If mstrNotes = "" Then
                        .Parameters.Add("@notes", DBNull.Value)
                    Else
                        .Parameters.Add("@notes", mstrNotes)
                    End If
                    If mintReportsTo = 0 Then
                        .Parameters.Add("@reports", DBNull.Value)
                    Else
                        .Parameters.Add("@reports", mintReportsTo)
                    End If
                    If mstrPhotoPath = "" Then
                        .Parameters.Add("@photopath", DBNull.Value)
                    Else
                        .Parameters.Add("@photopath", mstrPhotoPath)
                    End If
                    cmd.Parameters.Add("@new_id", intID).Direction = ParameterDirection.Output
                End With

                cmd.ExecuteNonQuery()

                mintEmployeeID = Convert.ToInt32(cmd.Parameters.Item("@new_id").Value)

                If intTempID > 0 Then
                    cmd = New SqlCommand
                    With cmd
                        .Connection = cn
                        .Transaction = trans
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "usp_employee_territory_delete"
                        .Parameters.Add("@employee_id", mintEmployeeID)
                        .ExecuteNonQuery()
                    End With
                    cmd = Nothing
                End If

                If Not sEmployee.Territories Is Nothing Then
                    For i = 0 To sEmployee.Territories.Length - 1
                        cmd = New SqlCommand
                        With cmd
                            .Connection = cn
                            .Transaction = trans
                            .CommandType = CommandType.StoredProcedure
                            .CommandText = "usp_employee_territory_insert"
                            .Parameters.Add("@employee_id", mintEmployeeID)
                            .Parameters.Add("@territory_id", sEmployee.Territories(i))
                            .ExecuteNonQuery()
                        End With
                        cmd = Nothing
                    Next
                End If

                intID = mintEmployeeID
                trans.Commit()
                trans = Nothing
                cn.Close()
            Catch exc As Exception
                trans.Rollback()
                trans = Nothing
                cn.Close()
                Throw exc
            End Try
        Else
            Return mobjBusErr
        End If
    End Function
End Class
