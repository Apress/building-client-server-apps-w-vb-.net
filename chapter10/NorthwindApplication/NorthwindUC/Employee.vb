Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Structures
Imports NorthwindTraders.NorthwindShared.Interfaces
Imports BusinessRules.Errors
Imports NorthwindTraders.NorthwindShared
Imports NorthwindTraders.NorthwindShared.Errors

Public Class Employee
    Inherits BusinessBase

    Private Const REMENTRY As String = "EmployeeDC.rem"
    Public Event Errs(ByVal obj As BusinessErrors)
    Private msEmployee As structEmployee

#Region " Private Attributes"
    Private mintEmployeeID As Integer = 0
    Private mstrLastName As String
    Private mstrFirstName As String
    Private mstrTitle As String
    Private mstrTitleOfCourtesy As String
    Private mdteBirthDate As Date = Now
    Private mdteHireDate As Date = Now
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
    Private mobjReportsTo As Employee
    Private mobjTerritoryMgr As TerritoryMgr

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

                If mstrLastName <> Value Then
                    mstrLastName = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Last Name", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Last Name", True)
                mstrLastName = Value
                mblnDirty = True
                Throw exc
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

                If mstrFirstName <> Value Then
                    mstrFirstName = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("First Name", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("First Name", True)
                mstrFirstName = Value
                mblnDirty = True
                Throw exc
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

                If mstrTitle <> Value Then
                    mstrTitle = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Title", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Title", True)
                mstrTitle = Value
                mblnDirty = True
                Throw exc
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

                If mstrTitleOfCourtesy <> Value Then
                    mstrTitleOfCourtesy = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Title Of Courtesy", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Title Of Courtesy", True)
                mstrTitleOfCourtesy = Value
                mblnDirty = True
                Throw exc
            End Try
        End Set
    End Property

    Public Property BirthDate() As Date
        Get
            Return mdteBirthDate
        End Get
        Set(ByVal Value As Date)
            Try
                If mdteBirthDate <> Value Then
                    mdteBirthDate = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Birth Date", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Birth Date", True)
                mdteBirthDate = Value
                mblnDirty = True
                Throw exc
            End Try

        End Set
    End Property

    Public Property HireDate() As Date
        Get
            Return mdteHireDate
        End Get
        Set(ByVal Value As Date)
            Try
                If mdteHireDate <> Value Then
                    mdteHireDate = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Hire Date", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Hire Date", True)
                mdteHireDate = Value
                mblnDirty = True
                Throw exc
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

                If mstrAddress <> Value Then
                    mstrAddress = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Address", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Address", True)
                mstrAddress = Value
                mblnDirty = True
                Throw exc
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

                If mstrCity <> Value Then
                    mstrCity = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("City", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("City", True)
                mstrCity = Value
                mblnDirty = True
                Throw exc
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
                If Value Is Nothing Then
                    Exit Property
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    mstrRegion = Nothing
                    Exit Property
                End If

                'Test for max length
                If Value.Length > 15 Then
                    Throw New MaximumLengthException(15)
                End If

                If mstrRegion <> Value Then
                    mstrRegion = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Region", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Region", True)
                mstrRegion = Value
                mblnDirty = True
                Throw exc
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

                If mstrPostalCode <> Value Then
                    mstrPostalCode = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Postal Code", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Postal Code", True)
                mstrPostalCode = Value
                mblnDirty = True
                Throw exc
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

                If mstrCountry <> Value Then
                    mstrCountry = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Country", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Country", True)
                mstrCountry = Value
                mblnDirty = True
                Throw exc
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

                If mstrHomePhone <> Value Then
                    mstrHomePhone = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Home Phone", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Home Phone", True)
                mstrHomePhone = Value
                mblnDirty = True
                Throw exc
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
                If Value Is Nothing Then
                    Exit Property
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    mstrExtension = Nothing
                    Exit Property
                End If

                'Test for max length
                If Value.Length > 4 Then
                    Throw New MaximumLengthException(4)
                End If

                If mstrExtension <> Value Then
                    mstrExtension = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Extension", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Extension", True)
                mstrExtension = Value
                mblnDirty = True
                Throw exc
            End Try
        End Set
    End Property

    Public Property Photo() As Byte()
        Get
            Return mbytPhoto
        End Get
        Set(ByVal Value As Byte())
            If Not mbytPhoto Is Value Then
                mbytPhoto = Value
                If Not Loading Then
                    mblnDirty = True
                End If
            End If
        End Set
    End Property

    Public Property Notes() As String
        Get
            Return mstrNotes
        End Get
        Set(ByVal Value As String)
            If mstrNotes <> Value Then
                mstrNotes = Value
                If Not Loading Then
                    mblnDirty = True
                End If
            End If
        End Set
    End Property

    Public Property PhotoPath() As String
        Get
            Return mstrPhotoPath
        End Get
        Set(ByVal Value As String)
            Try
                'Test for null value
                If Value Is Nothing Then
                    Exit Property
                End If

                'Test for empty string
                If Value.Length = 0 Then
                    mstrPhotoPath = Nothing
                    Exit Property
                End If

                'Test for max length
                If Value.Length > 255 Then
                    Throw New MaximumLengthException(255)
                End If

                If FileSystem.FileLen(Value) > 0 Then
                    'do nothing, if the file is not found, a
                    'FileNotFoundException will be thrown automatically
                End If

                If mstrPhotoPath <> Value Then
                    mstrPhotoPath = Value
                    If Not Loading Then
                        mobjRules.BrokenRule("Photo Path", False)
                        mblnDirty = True
                    End If
                End If
            Catch exc As Exception
                mobjRules.BrokenRule("Photo Path", True)
                mstrPhotoPath = Value
                mblnDirty = True
                Throw exc
            End Try
        End Set
    End Property

    Public Property ReportsTo() As Employee
        Get
            Return mobjReportsTo
        End Get
        Set(ByVal Value As Employee)
            If Not mobjReportsTo Is Value Then
                mobjReportsTo = Value
                If Not Loading Then
                    mblnDirty = True
                End If
            End If
        End Set
    End Property

    Public ReadOnly Property Territories() As TerritoryMgr
        Get
            Return mobjTerritoryMgr
        End Get
    End Property

#End Region

    Public Shadows Function IsDirty() As Boolean
        Dim i As Integer
        Dim dictEnt As DictionaryEntry
        Dim blnFound As Boolean

        If Not mblndirty Then
            If Not msEmployee.Territories Is Nothing Then
                If mobjTerritoryMgr.Count <> msEmployee.Territories.Length Then
                    mblndirty = True
                Else
                    For Each dictEnt In mobjTerritoryMgr
                        Dim objTerritory As Territory = _
                        CType(dictEnt.Value, Territory)
                        blnFound = False
                        For i = 0 To msEmployee.Territories.Length - 1
                            If objTerritory.TerritoryID = _
                            msEmployee.Territories(i) Then
                                blnFound = True
                                Exit For
                            End If
                        Next
                        If Not blnFound Then
                            mblndirty = True
                            Exit For
                        End If
                    Next
                End If
            End If
        End If

        Return mblndirty
    End Function

    Public Sub New()
        MyBase.New(REMENTRY)
        mobjRules = New BrokenRules
        mobjTerritoryMgr = New TerritoryMgr(False)
        With mobjrules
            .BrokenRule("Last Name", True)
            .BrokenRule("First Name", True)
            .BrokenRule("Title", True)
            .BrokenRule("Address", True)
            .BrokenRule("City", True)
            .BrokenRule("Postal Code", True)
            .BrokenRule("Country", True)
            .BrokenRule("Home Phone", True)
        End With
    End Sub

    Public Sub New(ByVal intID As Integer)
        MyBase.New(REMENTRY)
        mobjRules = New BrokenRules
        mintEmployeeID = intID
        mobjTerritoryMgr = New TerritoryMgr(False)
    End Sub

    Public Overrides Function ToString() As String
        Return mstrLastName & ", " & mstrFirstName
    End Function

    Public Sub LoadRecord(ByRef objTerritoryMgr As TerritoryMgr)
        Dim objIEmp As IEmployee

        objIEmp = CType(Activator.GetObject(GetType(IEmployee), _
        AppConstants.REMOTEOBJECTS & REMENTRY), IEmployee)
        msEmployee = objIEmp.LoadRecord(mintEmployeeID)
        objIEmp = Nothing

        LoadObject(objTerritoryMgr)
    End Sub

    Private Sub LoadObject(ByRef objTerritoryMgr As TerritoryMgr)
        Dim i As Integer

        With msEmployee
            Me.mintEmployeeID = .EmployeeID
            Me.mstrLastName = .LastName
            Me.mstrFirstName = .FirstName
            Me.mstrTitle = .Title
            Me.mstrTitleOfCourtesy = .TitleOfCourtesy
            Me.mdteBirthDate = .BirthDate
            Me.mdteHireDate = .HireDate
            Me.mstrAddress = .Address
            Me.mstrCity = .City
            Me.mstrRegion = .Region
            Me.mstrPostalCode = .PostalCode
            Me.mstrCountry = .Country
            Me.mstrHomePhone = .HomePhone
            Me.mstrExtension = .Extension
            Me.mbytPhoto = .Photo
            Me.mstrNotes = .Notes
            Me.mstrPhotoPath = .PhotoPath

            If .ReportsTo > 0 Then
                mobjReportsTo = New Employee(.ReportsTo)
                mobjReportsTo.FirstName = .ReportsToFirstName
                mobjReportsTo.LastName = .ReportsToLastName
            End If

            mobjTerritoryMgr.Clear()
            If Not .Territories Is Nothing Then
                For i = 0 To .Territories.Length - 1
                    mobjTerritoryMgr.Add((objTerritoryMgr.Item(.Territories(i))))
                Next
            End If
        End With
    End Sub

    Public Sub Rollback(ByRef objTerritoryMgr As TerritoryMgr)
        LoadObject(objTerritoryMgr)
    End Sub

    Public Sub Delete()
        Dim objIEmp As IEmployee

        objIEmp = CType(Activator.GetObject(GetType(IEmployee), _
        AppConstants.REMOTEOBJECTS & REMENTRY), IEmployee)
        objIEmp.Delete(mintEmployeeID)
        objIEmp = Nothing
    End Sub

    Public Sub Save()
        Dim objBusErr As BusinessErrors

        If mobjRules.Count = 0 Then
            If IsDirty() = True Then
                Dim objIEmp As IEmployee
                Dim intID As Integer
                Dim sEmployee As structEmployee
                Dim i As Integer
                Dim objTerritory As Territory
                Dim dictEnt As DictionaryEntry

                intID = mintEmployeeID

                With sEmployee
                    .EmployeeID = mintEmployeeID
                    .LastName = Me.mstrLastName
                    .FirstName = Me.mstrFirstName
                    .Title = Me.mstrTitle
                    .TitleOfCourtesy = Me.mstrTitleOfCourtesy
                    .BirthDate = Me.mdteBirthDate
                    .HireDate = Me.mdteHireDate
                    .Address = Me.mstrAddress
                    .City = Me.mstrCity
                    .Region = Me.mstrRegion
                    .PostalCode = Me.mstrPostalCode
                    .Country = Me.mstrCountry
                    .HomePhone = Me.mstrHomePhone
                    .Extension = Me.mstrExtension
                    .Photo = Me.mbytPhoto
                    .Notes = Me.mstrNotes
                    .PhotoPath = Me.mstrPhotoPath
                    If Not mobjReportsTo Is Nothing Then
                        .ReportsTo = mobjReportsTo.EmployeeID
                    End If

                    ReDim .Territories(mobjTerritoryMgr.Count - 1)
                    i = 0
                    For Each dictEnt In mobjTerritoryMgr
                        objTerritory = CType(dictEnt.Value, Territory)
                        .Territories(i) = objTerritory.TerritoryID
                        i += 1
                    Next
                End With

                objIEmp = CType(Activator.GetObject(GetType(IEmployee), _
                AppConstants.REMOTEOBJECTS & REMENTRY), IEmployee)

                objBusErr = objIEmp.Save(sEmployee, mintEmployeeID)
                objIEmp = Nothing

                If Not objBusErr Is Nothing Then
                    RaiseEvent Errs(objBusErr)
                Else
                    mblndirty = False
                    msEmployee = sEmployee
                    If intID = 0 Then
                        CallChangedEvent(Me, New _
                        ChangedEventArgs(ChangedEventArgs.eChange.Added))
                    Else
                        CallChangedEvent(Me, New _
                        ChangedEventArgs(ChangedEventArgs.eChange.Updated))
                    End If
                End If
            End If
        End If
    End Sub

End Class

Public Class EmployeeMgr
    Inherits System.Collections.DictionaryBase

    Private Shared mobjEmployeeMgr As EmployeeMgr

    Public Shared Function GetInstance() As EmployeeMgr
        If mobjEmployeeMgr Is Nothing Then
            mobjEmployeeMgr = New EmployeeMgr
        End If
        Return mobjEmployeeMgr
    End Function

    Protected Sub New()
        Load()
    End Sub

    Public Sub Add(ByVal obj As Employee)
        dictionary.Add(obj.EmployeeID, obj)
    End Sub

    Public Function Item(ByVal Key As Object) As Employee
        Return CType(dictionary.Item(Key), Employee)
    End Function

    Public Sub Remove(ByVal Key As Object)
        dictionary.Remove(Key)
    End Sub

    Private Sub Load()
        Dim objIEmployee As IEmployee
        Dim dRow As DataRow
        Dim ds As DataSet

        'Obtain a reference to the remote object
        objIEmployee = CType(Activator.GetObject(GetType(IEmployee), _
        AppConstants.REMOTEOBJECTS & "EmployeeDC.rem"), IEmployee)
        ds = objIEmployee.LoadProxy()
        objIEmployee = Nothing

        'Loop through the dataset adding region objects to the collection
        For Each dRow In ds.Tables(0).Rows
            'Assign the ID in the constructor since we can't assign it 
            'afterwards
            Dim objEmployee As New _
            Employee(Convert.ToInt32(dRow.Item("EmployeeID")))
            With objEmployee
                'Set the loading flag so the object isn't marked as edited
                .Loading = True
                .LastName = Convert.ToString(dRow.Item("LastName"))
                .FirstName = Convert.ToString(dRow.Item("FirstName"))
                .Title = Convert.ToString(dRow.Item("Title"))
                .Loading = False
            End With
            'Add the object to the collection
            Me.Add(objEmployee)
        Next

        ds = Nothing
    End Sub

    Public Sub Refresh()
        dictionary.Clear()
        Load()
    End Sub

End Class
