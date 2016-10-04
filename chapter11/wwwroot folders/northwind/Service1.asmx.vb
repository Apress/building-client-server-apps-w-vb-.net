Option Explicit On 
Option Strict On

Imports System.Web.Services
Imports NorthwindTraders.NorthwindDC
Imports NorthwindTraders.NorthwindShared.Structures

<System.Web.Services.WebService(Namespace := "http://tempuri.org/WebService1/Service1")> _
Public Class Service1
    Inherits System.Web.Services.WebService

#Region " Web Services Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()

        'Add your own initialization code after the InitializeComponent() call

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

    <WebMethod()> Public Function GetAllEmployees() As DataSet
        Dim objEmployee As New EmployeeDC
        Dim ds As DataSet

        ds = objEmployee.LoadProxy

        objEmployee = Nothing

        Return ds
    End Function

    <WebMethod()> Public Function GetEmployeeDetails(ByVal EmployeeID _
    As Integer) As structEmployee
        Dim objEmployee As New EmployeeDC
        Dim sEmployee As structEmployee

        sEmployee = objEmployee.LoadRecord(EmployeeID)

        objEmployee = Nothing

        Return sEmployee
    End Function

End Class
