Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Interfaces

Public Class employees
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents hypAdd As System.Web.UI.WebControls.HyperLink

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objIEmployee As IEmployee
        Dim ds As DataSet

        'Obtain a reference to the remote object
        objIEmployee = CType(Activator.GetObject(GetType(IEmployee), _
        "http://localhost:80/northwind/EmployeeDC.rem"), IEmployee)
        ds = objIEmployee.LoadProxy()
        objIEmployee = Nothing

        DataGrid1.DataSource = ds
        DataGrid1.DataBind()
    End Sub
End Class
