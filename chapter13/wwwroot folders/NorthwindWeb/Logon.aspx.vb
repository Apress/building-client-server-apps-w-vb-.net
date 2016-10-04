Option Explicit On 
Option Strict On

Imports System.Web
Imports System.Web.Security
Imports System.Security.Cryptography
Imports System.Configuration
Imports NorthwindTraders.NorthwindShared.Interfaces

Public Class Logon
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents txtUserName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtPassword As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnNewUser As System.Web.UI.WebControls.Button
    Protected WithEvents btnLogon As System.Web.UI.WebControls.Button
    Protected WithEvents lblMessage As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnNewUser_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnNewUser.Click
        Dim objSecurity As IWebUser
        Dim blnValid As Boolean
        Dim bytPassword(), bytEncrypted() As Byte
        Dim strEncryptedPassword As String
        Dim i As Integer
        Dim sec As New SHA1Managed
        Dim ue As New System.Text.UnicodeEncoding

        Try
            bytPassword = ue.GetBytes(txtPassword.Text)
            bytEncrypted = sec.ComputeHash(bytPassword)

            For i = 0 To bytEncrypted.Length - 1
                strEncryptedPassword += bytEncrypted(i).ToString
            Next

            objSecurity = CType(Activator.GetObject(GetType(IWebUser), _
            "http://localhost:80/northwind/WebUserDC.rem"), IWebUser)
            blnValid = objSecurity.AddUser(txtUserName.Text, strEncryptedPassword)
            objSecurity = Nothing
            If blnValid = False Then
                Throw New Exception("Failed to add user.")
            End If

            FormsAuthentication.SetAuthCookie(txtUserName.Text, False)
            Response.Redirect("secure/Employees.aspx")
        Catch exc As Exception
            lblMessage.Text = "An error occurred while adding user " & txtUserName.Text
        End Try
    End Sub

    Private Sub btnLogon_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnLogon.Click
        Dim objSecurity As IWebUser
        Dim blnValid As Boolean
        Dim bytPassword(), bytEncrypted() As Byte
        Dim strEncryptedPassword As String
        Dim i As Integer
        Dim sec As New SHA1Managed
        Dim ue As New System.Text.UnicodeEncoding

        bytPassword = ue.GetBytes(txtPassword.Text)
        bytEncrypted = sec.ComputeHash(bytPassword)

        For i = 0 To bytEncrypted.Length - 1
            strEncryptedPassword += bytEncrypted(i).ToString
        Next

        objSecurity = CType(Activator.GetObject(GetType(IWebUser), _
        "http://localhost:80/northwind/WebUserDC.rem"), IWebUser)

        blnValid = objSecurity.ValidateUser(txtUserName.Text, strEncryptedPassword)
        objSecurity = Nothing
        If blnValid = False Then
            Response.Redirect("AccessDenied.aspx")
        End If

        FormsAuthentication.SetAuthCookie(txtUserName.Text, False)
        Response.Redirect("secure/Employees.aspx")
    End Sub
End Class
