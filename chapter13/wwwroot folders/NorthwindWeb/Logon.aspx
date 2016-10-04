<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Logon.aspx.vb" Inherits="NorthwindWeb.Logon"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:label id="Label1" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
				Font-Size="Large">Logon Form</asp:label><asp:label id="lblMessage" style="Z-INDEX: 108; LEFT: 24px; POSITION: absolute; TOP: 208px"
				runat="server" Width="272px"></asp:label><asp:button id="btnLogon" style="Z-INDEX: 107; LEFT: 32px; POSITION: absolute; TOP: 152px" runat="server"
				Text="Logon"></asp:button><asp:button id="btnNewUser" style="Z-INDEX: 106; LEFT: 128px; POSITION: absolute; TOP: 152px"
				runat="server" Text="New User"></asp:button><asp:textbox id="txtPassword" style="Z-INDEX: 105; LEFT: 144px; POSITION: absolute; TOP: 104px"
				runat="server" TextMode="Password"></asp:textbox><asp:textbox id="txtUserName" style="Z-INDEX: 104; LEFT: 144px; POSITION: absolute; TOP: 64px"
				runat="server" Width="144px"></asp:textbox><asp:label id="Label3" style="Z-INDEX: 103; LEFT: 24px; POSITION: absolute; TOP: 104px" runat="server"
				Width="96px">Password:</asp:label><asp:label id="Label2" style="Z-INDEX: 102; LEFT: 24px; POSITION: absolute; TOP: 64px" runat="server"
				Width="88px">User Name:</asp:label></form>
	</body>
</HTML>
