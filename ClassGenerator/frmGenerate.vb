Option Explicit On 
Option Strict On

Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms


Public Class frmGenerate
    Inherits System.Windows.Forms.Form

    Private mstrCN As String
    Private mstrCatalog As String

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents gBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnTest As System.Windows.Forms.Button
    Friend WithEvents cboCatalog As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents gBoxTop As System.Windows.Forms.GroupBox
    Friend WithEvents gBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnRight As System.Windows.Forms.Button
    Friend WithEvents btnAllRight As System.Windows.Forms.Button
    Friend WithEvents btnAllLeft As System.Windows.Forms.Button
    Friend WithEvents btnLeft As System.Windows.Forms.Button
    Friend WithEvents lstSelectedTables As System.Windows.Forms.ListBox
    Friend WithEvents lstAvailableTables As System.Windows.Forms.ListBox
    Friend WithEvents gbox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents clbGenerated As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.gBox1 = New System.Windows.Forms.GroupBox()
        Me.btnTest = New System.Windows.Forms.Button()
        Me.cboCatalog = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUser = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.gBoxTop = New System.Windows.Forms.GroupBox()
        Me.gBox2 = New System.Windows.Forms.GroupBox()
        Me.btnAllLeft = New System.Windows.Forms.Button()
        Me.btnLeft = New System.Windows.Forms.Button()
        Me.btnAllRight = New System.Windows.Forms.Button()
        Me.btnRight = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lstSelectedTables = New System.Windows.Forms.ListBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lstAvailableTables = New System.Windows.Forms.ListBox()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.gbox3 = New System.Windows.Forms.GroupBox()
        Me.clbGenerated = New System.Windows.Forms.CheckedListBox()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.gBox1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.gBoxTop.SuspendLayout()
        Me.gBox2.SuspendLayout()
        Me.gbox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(325, 385)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.TabIndex = 6
        Me.btnNext.Text = "Next >"
        '
        'gBox1
        '
        Me.gBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnTest, Me.cboCatalog, Me.Label4, Me.GroupBox1, Me.Label1, Me.txtServer})
        Me.gBox1.Location = New System.Drawing.Point(-5, 50)
        Me.gBox1.Name = "gBox1"
        Me.gBox1.Size = New System.Drawing.Size(425, 325)
        Me.gBox1.TabIndex = 14
        Me.gBox1.TabStop = False
        '
        'btnTest
        '
        Me.btnTest.Location = New System.Drawing.Point(240, 290)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(100, 23)
        Me.btnTest.TabIndex = 19
        Me.btnTest.Text = "Test Connection"
        '
        'cboCatalog
        '
        Me.cboCatalog.Location = New System.Drawing.Point(135, 245)
        Me.cboCatalog.Name = "cboCatalog"
        Me.cboCatalog.Size = New System.Drawing.Size(180, 21)
        Me.cboCatalog.TabIndex = 18
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(80, 250)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 15)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Catalog"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtPassword, Me.txtUser, Me.Label3, Me.Label2, Me.RadioButton2, Me.RadioButton1})
        Me.GroupBox1.Location = New System.Drawing.Point(80, 75)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(255, 150)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Authentication"
        '
        'txtPassword
        '
        Me.txtPassword.Enabled = False
        Me.txtPassword.Location = New System.Drawing.Point(95, 115)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(135, 20)
        Me.txtPassword.TabIndex = 5
        Me.txtPassword.Text = ""
        '
        'txtUser
        '
        Me.txtUser.Enabled = False
        Me.txtUser.Location = New System.Drawing.Point(95, 90)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(135, 20)
        Me.txtUser.TabIndex = 4
        Me.txtUser.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(25, 115)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 15)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Password:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(25, 95)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "User Name:"
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(25, 65)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(100, 15)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "SQL Server"
        '
        'RadioButton1
        '
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(25, 35)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(100, 15)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Windows"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(70, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 15)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Server:"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(120, 30)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(135, 20)
        Me.txtServer.TabIndex = 14
        Me.txtServer.Text = "localhost"
        '
        'gBoxTop
        '
        Me.gBoxTop.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.gBoxTop.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label7})
        Me.gBoxTop.Location = New System.Drawing.Point(-5, -5)
        Me.gBoxTop.Name = "gBoxTop"
        Me.gBoxTop.Size = New System.Drawing.Size(420, 65)
        Me.gBoxTop.TabIndex = 15
        Me.gBoxTop.TabStop = False
        '
        'gBox2
        '
        Me.gBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnAllLeft, Me.btnLeft, Me.btnAllRight, Me.btnRight, Me.Label6, Me.lstSelectedTables, Me.Label5, Me.lstAvailableTables})
        Me.gBox2.Location = New System.Drawing.Point(-10, 55)
        Me.gBox2.Name = "gBox2"
        Me.gBox2.Size = New System.Drawing.Size(425, 320)
        Me.gBox2.TabIndex = 16
        Me.gBox2.TabStop = False
        Me.gBox2.Visible = False
        '
        'btnAllLeft
        '
        Me.btnAllLeft.Location = New System.Drawing.Point(200, 230)
        Me.btnAllLeft.Name = "btnAllLeft"
        Me.btnAllLeft.Size = New System.Drawing.Size(30, 30)
        Me.btnAllLeft.TabIndex = 7
        Me.btnAllLeft.Text = "<<"
        '
        'btnLeft
        '
        Me.btnLeft.Location = New System.Drawing.Point(200, 195)
        Me.btnLeft.Name = "btnLeft"
        Me.btnLeft.Size = New System.Drawing.Size(30, 30)
        Me.btnLeft.TabIndex = 6
        Me.btnLeft.Text = "<"
        '
        'btnAllRight
        '
        Me.btnAllRight.Location = New System.Drawing.Point(200, 155)
        Me.btnAllRight.Name = "btnAllRight"
        Me.btnAllRight.Size = New System.Drawing.Size(30, 30)
        Me.btnAllRight.TabIndex = 5
        Me.btnAllRight.Text = ">>"
        '
        'btnRight
        '
        Me.btnRight.Location = New System.Drawing.Point(200, 120)
        Me.btnRight.Name = "btnRight"
        Me.btnRight.Size = New System.Drawing.Size(30, 30)
        Me.btnRight.TabIndex = 4
        Me.btnRight.Text = ">"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(240, 50)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(160, 20)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "Selected Tables:"
        '
        'lstSelectedTables
        '
        Me.lstSelectedTables.Location = New System.Drawing.Point(240, 70)
        Me.lstSelectedTables.Name = "lstSelectedTables"
        Me.lstSelectedTables.Size = New System.Drawing.Size(160, 238)
        Me.lstSelectedTables.Sorted = True
        Me.lstSelectedTables.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(25, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(165, 20)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Available Tables:"
        '
        'lstAvailableTables
        '
        Me.lstAvailableTables.Location = New System.Drawing.Point(25, 70)
        Me.lstAvailableTables.Name = "lstAvailableTables"
        Me.lstAvailableTables.Size = New System.Drawing.Size(165, 238)
        Me.lstAvailableTables.Sorted = True
        Me.lstAvailableTables.TabIndex = 0
        '
        'btnBack
        '
        Me.btnBack.Enabled = False
        Me.btnBack.Location = New System.Drawing.Point(240, 385)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.TabIndex = 17
        Me.btnBack.Text = "< Back"
        '
        'gbox3
        '
        Me.gbox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.clbGenerated, Me.btnGenerate})
        Me.gbox3.Location = New System.Drawing.Point(-10, 50)
        Me.gbox3.Name = "gbox3"
        Me.gbox3.Size = New System.Drawing.Size(425, 325)
        Me.gbox3.TabIndex = 18
        Me.gbox3.TabStop = False
        Me.gbox3.Visible = False
        '
        'clbGenerated
        '
        Me.clbGenerated.Location = New System.Drawing.Point(180, 55)
        Me.clbGenerated.Name = "clbGenerated"
        Me.clbGenerated.Size = New System.Drawing.Size(225, 244)
        Me.clbGenerated.TabIndex = 2
        '
        'btnGenerate
        '
        Me.btnGenerate.Location = New System.Drawing.Point(45, 55)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(105, 30)
        Me.btnGenerate.TabIndex = 1
        Me.btnGenerate.Text = "Generate Classes"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(25, 20)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(340, 35)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Select Database Connection"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(5, 385)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 20
        Me.btnCancel.Text = "Cancel"
        '
        'frmGenerate
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(407, 413)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCancel, Me.gbox3, Me.btnBack, Me.btnNext, Me.gBox2, Me.gBoxTop, Me.gBox1})
        Me.Name = "frmGenerate"
        Me.Text = "Class Generator Wizard"
        Me.gBox1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.gBoxTop.ResumeLayout(False)
        Me.gBox2.ResumeLayout(False)
        Me.gbox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Connection Frame "

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged

        Dim rButton As RadioButton
        rButton = CType(sender, RadioButton)
        Select Case rButton.Name
            Case "RadioButton1"
                txtUser.Enabled = False
                txtPassword.Enabled = False
            Case "RadioButton2"
                txtUser.Enabled = True
                txtPassword.Enabled = True
        End Select
    End Sub

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles btnTest.Click

        Dim cn As SqlConnection = New SqlConnection()

        CreateConnectionString(False)

        Try
            cn.ConnectionString = mstrCN
            cn.Open()
            MessageBox.Show("Test Connection Succeeded!", "Connection", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch exc As Exception
            MessageBox.Show(exc.Message, "Connection", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            cn = Nothing
        End Try
    End Sub

    Private Sub cboCatalog_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles cboCatalog.DropDown

        Dim cn As SqlConnection = New SqlConnection()
        Dim strSQL As String = "Select name from sysdatabases order by name"
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
        Dim dReader As SqlDataReader

        Try
            Cursor.Current = Cursors.WaitCursor

            CreateConnectionString(True)

            cn.ConnectionString = mstrCN
            cn.Open()
            dReader = cmd.ExecuteReader
            cmd = Nothing

            While dReader.Read
                cboCatalog.Items.Add(dReader.GetString(dReader.GetOrdinal("name")))
            End While

        Catch exc As Exception
            MessageBox.Show(exc.Message)
        Finally
            dReader = Nothing
            cn = Nothing
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub CreateConnectionString(ByVal blnDrop As Boolean)
        mstrCatalog = "Master"

        If cboCatalog.Text <> "" And Not blnDrop Then
            mstrCatalog = cboCatalog.Text
        End If

        If RadioButton1.Checked Then
            mstrCN = "Server=" & txtServer.Text & ";Integrated Security=SSPI;" _
                & "Initial Catalog=" & mstrCatalog
        Else
            mstrCN = "Server=" & txtServer.Text & ";uid=" & txtUser.Text _
                & ";pwd=" & txtPassword.Text & ";Initial Catalog=" & mstrCatalog
        End If
    End Sub
#End Region

#Region " Select Table Frame "

    Private Sub LoadTables()
        Dim cn As SqlConnection = New SqlConnection()
        Dim strSQL As String = "select * from sysobjects where xtype " _
            & "= 'U' and name <> 'dtproperties' order by name"
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
        Dim dReader As SqlDataReader

        Cursor.Current = Cursors.WaitCursor

        cn.ConnectionString = mstrCN
        cn.Open()
        dReader = cmd.ExecuteReader
        cmd = Nothing

        While dReader.Read
            lstAvailableTables.Items.Add(dReader.GetString(dReader.GetOrdinal("name")))
        End While

        dReader = Nothing
        cn = Nothing

        Cursor.Current = Cursors.Default
    End Sub

    Private Sub MoveTables_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btnRight.Click, btnAllRight.Click, _
        btnLeft.Click, btnAllLeft.Click

        Dim btn As Button = CType(sender, Button)
        Dim i As Integer = 0

        Select Case btn.Text
            Case ">"
                lstSelectedTables.Items.Add(lstAvailableTables.SelectedItem)
                lstAvailableTables.Items.RemoveAt(lstAvailableTables.SelectedIndex)
            Case ">>"
                For i = 0 To lstAvailableTables.Items.Count - 1
                    lstSelectedTables.Items.Add(lstAvailableTables.Items(0))
                    lstAvailableTables.Items.RemoveAt(0)
                Next i
            Case "<"
                lstAvailableTables.Items.Add(lstSelectedTables.SelectedItem)
                lstSelectedTables.Items.RemoveAt(lstSelectedTables.SelectedIndex)
            Case "<<"
                For i = 0 To lstSelectedTables.Items.Count - 1
                    lstAvailableTables.Items.Add(lstSelectedTables.Items(0))
                    lstSelectedTables.Items.RemoveAt(0)
                Next
        End Select
    End Sub
#End Region

#Region " Generate Frame "

    Private Sub btnGenerate_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnGenerate.Click
        Dim i As Integer
        Dim cn As SqlConnection = New SqlConnection(mstrCN)

        cn.Open()

        'Perform the class generation here, as we finish, check the box
        For i = 0 To clbGenerated.Items.Count - 1
            'Create the code module
            Dim fs As New IO.FileStream("C:\" & clbGenerated.Items(i).ToString & ".vb", _
                IO.FileMode.Create)
            Dim sw As New IO.StreamWriter(fs)

            'Add the header information to the module
            sw.WriteLine("Option Explicit On")
            sw.WriteLine("Option Strict On")
            sw.WriteLine()
            sw.WriteLine("Public Class {0}", clbGenerated.Items(i).ToString)
            sw.WriteLine("")

            'Get the list of columns in the current table
            Dim strSQL As String = "select 	a.name colname, a.length, c.name typename " _
                & "from syscolumns a " _
                & "join	sysobjects b on a.id = b.id " _
                & "join	systypes c on a.xtype = c.xtype " _
                & "where b.name = '" & clbGenerated.Items(i).ToString & "' " _
                & "and c.name <> 'sysname' order by a.number"
            Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
            Dim ds As New DataSet()
            Dim da As New SqlClient.SqlDataAdapter(cmd)
            Dim dRow As DataRow
            Dim strPrefix As String
            Dim strType As String
            Dim strColumn As String
            Dim chrColumn() As Char
            Dim j As Integer

            da.Fill(ds)

            sw.WriteLine("#Region " & Chr(34) & " Private Member Variables" & Chr(34))
            sw.WriteLine()

            'Loop through the columns and create the private module level variables
            For Each dRow In ds.Tables(0).Rows
                'Perform a select case on the data type to get the right
                'variable prefix
                Select Case Convert.ToString(dRow.Item("typename"))
                    Case "tinyint", "smallint", "int", "bigint"
                        strPrefix = "mint"
                        strType = "Integer"
                    Case "smalldatetime", "datetime"
                        strPrefix = "mdte"
                        strType = "Date"
                    Case "real", "decimal", "money", "float", "numeric", "smallmoney"
                        strPrefix = "mdec"
                        strType = "Decimal"
                    Case "varchar", "nvarchar", "text", "ntext"
                        strPrefix = "mstr"
                        strType = "String"
                    Case "char", "nchar"
                        If Convert.ToInt32(dRow.Item("length")) > 1 Then
                            strPrefix = "mstr"
                            strType = "String"
                        Else
                            strPrefix = "mchr"
                            strType = "Char"
                        End If
                    Case "timestamp"
                        strPrefix = "mlng"
                        strType = "Long"
                    Case "image"
                        strPrefix = "mbyt"
                        strType = "Byte()"
                End Select

                strColumn = Convert.ToString(dRow.Item("colname"))
                chrColumn = strColumn.ToCharArray

                chrColumn(0) = Char.ToUpper(chrColumn(0))

                For j = 1 To chrColumn.Length - 1
                    If chrColumn(j) = "_"c Then
                        chrColumn(j + 1) = Char.ToUpper(chrColumn(j + 1))
                    End If
                Next

                strColumn = ""

                For j = 0 To chrColumn.Length - 1
                    If chrColumn(j) <> "_"c Then
                        strColumn += chrColumn(j).ToString
                    End If
                Next

                sw.WriteLine("Private {0}" & strColumn & " As {1}", strPrefix, strType)
            Next

            sw.WriteLine("")
            sw.WriteLine("#End Region")
            sw.WriteLine("")

            sw.WriteLine("#Region " & Chr(34) & " Public Accessors" & Chr(34))
            sw.WriteLine()

            'Loop through the columns and create the public accessors
            For Each dRow In ds.Tables(0).Rows
                'Perform a select case on the data type to get the right variable prefix
                Select Case Convert.ToString(dRow.Item("typename"))
                    Case "tinyint", "smallint", "int", "bigint"
                        strPrefix = "mint"
                        strType = "Integer"
                    Case "smalldatetime", "datetime"
                        strPrefix = "mdte"
                        strType = "Date"
                    Case "real", "decimal", "money", "float", "numeric", "smallmoney"
                        strPrefix = "mdec"
                        strType = "Decimal"
                    Case "varchar", "nvarchar", "text", "ntext"
                        strPrefix = "mstr"
                        strType = "String"
                    Case "char", "nchar"
                        If Convert.ToInt32(dRow.Item("length")) > 1 Then
                            strPrefix = "mstr"
                            strType = "String"
                        Else
                            strPrefix = "mchr"
                            strType = "Char"
                        End If
                    Case "timestamp"
                        strPrefix = "mlng"
                        strType = "Long"
                    Case "image"
                        strPrefix = "mbyt"
                        strType = "Byte()"
                End Select

                strColumn = Convert.ToString(dRow.Item("colname"))
                chrColumn = strColumn.ToCharArray

                chrColumn(0) = Char.ToUpper(chrColumn(0))

                For j = 1 To chrColumn.Length - 1
                    If chrColumn(j) = "_"c Then
                        chrColumn(j + 1) = Char.ToUpper(chrColumn(j + 1))
                    End If
                Next

                strColumn = ""

                For j = 0 To chrColumn.Length - 1
                    If chrColumn(j) <> "_"c Then
                        strColumn += chrColumn(j).ToString
                    End If
                Next

                sw.WriteLine("Public Property {0}() As {1}", strColumn, strType)
                sw.WriteLine("     Get")
                sw.WriteLine("          Return {0}{1}", strPrefix, strColumn)
                sw.WriteLine("     End Get")
                sw.WriteLine("     Set(ByVal Value As {0})", strType)
                sw.WriteLine("          {0}{1} = Value", strPrefix, strColumn)
                sw.WriteLine("     End Set")
                sw.WriteLine("End Property")
                sw.WriteLine("")
            Next

            sw.WriteLine("#End Region")
            sw.WriteLine("")

            sw.WriteLine("End Class")
            sw.Close()
            fs.Close()

            clbGenerated.SetItemChecked(i, True)
        Next
        MessageBox.Show("Done")
    End Sub

#End Region

    Private Sub LoadItemsToGenerate()
        Dim i As Integer

        'Move all of the selected tables to this list box
        For i = 0 To lstSelectedTables.Items.Count - 1
            clbGenerated.Items.Add(lstSelectedTables.Items(i))
        Next
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btnNext.Click

        If btnNext.Text = "Finished" Then
            Close()
        End If

        If gBox1.Visible Then
            If cboCatalog.Text = "" Then
                MessageBox.Show("Select a database first.", "No Database", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                CreateConnectionString(False)
                LoadTables()
                Label7.Text = "Select Tables To Generate Classes From"
                gBox2.Visible = True
                gBox1.Visible = False
                btnBack.Enabled = True
            End If
        Else
            If gBox2.Visible Then
                gbox3.Visible = True
                gBox2.Visible = False
                btnNext.Text = "Finished"
                LoadItemsToGenerate()
            End If
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        If gBox2.Visible Then
            Label7.Text = "Select Database Connection"
            gBox1.Visible = True
            gBox2.Visible = False
        Else
            Label7.Text = "Select Tables To Generate Classes From"
            gBox2.Visible = True
            gbox3.Visible = False
            btnNext.Text = "Next >"
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Close()
    End Sub
End Class