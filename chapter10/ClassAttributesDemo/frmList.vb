Option Explicit On 

Imports System.Reflection

Public Class frmList
    Inherits System.Windows.Forms.Form

    Private mobjCLMgr As ComputerListMgr
    Private mobjBKMgr As BookListMgr

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
    Friend WithEvents lvwList As System.Windows.Forms.ListView
    Friend WithEvents btnBooks As System.Windows.Forms.Button
    Friend WithEvents btnComputers As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lvwList = New System.Windows.Forms.ListView()
        Me.btnBooks = New System.Windows.Forms.Button()
        Me.btnComputers = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lvwList
        '
        Me.lvwList.Location = New System.Drawing.Point(16, 16)
        Me.lvwList.Name = "lvwList"
        Me.lvwList.Size = New System.Drawing.Size(408, 224)
        Me.lvwList.TabIndex = 0
        Me.lvwList.View = System.Windows.Forms.View.Details
        '
        'btnBooks
        '
        Me.btnBooks.Location = New System.Drawing.Point(336, 248)
        Me.btnBooks.Name = "btnBooks"
        Me.btnBooks.Size = New System.Drawing.Size(88, 23)
        Me.btnBooks.TabIndex = 1
        Me.btnBooks.Text = "Book List"
        '
        'btnComputers
        '
        Me.btnComputers.Location = New System.Drawing.Point(16, 248)
        Me.btnComputers.Name = "btnComputers"
        Me.btnComputers.Size = New System.Drawing.Size(88, 23)
        Me.btnComputers.TabIndex = 2
        Me.btnComputers.Text = "Computer List"
        '
        'frmList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(440, 278)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnComputers, Me.btnBooks, Me.lvwList})
        Me.Name = "frmList"
        Me.Text = "Class Attribute Demo"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnComputers_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnComputers.Click
        Dim t As Type = GetType(ComputerList)
        mobjCLMgr = New ComputerListMgr()

        mobjCLMgr.Add(New ComputerList("Lightning", "P4", 1.4, 699.0, "Dell"))
        mobjCLMgr.Add(New ComputerList("Thunder", "P4", 1.7, 799.0, "Dell"))
        LoadList(t, CType(mobjCLMgr, CollectionBase))
    End Sub

    Private Sub LoadList(ByVal t As Type, ByRef col As CollectionBase)
        Dim p As PropertyInfo()
        Dim i, j As Integer
        Dim SortedL As New Collections.SortedList()

        lvwList.Clear()

        p = t.GetProperties(BindingFlags.Public Or BindingFlags.Instance)

        For i = 0 To p.Length - 1
            Dim a As Object()
            a = p(i).GetCustomAttributes(False)
            If a.Length > 0 Then
                For j = 0 To a.Length - 1
                    If a(j).GetType Is GetType(ListAttribute) Then
                        Dim la As ListAttribute = CType(a(j), ListAttribute)
                        SortedL.Add(la.Column, p(i))
                    End If
                Next
            End If
        Next

        For i = 0 To SortedL.Count - 1
            Dim pi As PropertyInfo = CType(SortedL.Item(i), PropertyInfo)
            Dim a As Object = pi.GetCustomAttributes(False)
            For j = 0 To a.Length - 1
                If a(j).GetType Is GetType(ListAttribute) Then
                    Dim la As ListAttribute = CType(a(j), ListAttribute)
                    lvwList.Columns.Add(la.Heading, lvwList.Width / SortedL.Count - 2, HorizontalAlignment.Left)
                    Exit For
                End If
            Next
        Next

        Dim obj As Object
        Dim myObject() As Object

        For Each obj In col
            Dim cl As Object = Convert.ChangeType(obj, t)
            Dim k As Integer
            Dim lst As New ListViewItem()
            For i = 0 To SortedL.Count - 1
                Dim pr As PropertyInfo = CType(SortedL.Item(i), PropertyInfo)
                Dim strValue As String = ""
                strValue = Convert.ToString(t.InvokeMember(pr.Name, BindingFlags.GetProperty, _
                Nothing, cl, myObject))
                If i = 0 Then
                    lst.SubItems(i).Text = strValue
                Else
                    lst.SubItems.Add(strValue)
                End If
            Next
            lvwList.Items.Add(lst)
        Next
    End Sub

    Private Sub btnBooks_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnBooks.Click
        Dim t As Type = GetType(BookList)
        mobjBKMgr = New BookListMgr()

        mobjBKMgr.Add(New BookList("Life With .NET", "Anonymous", 49.95, "Apress"))
        mobjBKMgr.Add(New BookList("Life With Java", "Unknown", 19.95, "ABC Publishing"))
        LoadList(t, CType(mobjBKMgr, CollectionBase))
    End Sub
End Class
