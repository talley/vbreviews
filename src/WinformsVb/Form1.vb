Imports DbHelper = Microsoft.ApplicationBlocks.Data.SqlHelper
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Public Class Form1
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbltotal As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lbltotal = New System.Windows.Forms.Label
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(128, 8)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(632, 408)
        Me.DataGrid1.TabIndex = 1
        '
        'ListBox1
        '
        Me.ListBox1.Location = New System.Drawing.Point(0, 8)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(120, 407)
        Me.ListBox1.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(128, 424)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Order Total:"
        '
        'lbltotal
        '
        Me.lbltotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbltotal.Location = New System.Drawing.Point(240, 424)
        Me.lbltotal.Name = "lbltotal"
        Me.lbltotal.Size = New System.Drawing.Size(120, 23)
        Me.lbltotal.TabIndex = 4
        Me.lbltotal.Text = "Label2"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(768, 445)
        Me.Controls.Add(Me.lbltotal)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.DataGrid1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim rows As DataRowCollection = getCustomerNames().Rows
        For Each r As DataRow In rows
            REM Me.ListView1.Items.Add(r.Item("CustomerID"))
            Me.ListBox1.Items.Add(r.Item("CustomerID"))
        Next


    End Sub
    Function GetCustomerTotal(ByVal id As String)
        Dim total As Decimal
        Dim query = "SELECT dbo.getCustomerTotal('" & id & "')"
        Dim cs As String = "Server=.;Database=northwind;Trusted_Connection=True;"
        Dim dt As New DataTable
        Dim sqlcon As New SqlConnection(cs)
        sqlcon.Open()
        Dim sqlcmd As SqlCommand = New SqlCommand(query, sqlcon)
        Dim temp = sqlcmd.ExecuteScalar()
        If temp Is Nothing Then
            total = 0
        Else
            total = DirectCast(temp, Decimal)
        End If

        sqlcon.Close()
        Return total
    End Function
    Function getCustomerNames() As DataTable
        Dim cs As String = "Server=.;Database=northwind;Trusted_Connection=True;"
        Dim dt As New DataTable
        Dim sqlcon As New SqlConnection(cs)
        Dim query As String = "SELECT CustomerID FROM Customers"
        sqlcon.Open()
        Dim ds As DataSet = DbHelper.ExecuteDataset(cs, CommandType.Text, query)
        dt = ds.Tables(0)
        sqlcon.Close()
        Return dt
    End Function

    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click
        Dim lst As ListBox = DirectCast(sender, ListBox)
        Dim id As String = DirectCast(lst.SelectedItem, String)
        GetCustomerOrders(id)
        Me.DataGrid1.DataSource = GetCustomerOrders(id)
        Me.lbltotal.BackColor = Color.Green
        Me.lbltotal.ForeColor = Color.White
        Me.lbltotal.Text = GetCustomerTotal(id)
    End Sub

    Function GetCustomerOrders(ByVal id As String) As DataTable
        Dim cs As String = "Server=.;Database=northwind;Trusted_Connection=True;"
        Dim dt As New DataTable
        Dim sqlcon As New SqlConnection(cs)
        Dim query As New StringBuilder
        query.Append("SELECT c.CustomerID,CAST(o.OrderDate AS DateTime)as OrderDate,od.Quantity ,od.UnitPrice,od.Discount,(od.Quantity*od.UnitPrice)*(1-od.Discount)as Total")
        query.Append(" FROM Customers as c INNER JOIN Orders as o ON c.CustomerID=o.CustomerID INNER JOIN [Order Details] as od ON o.OrderID=od.OrderID")
        query.Append(" WHERE c.CustomerID='" & id & "'")
        sqlcon.Open()
        Dim ds As DataSet = DbHelper.ExecuteDataset(cs, CommandType.Text, query.ToString)
        dt = ds.Tables(0)
        sqlcon.Close()
        Return dt
    End Function
End Class
