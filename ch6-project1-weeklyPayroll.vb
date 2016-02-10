Public Class frmMain
    Inherits System.Windows.Forms.Form

'Windows Generated Code
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
    Friend WithEvents btnWeeklyDataForm As System.Windows.Forms.Button
    Friend WithEvents btnTotalWeekPayroll As System.Windows.Forms.Button
    Friend WithEvents btnClearForm As System.Windows.Forms.Button
    Friend WithEvents clmEmployeeName As System.Windows.Forms.ColumnHeader
    Friend WithEvents clmHours As System.Windows.Forms.ColumnHeader
    Friend WithEvents clmRate As System.Windows.Forms.ColumnHeader
    Friend WithEvents lstEmployee As System.Windows.Forms.ListView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.btnWeeklyDataForm = New System.Windows.Forms.Button
        Me.btnTotalWeekPayroll = New System.Windows.Forms.Button
        Me.btnClearForm = New System.Windows.Forms.Button
        Me.lstEmployee = New System.Windows.Forms.ListView
        Me.clmEmployeeName = New System.Windows.Forms.ColumnHeader
        Me.clmHours = New System.Windows.Forms.ColumnHeader
        Me.clmRate = New System.Windows.Forms.ColumnHeader
        Me.SuspendLayout()
        '
        'btnWeeklyDataForm
        '
        Me.btnWeeklyDataForm.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWeeklyDataForm.Font = New System.Drawing.Font("Trebuchet MS", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWeeklyDataForm.Location = New System.Drawing.Point(344, 8)
        Me.btnWeeklyDataForm.Name = "btnWeeklyDataForm"
        Me.btnWeeklyDataForm.Size = New System.Drawing.Size(144, 96)
        Me.btnWeeklyDataForm.TabIndex = 0
        Me.btnWeeklyDataForm.Text = "Enter Weekly Data"
        '
        'btnTotalWeekPayroll
        '
        Me.btnTotalWeekPayroll.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnTotalWeekPayroll.Font = New System.Drawing.Font("Trebuchet MS", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTotalWeekPayroll.Location = New System.Drawing.Point(344, 112)
        Me.btnTotalWeekPayroll.Name = "btnTotalWeekPayroll"
        Me.btnTotalWeekPayroll.Size = New System.Drawing.Size(144, 96)
        Me.btnTotalWeekPayroll.TabIndex = 1
        Me.btnTotalWeekPayroll.Text = "Total Weekly Payroll"
        '
        'btnClearForm
        '
        Me.btnClearForm.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClearForm.Font = New System.Drawing.Font("Trebuchet MS", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClearForm.Location = New System.Drawing.Point(344, 216)
        Me.btnClearForm.Name = "btnClearForm"
        Me.btnClearForm.Size = New System.Drawing.Size(144, 96)
        Me.btnClearForm.TabIndex = 2
        Me.btnClearForm.Text = "Clear Weekly Data"
        '
        'lstEmployee
        '
        Me.lstEmployee.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstEmployee.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.clmEmployeeName, Me.clmHours, Me.clmRate})
        Me.lstEmployee.GridLines = True
        Me.lstEmployee.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lstEmployee.Location = New System.Drawing.Point(0, 0)
        Me.lstEmployee.Name = "lstEmployee"
        Me.lstEmployee.Size = New System.Drawing.Size(336, 328)
        Me.lstEmployee.TabIndex = 3
        Me.lstEmployee.View = System.Windows.Forms.View.Details
        '
        'clmEmployeeName
        '
        Me.clmEmployeeName.Text = "Employee Name"
        Me.clmEmployeeName.Width = 197
        '
        'clmHours
        '
        Me.clmHours.Text = "Hours"
        '
        'clmRate
        '
        Me.clmRate.Text = "Rate"
        Me.clmRate.Width = 78
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(496, 322)
        Me.Controls.Add(Me.lstEmployee)
        Me.Controls.Add(Me.btnClearForm)
        Me.Controls.Add(Me.btnTotalWeekPayroll)
        Me.Controls.Add(Me.btnWeeklyDataForm)
        Me.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(512, 360)
        Me.MinimumSize = New System.Drawing.Size(512, 360)
        Me.Name = "frmMain"
        Me.Text = "Payroll"
        Me.ResumeLayout(False)

    End Sub

#End Region
'End of Windows Generated Code

    'Chapter 6: Weekly Payroll
    'Programmer: Sebastian Kaye
    'Date: February 10, 2016
    'Purpose: The purpose of this program is to calculate the payroll from a list view box

    'Clears the entire list
    Private Sub btnClearForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearForm.Click

        lstEmployee.Items.Clear()

    End Sub

    Private Sub btnWeeklyDataForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWeeklyDataForm.Click

        'Creates a new form
        Dim weeklyInput As New frmInputWeeklyData

        'The location to jump back if there is an error
redrawForm2:
        'Shows the new form
        weeklyInput.ShowDialog()

        'Creates a new dataset for employee data
        'Adds the name to the dataset
        Dim employeeData As New ListViewItem(weeklyInput.txtEmployeeName.Text)

        'Adds the number of hours worked and the pay rate to the set
        employeeData.SubItems.Add(weeklyInput.numHours.Value)
        employeeData.SubItems.Add(Format$(weeklyInput.numRate.Value, "Currency"))

        'Checks to see that there is no null data
        If Not weeklyInput.txtEmployeeName.Text = "" And Not weeklyInput.numRate.Value = 0 Then

            lstEmployee.Items.Add(employeeData)

        End If

        'an error message is displayed if there is no pay rate and jumps back to the form load
        If weeklyInput.txtEmployeeName.Text = "" Or weeklyInput.numRate.Value = 0 Then

            errorMessage()
            GoTo redrawForm2

        End If

    End Sub

    'Totals all of the payroll
    Private Sub btnTotalWeekPayroll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTotalWeekPayroll.Click

        Dim counter, intHours As Integer
        Dim strRate As String = ""
        Dim strResults As String = ""
        Dim sngRate As Single = 0
        Dim sngPayChecks As Single = 0
        Dim sngListCount As Single = 0

        'The number of items on the list is assigned to intListCount
        sngListCount = Convert.ToSingle(lstEmployee.Items.Count)

        'This for loop will iterate through each item in the list
        For counter = 0 To sngListCount - 1

            'Converts the hours worked to the employee at location counter to an integer
            intHours = Convert.ToInt32(lstEmployee.Items(counter).SubItems(1).Text)
            'Converts the pay rate to a string
            strRate = lstEmployee.Items(counter).SubItems(2).Text
            'Replaces the $ with 0 and converts the pay rate to type Single
            sngRate = strRate.Replace("$", "0")
            'Calculates the actual paycheck amount for the employee at the location counter
            sngPayChecks += (intHours * sngRate)

        Next

        'Outputs the results to a message box
        strResults = "Total payroll: " & Format$(sngPayChecks, "Currency") & ControlChars.NewLine
        strResults &= "Total number of employees: " & sngListCount
        MessageBox.Show(strResults, "Total Paychecks")

    End Sub

    'Displays an error message
    Sub errorMessage()

        MsgBox("Uh Oh! Something went wrong. Please re-enter.", MsgBoxStyle.OKOnly, "Oh No!")

    End Sub

End Class
