Public Class Form1
    Inherits System.Windows.Forms.Form

'WARNING WARNING WARNING Below is the Windows generated code. Don't edit it unless you want everything to go wrong.
'If you copy this then the actual form itself should copy as well.
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents badLetterCheck As System.Windows.Forms.ListBox
    Friend WithEvents input As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label2 = New System.Windows.Forms.Label
        Me.badLetterCheck = New System.Windows.Forms.ListBox
        Me.input = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(0, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(184, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Bad Character Checker"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'badLetterCheck
        '
        Me.badLetterCheck.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.badLetterCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.badLetterCheck.ItemHeight = 20
        Me.badLetterCheck.Location = New System.Drawing.Point(8, 32)
        Me.badLetterCheck.Name = "badLetterCheck"
        Me.badLetterCheck.Size = New System.Drawing.Size(176, 162)
        Me.badLetterCheck.TabIndex = 3
        '
        'input
        '
        Me.input.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.input.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.input.Location = New System.Drawing.Point(8, 208)
        Me.input.Name = "input"
        Me.input.Size = New System.Drawing.Size(176, 29)
        Me.input.TabIndex = 4
        Me.input.Text = ""
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(8, 240)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(176, 32)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Check"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(192, 282)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.input)
        Me.Controls.Add(Me.badLetterCheck)
        Me.Controls.Add(Me.Label2)
        Me.Name = "Form1"
        Me.Text = "Bad Character Checker"
        Me.ResumeLayout(False)

    End Sub

#End Region
'WARNING WARNING WARNING This is the end of the Windows generated code.

    Private Sub init()

        badLetterCheck.Items.Clear()

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        init()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim a, b, c As Integer
        'the unallowed characters
        Dim badLetters As String = "abcdefghijklmnopqrstuvwxyz ;'""!@#$%^&*()_-+=,<>/?\|]}[{`~"
        'the allowed characters
        Dim goodLetters As String = "1234567890."
        'the arrays
        Dim stringToCheck() As Char
        Dim badLettersArray() As Char
        Dim goodLettersArray() As Char

        init()

        'creates the arrays and forces the input to a lowercase form
        stringToCheck = input.Text.ToLower
        badLettersArray = badLetters
        goodLettersArray = goodLetters

        '=====the "wrapper" for loop=====
        'the for loop gets the text's length and examines each character
        For a = 0 To input.TextLength - 1

            'these next 2 for loops check to see if the character is "good" or "bad"
            For b = 0 To badLetters.Length - 1

                'this if statement checks to see if the character at location (a) in the counter
                '   is equal to location (b) in the bad character counter
                If stringToCheck(a) = badLettersArray(b) Then

                    badLetterCheck.Items.Add(stringToCheck(a) & "  -  True")

                End If

            Next
            'checks for good characters
            For c = 0 To goodLetters.Length - 1

                'this if statement checks to see if the character at location (a) in the counter
                '   is equal to location (c) in the good character counter
                If stringToCheck(a) = goodLettersArray(c) Then

                    badLetterCheck.Items.Add(stringToCheck(a) & "  -  False")

                End If

            Next

        Next

    End Sub

End Class
