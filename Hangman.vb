Public Class frmMain
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
    Friend WithEvents letter1 As System.Windows.Forms.Label
    Friend WithEvents btnPlay As System.Windows.Forms.Button
    Friend WithEvents line1 As System.Windows.Forms.Button
    Friend WithEvents letter2 As System.Windows.Forms.Label
    Friend WithEvents letter3 As System.Windows.Forms.Label
    Friend WithEvents letter4 As System.Windows.Forms.Label
    Friend WithEvents letter5 As System.Windows.Forms.Label
    Friend WithEvents letter6 As System.Windows.Forms.Label
    Friend WithEvents line2 As System.Windows.Forms.Button
    Friend WithEvents line3 As System.Windows.Forms.Button
    Friend WithEvents line4 As System.Windows.Forms.Button
    Friend WithEvents line5 As System.Windows.Forms.Button
    Friend WithEvents line6 As System.Windows.Forms.Button
    Friend WithEvents lblWrongGuesses As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.letter1 = New System.Windows.Forms.Label
        Me.btnPlay = New System.Windows.Forms.Button
        Me.line1 = New System.Windows.Forms.Button
        Me.letter2 = New System.Windows.Forms.Label
        Me.letter3 = New System.Windows.Forms.Label
        Me.letter4 = New System.Windows.Forms.Label
        Me.letter5 = New System.Windows.Forms.Label
        Me.letter6 = New System.Windows.Forms.Label
        Me.line2 = New System.Windows.Forms.Button
        Me.line3 = New System.Windows.Forms.Button
        Me.line4 = New System.Windows.Forms.Button
        Me.line5 = New System.Windows.Forms.Button
        Me.line6 = New System.Windows.Forms.Button
        Me.lblWrongGuesses = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'letter1
        '
        Me.letter1.BackColor = System.Drawing.Color.LimeGreen
        Me.letter1.Font = New System.Drawing.Font("Trebuchet MS", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.letter1.ForeColor = System.Drawing.Color.White
        Me.letter1.Location = New System.Drawing.Point(32, 120)
        Me.letter1.Name = "letter1"
        Me.letter1.Size = New System.Drawing.Size(48, 48)
        Me.letter1.TabIndex = 0
        Me.letter1.Text = "A"
        Me.letter1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnPlay
        '
        Me.btnPlay.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPlay.Font = New System.Drawing.Font("Trebuchet MS", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPlay.Location = New System.Drawing.Point(144, 32)
        Me.btnPlay.Name = "btnPlay"
        Me.btnPlay.Size = New System.Drawing.Size(144, 56)
        Me.btnPlay.TabIndex = 1
        Me.btnPlay.Text = "Play"
        '
        'line1
        '
        Me.line1.Enabled = False
        Me.line1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.line1.Location = New System.Drawing.Point(32, 176)
        Me.line1.Name = "line1"
        Me.line1.Size = New System.Drawing.Size(48, 2)
        Me.line1.TabIndex = 2
        Me.line1.TabStop = False
        '
        'letter2
        '
        Me.letter2.BackColor = System.Drawing.Color.LimeGreen
        Me.letter2.Font = New System.Drawing.Font("Trebuchet MS", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.letter2.ForeColor = System.Drawing.Color.White
        Me.letter2.Location = New System.Drawing.Point(96, 120)
        Me.letter2.Name = "letter2"
        Me.letter2.Size = New System.Drawing.Size(48, 48)
        Me.letter2.TabIndex = 3
        Me.letter2.Text = "A"
        Me.letter2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'letter3
        '
        Me.letter3.BackColor = System.Drawing.Color.LimeGreen
        Me.letter3.Font = New System.Drawing.Font("Trebuchet MS", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.letter3.ForeColor = System.Drawing.Color.White
        Me.letter3.Location = New System.Drawing.Point(160, 120)
        Me.letter3.Name = "letter3"
        Me.letter3.Size = New System.Drawing.Size(48, 48)
        Me.letter3.TabIndex = 4
        Me.letter3.Text = "A"
        Me.letter3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'letter4
        '
        Me.letter4.BackColor = System.Drawing.Color.LimeGreen
        Me.letter4.Font = New System.Drawing.Font("Trebuchet MS", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.letter4.ForeColor = System.Drawing.Color.White
        Me.letter4.Location = New System.Drawing.Point(224, 120)
        Me.letter4.Name = "letter4"
        Me.letter4.Size = New System.Drawing.Size(48, 48)
        Me.letter4.TabIndex = 5
        Me.letter4.Text = "A"
        Me.letter4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'letter5
        '
        Me.letter5.BackColor = System.Drawing.Color.LimeGreen
        Me.letter5.Font = New System.Drawing.Font("Trebuchet MS", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.letter5.ForeColor = System.Drawing.Color.White
        Me.letter5.Location = New System.Drawing.Point(288, 120)
        Me.letter5.Name = "letter5"
        Me.letter5.Size = New System.Drawing.Size(48, 48)
        Me.letter5.TabIndex = 6
        Me.letter5.Text = "A"
        Me.letter5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'letter6
        '
        Me.letter6.BackColor = System.Drawing.Color.LimeGreen
        Me.letter6.Font = New System.Drawing.Font("Trebuchet MS", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.letter6.ForeColor = System.Drawing.Color.White
        Me.letter6.Location = New System.Drawing.Point(352, 120)
        Me.letter6.Name = "letter6"
        Me.letter6.Size = New System.Drawing.Size(48, 48)
        Me.letter6.TabIndex = 7
        Me.letter6.Text = "A"
        Me.letter6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'line2
        '
        Me.line2.Enabled = False
        Me.line2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.line2.Location = New System.Drawing.Point(96, 176)
        Me.line2.Name = "line2"
        Me.line2.Size = New System.Drawing.Size(48, 2)
        Me.line2.TabIndex = 8
        Me.line2.TabStop = False
        '
        'line3
        '
        Me.line3.Enabled = False
        Me.line3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.line3.Location = New System.Drawing.Point(160, 176)
        Me.line3.Name = "line3"
        Me.line3.Size = New System.Drawing.Size(48, 2)
        Me.line3.TabIndex = 9
        Me.line3.TabStop = False
        '
        'line4
        '
        Me.line4.Enabled = False
        Me.line4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.line4.Location = New System.Drawing.Point(224, 176)
        Me.line4.Name = "line4"
        Me.line4.Size = New System.Drawing.Size(48, 2)
        Me.line4.TabIndex = 10
        Me.line4.TabStop = False
        '
        'line5
        '
        Me.line5.Enabled = False
        Me.line5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.line5.Location = New System.Drawing.Point(288, 176)
        Me.line5.Name = "line5"
        Me.line5.Size = New System.Drawing.Size(48, 2)
        Me.line5.TabIndex = 11
        Me.line5.TabStop = False
        '
        'line6
        '
        Me.line6.Enabled = False
        Me.line6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.line6.Location = New System.Drawing.Point(352, 176)
        Me.line6.Name = "line6"
        Me.line6.Size = New System.Drawing.Size(48, 2)
        Me.line6.TabIndex = 12
        Me.line6.TabStop = False
        '
        'lblWrongGuesses
        '
        Me.lblWrongGuesses.Location = New System.Drawing.Point(16, 488)
        Me.lblWrongGuesses.Name = "lblWrongGuesses"
        Me.lblWrongGuesses.Size = New System.Drawing.Size(152, 23)
        Me.lblWrongGuesses.TabIndex = 13
        Me.lblWrongGuesses.Text = "Wrong Guesses:"
        Me.lblWrongGuesses.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Location = New System.Drawing.Point(448, 376)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(120, 26)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = ""
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Trebuchet MS", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Items.AddRange(New Object() {"SEARCH", "OUTPUT", "VISUAL", "SYMBOL", "STUDIO", "MEMBER", "SOURCE", "SORTED", "COLUMN", "DESIGN", "BUTTON", "EXTEND", "DURING", "ROUTES", "DEFINE", "FINISH", "PLEASE", "WINDOW", "MONDAY", "SUNDAY", "ANSWER", "TALENT"})
        Me.ListBox1.Location = New System.Drawing.Point(448, 288)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(120, 84)
        Me.ListBox1.TabIndex = 15
        Me.ListBox1.TabStop = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(434, 522)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.lblWrongGuesses)
        Me.Controls.Add(Me.line6)
        Me.Controls.Add(Me.line5)
        Me.Controls.Add(Me.line4)
        Me.Controls.Add(Me.line3)
        Me.Controls.Add(Me.line2)
        Me.Controls.Add(Me.letter6)
        Me.Controls.Add(Me.letter5)
        Me.Controls.Add(Me.letter4)
        Me.Controls.Add(Me.letter3)
        Me.Controls.Add(Me.letter2)
        Me.Controls.Add(Me.line1)
        Me.Controls.Add(Me.btnPlay)
        Me.Controls.Add(Me.letter1)
        Me.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(440, 550)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(440, 550)
        Me.Name = "frmMain"
        Me.Text = "Hangman"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim word As String                      'holds the word to be guessed
    Dim wordLength As Integer               'holds the length of the word
    Dim char1 As String                     'holds the first letter of the word
    Dim char2 As String                     'holds the second letter of the word
    Dim char3 As String                     'holds the third letter of the word
    Dim char4 As String                     'holds the fourth letter of the word
    Dim char5 As String                     'holds the fifth letter of the word
    Dim char6 As String                     'holds the sixth letter of the word
    Dim gameStarted As Boolean = False      'switches to true when the user clicks 'Play'
    Dim rightTries As Integer               'number of correct letter guesses
    Dim wrongTries As Integer               'number of incorrect letter guesses

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        clearLetters()
        hideCharacterLength()

    End Sub

    Public Sub drawHanger()

        Dim pnt As Graphics = Me.CreateGraphics
        Dim pencil As New Pen(Color.Black, 2)
        pnt.DrawLine(pencil, 46, 434, 115, 394)
        pnt.DrawLine(pencil, 115, 394, 184, 434)
        pnt.DrawLine(pencil, 115, 395, 115, 220)
        pnt.DrawLine(pencil, 115, 220, 250, 220)
        pnt.DrawLine(pencil, 250, 220, 250, 250)

    End Sub

    Public Sub resetGame()

        lblWrongGuesses.Text = ""
        Dim pnt As Graphics = Me.CreateGraphics
        'pnt.Clear(Color.FromArgb(0, 192, 0))
        btnPlay.Visible = True

    End Sub

    Public Sub clearLetters()

        letter1.Text = ""
        letter2.Text = ""
        letter3.Text = ""
        letter4.Text = ""
        letter5.Text = ""
        letter6.Text = ""

    End Sub

    Public Sub checkForWinner()

        If rightTries = wordLength Then
            MsgBox("You Win! :D")
            resetGame()
        End If

    End Sub

    Public Sub startDrawingWhenLosing(ByVal tries As Integer)

        Dim pnt As Graphics = Me.CreateGraphics
        Dim pencil As New Pen(Color.Black, 2)
        If tries = 1 Then
            pnt.DrawEllipse(pencil, 228, 250, 40, 40)
        ElseIf tries = 2 Then
            pnt.DrawLine(pencil, 248, 290, 248, 370)
        ElseIf tries = 3 Then
            pnt.DrawLine(pencil, 248, 315, 213, 293)
        ElseIf tries = 4 Then
            pnt.DrawLine(pencil, 248, 315, 283, 293)
        ElseIf tries = 5 Then
            pnt.DrawLine(pencil, 248, 369, 213, 391)
        ElseIf tries = 6 Then
            pnt.DrawLine(pencil, 248, 369, 283, 391)
            TextBox1.Clear()
            Try
                letter1.Text = char1
                letter2.Text = char2
                letter3.Text = char3
                letter4.Text = char4
                letter5.Text = char5
                letter6.Text = char6
            Catch ex As Exception
            End Try
            MsgBox("You Lose! :(")
            resetGame()
        End If

    End Sub

    Public Sub assignLetters()

        If wordLength = 6 Then

            char1 = word.Chars(0).ToString.ToUpper
            char2 = word.Chars(1).ToString.ToUpper
            char3 = word.Chars(2).ToString.ToUpper
            char4 = word.Chars(3).ToString.ToUpper
            char5 = word.Chars(4).ToString.ToUpper
            char6 = word.Chars(5).ToString.ToUpper

        End If

    End Sub

    Public Sub hideCharacterLength()

        line1.Visible = True
        line2.Visible = True
        line3.Visible = True
        line4.Visible = True
        line5.Visible = True
        line6.Visible = True

    End Sub

    Public Sub showLength(ByVal VisibleLines As Integer)

        hideCharacterLength()
        If VisibleLines = 6 Then

            line1.Visible = True
            line2.Visible = True
            line3.Visible = True
            line4.Visible = True
            line5.Visible = True
            line6.Visible = True

        End If

    End Sub

    Private Sub btnPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlay.Click

        char1 = ""
        char2 = ""
        char3 = ""
        char4 = ""
        char5 = ""
        char6 = ""
        drawHanger()
        Dim number As Integer
        Randomize()
        number = Int(Rnd() * ListBox1.Items.Count - 1) + 1
        word = ListBox1.Items(number)
        wordLength = word.Length
        showLength(word.Length)
        clearLetters()
        assignLetters()
        TextBox1.Focus()
        rightTries = 0
        wrongTries = 0
        lblWrongGuesses.Text = "Wrong Guesses: "
        gameStarted = True
        btnPlay.Visible = False

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

        If gameStarted = True Then
            If TextBox1.Text = "" Then
            Else
                If TextBox1.Text.ToUpper = char1 And letter1.Text = "" And char1 <> "" Then
                    letter1.Text = char1
                    rightTries += 1
                    checkForWinner()
                ElseIf TextBox1.Text.ToUpper = char2 And letter2.Text = "" And char2 <> "" Then
                    letter2.Text = char2
                    rightTries += 1
                    checkForWinner()
                ElseIf TextBox1.Text.ToUpper = char3 And letter3.Text = "" And char3 <> "" Then
                    letter3.Text = char3
                    rightTries += 1
                    checkForWinner()
                ElseIf TextBox1.Text.ToUpper = char4 And letter4.Text = "" And char4 <> "" Then
                    letter4.Text = char4
                    rightTries += 1
                    checkForWinner()
                ElseIf TextBox1.Text.ToUpper = char5 And letter5.Text = "" And char5 <> "" Then
                    letter5.Text = char5
                    rightTries += 1
                    checkForWinner()
                ElseIf TextBox1.Text.ToUpper = char6 And letter6.Text = "" And char6 <> "" Then
                    letter6.Text = char6
                    rightTries += 1
                    checkForWinner()
                Else
                    wrongTries += 1
                    startDrawingWhenLosing(wrongTries)
                    lblWrongGuesses.Text = lblWrongGuesses.Text & " " & TextBox1.Text
                End If
            End If
        Else
        End If
        TextBox1.Text = ""

    End Sub
End Class
