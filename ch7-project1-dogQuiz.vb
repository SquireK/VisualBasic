Option Strict On
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
    Friend WithEvents grpAnswer As System.Windows.Forms.GroupBox
    Friend WithEvents radAnswer1 As System.Windows.Forms.RadioButton
    Friend WithEvents radAnswer2 As System.Windows.Forms.RadioButton
    Friend WithEvents radAnswer3 As System.Windows.Forms.RadioButton
    Friend WithEvents radAnswer4 As System.Windows.Forms.RadioButton
    Friend WithEvents btnPrevious As System.Windows.Forms.Button
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuQuestions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFirst As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLast As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSeperator As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNext As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrevious As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOptions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFont As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents cmuShortcut As System.Windows.Forms.ContextMenu
    Friend WithEvents cmuNext As System.Windows.Forms.MenuItem
    Friend WithEvents cmuPrevious As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents cmuFont As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents cmuFinish As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents cmuExit As System.Windows.Forms.MenuItem
    Friend WithEvents stbQuestion As System.Windows.Forms.StatusBar
    Friend WithEvents picQuestion As System.Windows.Forms.PictureBox
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents fodQuestion As System.Windows.Forms.FontDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.mnuMain = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuQuestions = New System.Windows.Forms.MenuItem
        Me.mnuFirst = New System.Windows.Forms.MenuItem
        Me.mnuLast = New System.Windows.Forms.MenuItem
        Me.mnuSeperator = New System.Windows.Forms.MenuItem
        Me.mnuNext = New System.Windows.Forms.MenuItem
        Me.mnuPrevious = New System.Windows.Forms.MenuItem
        Me.mnuOptions = New System.Windows.Forms.MenuItem
        Me.mnuFont = New System.Windows.Forms.MenuItem
        Me.stbQuestion = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.picQuestion = New System.Windows.Forms.PictureBox
        Me.grpAnswer = New System.Windows.Forms.GroupBox
        Me.radAnswer4 = New System.Windows.Forms.RadioButton
        Me.radAnswer3 = New System.Windows.Forms.RadioButton
        Me.radAnswer2 = New System.Windows.Forms.RadioButton
        Me.radAnswer1 = New System.Windows.Forms.RadioButton
        Me.btnPrevious = New System.Windows.Forms.Button
        Me.btnNext = New System.Windows.Forms.Button
        Me.fodQuestion = New System.Windows.Forms.FontDialog
        Me.cmuShortcut = New System.Windows.Forms.ContextMenu
        Me.cmuNext = New System.Windows.Forms.MenuItem
        Me.cmuPrevious = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.cmuFont = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.cmuFinish = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.cmuExit = New System.Windows.Forms.MenuItem
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAnswer.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuQuestions, Me.mnuOptions})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 0
        Me.mnuExit.Text = "E&xit"
        '
        'mnuQuestions
        '
        Me.mnuQuestions.Index = 1
        Me.mnuQuestions.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFirst, Me.mnuLast, Me.mnuSeperator, Me.mnuNext, Me.mnuPrevious})
        Me.mnuQuestions.Text = "&Questions"
        '
        'mnuFirst
        '
        Me.mnuFirst.Index = 0
        Me.mnuFirst.Shortcut = System.Windows.Forms.Shortcut.CtrlF
        Me.mnuFirst.Text = "&First"
        '
        'mnuLast
        '
        Me.mnuLast.Index = 1
        Me.mnuLast.Shortcut = System.Windows.Forms.Shortcut.CtrlL
        Me.mnuLast.Text = "&Last"
        '
        'mnuSeperator
        '
        Me.mnuSeperator.Index = 2
        Me.mnuSeperator.Text = "-"
        '
        'mnuNext
        '
        Me.mnuNext.Index = 3
        Me.mnuNext.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuNext.Text = "&Next"
        '
        'mnuPrevious
        '
        Me.mnuPrevious.Index = 4
        Me.mnuPrevious.Shortcut = System.Windows.Forms.Shortcut.CtrlP
        Me.mnuPrevious.Text = "&Previous"
        '
        'mnuOptions
        '
        Me.mnuOptions.Index = 2
        Me.mnuOptions.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFont})
        Me.mnuOptions.Text = "&Options"
        '
        'mnuFont
        '
        Me.mnuFont.Index = 0
        Me.mnuFont.Text = "&Font"
        '
        'stbQuestion
        '
        Me.stbQuestion.Font = New System.Drawing.Font("Trebuchet MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.stbQuestion.Location = New System.Drawing.Point(0, 353)
        Me.stbQuestion.Name = "stbQuestion"
        Me.stbQuestion.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1})
        Me.stbQuestion.ShowPanels = True
        Me.stbQuestion.Size = New System.Drawing.Size(582, 22)
        Me.stbQuestion.TabIndex = 0
        Me.stbQuestion.Text = "StatusBar1"
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.StatusBarPanel1.Text = "StatusBarPanel1"
        Me.StatusBarPanel1.Width = 566
        '
        'picQuestion
        '
        Me.picQuestion.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picQuestion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picQuestion.Location = New System.Drawing.Point(8, 8)
        Me.picQuestion.Name = "picQuestion"
        Me.picQuestion.Size = New System.Drawing.Size(320, 304)
        Me.picQuestion.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picQuestion.TabIndex = 1
        Me.picQuestion.TabStop = False
        '
        'grpAnswer
        '
        Me.grpAnswer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpAnswer.Controls.Add(Me.radAnswer4)
        Me.grpAnswer.Controls.Add(Me.radAnswer3)
        Me.grpAnswer.Controls.Add(Me.radAnswer2)
        Me.grpAnswer.Controls.Add(Me.radAnswer1)
        Me.grpAnswer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.grpAnswer.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpAnswer.Location = New System.Drawing.Point(352, 8)
        Me.grpAnswer.Name = "grpAnswer"
        Me.grpAnswer.Size = New System.Drawing.Size(216, 152)
        Me.grpAnswer.TabIndex = 0
        Me.grpAnswer.TabStop = False
        Me.grpAnswer.Text = "Choose Answer"
        '
        'radAnswer4
        '
        Me.radAnswer4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radAnswer4.Font = New System.Drawing.Font("Trebuchet MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radAnswer4.Location = New System.Drawing.Point(16, 120)
        Me.radAnswer4.Name = "radAnswer4"
        Me.radAnswer4.Size = New System.Drawing.Size(192, 24)
        Me.radAnswer4.TabIndex = 3
        Me.radAnswer4.Text = "RadioButton4"
        '
        'radAnswer3
        '
        Me.radAnswer3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radAnswer3.Font = New System.Drawing.Font("Trebuchet MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radAnswer3.Location = New System.Drawing.Point(16, 88)
        Me.radAnswer3.Name = "radAnswer3"
        Me.radAnswer3.Size = New System.Drawing.Size(192, 24)
        Me.radAnswer3.TabIndex = 2
        Me.radAnswer3.Text = "RadioButton3"
        '
        'radAnswer2
        '
        Me.radAnswer2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radAnswer2.Font = New System.Drawing.Font("Trebuchet MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radAnswer2.Location = New System.Drawing.Point(16, 56)
        Me.radAnswer2.Name = "radAnswer2"
        Me.radAnswer2.Size = New System.Drawing.Size(192, 24)
        Me.radAnswer2.TabIndex = 1
        Me.radAnswer2.Text = "RadioButton2"
        '
        'radAnswer1
        '
        Me.radAnswer1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radAnswer1.Font = New System.Drawing.Font("Trebuchet MS", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radAnswer1.Location = New System.Drawing.Point(16, 24)
        Me.radAnswer1.Name = "radAnswer1"
        Me.radAnswer1.Size = New System.Drawing.Size(192, 24)
        Me.radAnswer1.TabIndex = 0
        Me.radAnswer1.Text = "RadioButton1"
        '
        'btnPrevious
        '
        Me.btnPrevious.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrevious.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPrevious.Font = New System.Drawing.Font("Trebuchet MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrevious.Location = New System.Drawing.Point(376, 176)
        Me.btnPrevious.Name = "btnPrevious"
        Me.btnPrevious.Size = New System.Drawing.Size(72, 32)
        Me.btnPrevious.TabIndex = 1
        Me.btnPrevious.Text = "Previous"
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNext.Font = New System.Drawing.Font("Trebuchet MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext.Location = New System.Drawing.Point(464, 176)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(72, 32)
        Me.btnNext.TabIndex = 2
        Me.btnNext.Text = "Next"
        '
        'fodQuestion
        '
        Me.fodQuestion.MaxSize = 14
        Me.fodQuestion.MinSize = 8
        Me.fodQuestion.ScriptsOnly = True
        Me.fodQuestion.ShowEffects = False
        '
        'cmuShortcut
        '
        Me.cmuShortcut.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmuNext, Me.cmuPrevious, Me.MenuItem1, Me.cmuFont, Me.MenuItem2, Me.cmuFinish, Me.MenuItem3, Me.cmuExit})
        '
        'cmuNext
        '
        Me.cmuNext.Index = 0
        Me.cmuNext.Text = "Next"
        '
        'cmuPrevious
        '
        Me.cmuPrevious.Index = 1
        Me.cmuPrevious.Text = "Previous"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.Text = "-"
        '
        'cmuFont
        '
        Me.cmuFont.Index = 3
        Me.cmuFont.Text = "Font"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 4
        Me.MenuItem2.Text = "-"
        '
        'cmuFinish
        '
        Me.cmuFinish.Index = 5
        Me.cmuFinish.Text = "Finish Quiz and Display Result"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 6
        Me.MenuItem3.Text = "-"
        '
        'cmuExit
        '
        Me.cmuExit.Index = 7
        Me.cmuExit.Text = "Exit"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(582, 375)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrevious)
        Me.Controls.Add(Me.grpAnswer)
        Me.Controls.Add(Me.picQuestion)
        Me.Controls.Add(Me.stbQuestion)
        Me.Font = New System.Drawing.Font("Trebuchet MS", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Menu = Me.mnuMain
        Me.MinimumSize = New System.Drawing.Size(598, 413)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Know Your Dogs Quiz"
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAnswer.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Chapter 7:         Know Your Dogs Quiz
    'Programmer:        Sebastian Kaye
    'Date:              February 25, 2016
    'Description:       This program generates a seven question quiz. Each question
    '                   includes an image and four choices in RadioButton controls.
    '                   The user must answer all questions correctly to pass the quiz.
    '                   The solution is stored in the gshrSolution array. The user answers
    '                   are stored in the gshrAnswers array. Questions choices and image
    '                   file names are stored in the gstrQuestions array.

    Dim gshrSolution As Short() = {3, 0, 1, 2, 3, 3, 0}
    Dim gshrAnswers As Short() = {-1, -1, -1, -1, -1, -1, -1}
    Dim gshrCurrentQuestion As Short = 0
    Dim gstrQuestions As String(,) = { _
    {"English Springer Spaniel", "Akita", "Husky", "Bearded Collie", "dog051.jpg"}, _
    {"Pug", "Whippet", "Fox Terrier", "English Setter", "dog50.jpg"}, _
    {"Chow Chow", "Beagle", "Labrador Retriever", "Foxhound", "dog068.jpg"}, _
    {"Great Dane", "Dachshund", "Samoyed", "Newfoundland", "dog044.jpg"}, _
    {"Samoyed", "Pekingese", "English Setter", "Pomeranian", "dog088.jpg"}, _
    {"Old English Sheepdog", "Saluki", "Rottweiler", "poodle", "dog074.jpg"}, _
    {"Collie", "Boston Terrier", "Bichon Frise", "Akita", "dog007.jpg"}}

    'This function checks a user's input with the correct answer
    Private Function QuizResult() As Boolean

        Dim shrIndex As Short

        'This for-if statement check to see if the answer is incorrect and returns to
        'the program if the answer is false.
        For shrIndex = 0 To 6
            If gshrSolution(shrIndex) <> gshrAnswers(shrIndex) Then
                Return False
            End If
        Next

        'If the answer is correct, then the function returns with a true value.
        QuizResult = True

    End Function

    'This has more than a comment can explain
    Private Sub DisplayQuestion(ByVal shrQuestion As Short)

        Dim shrIndex As Short

        'Check if user clicked 'Previous' button with question 1 displayed
        If shrQuestion < 0 Then
            MessageBox.Show("You are already at the beginning of the quiz", "Know Your Dogs Quiz")
            gshrCurrentQuestion = 0
            Exit Sub
        End If

        'Store user's answer to previous question in gshrAnswers array
        'in the case that the first radio button has been selected the number representing
        'the selected radiobutton is assigned to gshrCurrentQuestion and unchecked
        Select Case True
            Case radAnswer1.Checked
                gshrAnswers(gshrCurrentQuestion) = 0
                radAnswer1.Checked = False
            Case radAnswer2.Checked
                gshrAnswers(gshrCurrentQuestion) = 1
                radAnswer2.Checked = False
            Case radAnswer3.Checked
                gshrAnswers(gshrCurrentQuestion) = 2
                radAnswer3.Checked = False
            Case radAnswer4.Checked
                gshrAnswers(gshrCurrentQuestion) = 3
                radAnswer4.Checked = False
        End Select

        'Scoring of the quiz and the closing of the program
        If shrQuestion > gstrQuestions.GetUpperBound(0) Then
            'Askes the user if they want to see their score
            'If not, they are taken back to the quiz.
            If MessageBox.Show("You have answered the last question. Would you like to display the final quiz result and exit the quiz?" _
            & ControlChars.NewLine & ControlChars.NewLine & "Click No to return.", "End of quiz", _
            MessageBoxButtons.YesNo, MessageBoxIcon.Question, _
            MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                'If they pass
                If QuizResult() Then
                    MessageBox.Show("Congratulations! You passed!", "Know Your Dogs Quiz")
                Else 'Or if they don't
                    MessageBox.Show("Sorry, you did not pass the quiz.", "Know Your Dogs Quiz")
                End If
                Me.Close()
            Else
                gshrCurrentQuestion = Convert.ToInt16(gstrQuestions.GetUpperBound(0))
            End If
            Exit Sub
        End If

        'Update controls for next question
        For shrIndex = 0 To 3
            grpAnswer.Controls(shrIndex).Text = gstrQuestions(shrQuestion, shrIndex)
        Next
        'Displays the image
        picQuestion.Image = Image.FromFile(gstrQuestions(shrQuestion, 4))
        stbQuestion.Panels(0).Text = "Question " & shrQuestion + 1 & " of " & _
        gstrQuestions.GetUpperBound(0) + 1

        Select Case gshrAnswers(shrQuestion)
            Case 0
                radAnswer1.Checked = True
            Case 1
                radAnswer2.Checked = True
            Case 2
                radAnswer3.Checked = True
            Case 3
                radAnswer4.Checked = True
        End Select

        gshrCurrentQuestion = shrQuestion

    End Sub

#Region " Menus and Buttons "
    'Closes the application
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click

        Me.Close()

    End Sub
    Private Sub cmuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmuExit.Click
        mnuExit.PerformClick()
    End Sub
    'Sets the RadioButton fonts
    Private Sub mnuFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFont.Click

        'Tells the font menu which font the radiobuttons are using
        fodQuestion.Font = radAnswer1.Font

        'Assigns the new font to the buttons
        If fodQuestion.ShowDialog() = DialogResult.OK Then
            radAnswer1.Font = fodQuestion.Font
            radAnswer2.Font = fodQuestion.Font
            radAnswer3.Font = fodQuestion.Font
            radAnswer4.Font = fodQuestion.Font
        End If

    End Sub
    Private Sub cmuFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmuFont.Click
        mnuFont.PerformClick()
    End Sub
    'Displays the first question
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DisplayQuestion(0)
    End Sub
    Private Sub mnuFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFirst.Click
        DisplayQuestion(0)
    End Sub
    'Displays the last question
    Private Sub mnuLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLast.Click
        DisplayQuestion(Convert.ToInt16(gstrQuestions.GetUpperBound(0)))
    End Sub
    'Displays the next question
    Private Sub mnuNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNext.Click
        DisplayQuestion(gshrCurrentQuestion + Convert.ToInt16(1))
    End Sub
    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        mnuNext.PerformClick()
    End Sub
    Private Sub cmuNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmuNext.Click
        mnuNext.PerformClick()
    End Sub
    'Displays the previous Question
    Private Sub mnuPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrevious.Click
        DisplayQuestion(gshrCurrentQuestion - Convert.ToInt16(1))
    End Sub
    Private Sub btnPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious.Click
        mnuPrevious.PerformClick()
    End Sub
    Private Sub cmuPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmuPrevious.Click
        mnuPrevious.PerformClick()
    End Sub
    'Navigates past the last question
    Private Sub cmuFinish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmuFinish.Click
        DisplayQuestion(Convert.ToInt16(gstrQuestions.GetUpperBound(0) + 1))
    End Sub
#End Region

End Class
