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
    Friend WithEvents btnMidRight As System.Windows.Forms.Button
    Friend WithEvents btnTopLeft As System.Windows.Forms.Button
    Friend WithEvents btnTopRight As System.Windows.Forms.Button
    Friend WithEvents btnTopCenter As System.Windows.Forms.Button
    Friend WithEvents btnMidCenter As System.Windows.Forms.Button
    Friend WithEvents btnMidLeft As System.Windows.Forms.Button
    Friend WithEvents btnLowLeft As System.Windows.Forms.Button
    Friend WithEvents btnLowCenter As System.Windows.Forms.Button
    Friend WithEvents btnLowRight As System.Windows.Forms.Button
    Friend WithEvents btnSinglePlayer As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents btnMultiPlayer As System.Windows.Forms.Button
    Friend WithEvents lblPlayer1Score As System.Windows.Forms.Label
    Friend WithEvents lblPlayer2Score As System.Windows.Forms.Label
    Friend WithEvents lblTurn As System.Windows.Forms.Label
    Friend WithEvents btnClearBoard As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.btnTopLeft = New System.Windows.Forms.Button
        Me.btnTopCenter = New System.Windows.Forms.Button
        Me.btnTopRight = New System.Windows.Forms.Button
        Me.btnMidLeft = New System.Windows.Forms.Button
        Me.btnMidCenter = New System.Windows.Forms.Button
        Me.btnMidRight = New System.Windows.Forms.Button
        Me.btnLowLeft = New System.Windows.Forms.Button
        Me.btnLowCenter = New System.Windows.Forms.Button
        Me.btnLowRight = New System.Windows.Forms.Button
        Me.btnSinglePlayer = New System.Windows.Forms.Button
        Me.btnReset = New System.Windows.Forms.Button
        Me.lblPlayer1Score = New System.Windows.Forms.Label
        Me.lblPlayer2Score = New System.Windows.Forms.Label
        Me.btnMultiPlayer = New System.Windows.Forms.Button
        Me.lblTurn = New System.Windows.Forms.Label
        Me.btnClearBoard = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnTopLeft
        '
        Me.btnTopLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnTopLeft.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTopLeft.Location = New System.Drawing.Point(16, 16)
        Me.btnTopLeft.Name = "btnTopLeft"
        Me.btnTopLeft.Size = New System.Drawing.Size(96, 96)
        Me.btnTopLeft.TabIndex = 0
        Me.btnTopLeft.TabStop = False
        '
        'btnTopCenter
        '
        Me.btnTopCenter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnTopCenter.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTopCenter.Location = New System.Drawing.Point(112, 16)
        Me.btnTopCenter.Name = "btnTopCenter"
        Me.btnTopCenter.Size = New System.Drawing.Size(96, 96)
        Me.btnTopCenter.TabIndex = 1
        Me.btnTopCenter.TabStop = False
        '
        'btnTopRight
        '
        Me.btnTopRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnTopRight.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTopRight.Location = New System.Drawing.Point(208, 16)
        Me.btnTopRight.Name = "btnTopRight"
        Me.btnTopRight.Size = New System.Drawing.Size(96, 96)
        Me.btnTopRight.TabIndex = 2
        Me.btnTopRight.TabStop = False
        '
        'btnMidLeft
        '
        Me.btnMidLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMidLeft.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMidLeft.Location = New System.Drawing.Point(16, 112)
        Me.btnMidLeft.Name = "btnMidLeft"
        Me.btnMidLeft.Size = New System.Drawing.Size(96, 96)
        Me.btnMidLeft.TabIndex = 3
        Me.btnMidLeft.TabStop = False
        '
        'btnMidCenter
        '
        Me.btnMidCenter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMidCenter.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMidCenter.Location = New System.Drawing.Point(112, 112)
        Me.btnMidCenter.Name = "btnMidCenter"
        Me.btnMidCenter.Size = New System.Drawing.Size(96, 96)
        Me.btnMidCenter.TabIndex = 4
        Me.btnMidCenter.TabStop = False
        '
        'btnMidRight
        '
        Me.btnMidRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMidRight.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMidRight.Location = New System.Drawing.Point(208, 112)
        Me.btnMidRight.Name = "btnMidRight"
        Me.btnMidRight.Size = New System.Drawing.Size(96, 96)
        Me.btnMidRight.TabIndex = 5
        Me.btnMidRight.TabStop = False
        '
        'btnLowLeft
        '
        Me.btnLowLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLowLeft.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLowLeft.Location = New System.Drawing.Point(16, 208)
        Me.btnLowLeft.Name = "btnLowLeft"
        Me.btnLowLeft.Size = New System.Drawing.Size(96, 96)
        Me.btnLowLeft.TabIndex = 6
        Me.btnLowLeft.TabStop = False
        '
        'btnLowCenter
        '
        Me.btnLowCenter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLowCenter.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLowCenter.Location = New System.Drawing.Point(112, 208)
        Me.btnLowCenter.Name = "btnLowCenter"
        Me.btnLowCenter.Size = New System.Drawing.Size(96, 96)
        Me.btnLowCenter.TabIndex = 7
        Me.btnLowCenter.TabStop = False
        '
        'btnLowRight
        '
        Me.btnLowRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLowRight.Font = New System.Drawing.Font("Trebuchet MS", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLowRight.Location = New System.Drawing.Point(208, 208)
        Me.btnLowRight.Name = "btnLowRight"
        Me.btnLowRight.Size = New System.Drawing.Size(96, 96)
        Me.btnLowRight.TabIndex = 8
        Me.btnLowRight.TabStop = False
        '
        'btnSinglePlayer
        '
        Me.btnSinglePlayer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSinglePlayer.Font = New System.Drawing.Font("Trebuchet MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSinglePlayer.Location = New System.Drawing.Point(328, 16)
        Me.btnSinglePlayer.Name = "btnSinglePlayer"
        Me.btnSinglePlayer.Size = New System.Drawing.Size(160, 40)
        Me.btnSinglePlayer.TabIndex = 9
        Me.btnSinglePlayer.TabStop = False
        Me.btnSinglePlayer.Text = "Single Player"
        '
        'btnReset
        '
        Me.btnReset.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnReset.Font = New System.Drawing.Font("Trebuchet MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReset.Location = New System.Drawing.Point(328, 232)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(160, 40)
        Me.btnReset.TabIndex = 10
        Me.btnReset.TabStop = False
        Me.btnReset.Text = "Back To Menu"
        Me.btnReset.Visible = False
        '
        'lblPlayer1Score
        '
        Me.lblPlayer1Score.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayer1Score.Font = New System.Drawing.Font("Trebuchet MS", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayer1Score.ForeColor = System.Drawing.Color.Black
        Me.lblPlayer1Score.Location = New System.Drawing.Point(320, 136)
        Me.lblPlayer1Score.Name = "lblPlayer1Score"
        Me.lblPlayer1Score.Size = New System.Drawing.Size(176, 32)
        Me.lblPlayer1Score.TabIndex = 11
        Me.lblPlayer1Score.Text = "Player 1's Score:"
        Me.lblPlayer1Score.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPlayer2Score
        '
        Me.lblPlayer2Score.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayer2Score.Font = New System.Drawing.Font("Trebuchet MS", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayer2Score.ForeColor = System.Drawing.Color.Black
        Me.lblPlayer2Score.Location = New System.Drawing.Point(320, 168)
        Me.lblPlayer2Score.Name = "lblPlayer2Score"
        Me.lblPlayer2Score.Size = New System.Drawing.Size(176, 32)
        Me.lblPlayer2Score.TabIndex = 12
        Me.lblPlayer2Score.Text = "Computer: 0"
        Me.lblPlayer2Score.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnMultiPlayer
        '
        Me.btnMultiPlayer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMultiPlayer.Font = New System.Drawing.Font("Trebuchet MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMultiPlayer.Location = New System.Drawing.Point(328, 72)
        Me.btnMultiPlayer.Name = "btnMultiPlayer"
        Me.btnMultiPlayer.Size = New System.Drawing.Size(160, 40)
        Me.btnMultiPlayer.TabIndex = 13
        Me.btnMultiPlayer.TabStop = False
        Me.btnMultiPlayer.Text = "Multi Player"
        '
        'lblTurn
        '
        Me.lblTurn.BackColor = System.Drawing.SystemColors.Control
        Me.lblTurn.Font = New System.Drawing.Font("Trebuchet MS", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTurn.ForeColor = System.Drawing.Color.Black
        Me.lblTurn.Location = New System.Drawing.Point(320, 200)
        Me.lblTurn.Name = "lblTurn"
        Me.lblTurn.Size = New System.Drawing.Size(176, 32)
        Me.lblTurn.TabIndex = 14
        Me.lblTurn.Text = "Player X's Turn"
        Me.lblTurn.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnClearBoard
        '
        Me.btnClearBoard.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClearBoard.Font = New System.Drawing.Font("Trebuchet MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClearBoard.Location = New System.Drawing.Point(328, 280)
        Me.btnClearBoard.Name = "btnClearBoard"
        Me.btnClearBoard.Size = New System.Drawing.Size(160, 40)
        Me.btnClearBoard.TabIndex = 15
        Me.btnClearBoard.TabStop = False
        Me.btnClearBoard.Text = "Clear Board"
        Me.btnClearBoard.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(10, 25)
        Me.ClientSize = New System.Drawing.Size(506, 324)
        Me.Controls.Add(Me.btnClearBoard)
        Me.Controls.Add(Me.lblTurn)
        Me.Controls.Add(Me.btnMultiPlayer)
        Me.Controls.Add(Me.lblPlayer2Score)
        Me.Controls.Add(Me.lblPlayer1Score)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnSinglePlayer)
        Me.Controls.Add(Me.btnLowRight)
        Me.Controls.Add(Me.btnLowCenter)
        Me.Controls.Add(Me.btnLowLeft)
        Me.Controls.Add(Me.btnMidRight)
        Me.Controls.Add(Me.btnMidCenter)
        Me.Controls.Add(Me.btnMidLeft)
        Me.Controls.Add(Me.btnTopRight)
        Me.Controls.Add(Me.btnTopCenter)
        Me.Controls.Add(Me.btnTopLeft)
        Me.Font = New System.Drawing.Font("Trebuchet MS", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tic Tac Toe"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim turn As String = "Red" 'This variable logs which person's turn it currently is. 'Red' = Player 1 'Blue" = Player 2
    Dim turnCounter As Integer = 1 'This variable keeps track of the turns in a game.
    Dim gameMode As Integer = 0 'This variable tells the program what game mode the user/users are playing.
    Dim playerScores As Integer() = {0, 0} 'This integer array keeps track of each player's score.

    'Sebastian Kaye
    'Mr. Scott
    'Basic Programming Period 6
    'March 17, 2016     Happy St. Patrick's Day!

    'Purpose:
    'This program is a game of tic tac toe. It can either be played singleplayer or multiplayer. The singleplayer will be up against the computer
    'which checks for win opportunities based on the input of the user. The multiplayer begins with player 1 and switches back and forth between
    'player 1 and player 2.

    'This sub checks to see if someone has won the game, congratulates that player, and adds 1 to the player who won the game's score.
    Private Function checkForWin(ByVal button1 As Object, ByVal button2 As Object, ByVal button3 As Object)

        If button1.text = "X" And button2.text = "X" And button3.text = "X" Then
            playerScores(0) += 1
            MsgBox("Player 1 has won!", MsgBoxStyle.OkOnly)
            lblPlayer1Score.Text = "Player 1's Score: " & playerScores(0)
            toggleButtons(False, False)
            resetBoard()
            Return True
        ElseIf button1.text = "O" And button2.text = "O" And button3.text = "O" Then
            playerScores(1) += 1
            'Since, in gamemode 2, the user is playing against the computer, the winning message needs to be different from the default message.
            If gameMode = 1 Then
                MsgBox("Player 2 has won!", MsgBoxStyle.OkOnly)
                lblPlayer2Score.Text = "Player 2's Score: " & playerScores(1)
            ElseIf gameMode = 2 Then
                MsgBox("The Computer has won!", MsgBoxStyle.OkOnly)
                lblPlayer2Score.Text = "Computer: " & playerScores(1)
            End If
            toggleButtons(False, False)
            resetBoard()
            Return True
        End If
        Return False

    End Function

    'This sub sets the text boxes' new text and decides which turn it is.
    Private Sub determineNextTurn()

        'The code below handles a win and stalemate scenerio. If there is no win, then the game will configure the game for the next player.
        turnCounter += 1 'This adds 1 to the current value turnCounter.
        'Explanation for turnCounter:
        '   After the fourth turn, one of the players has put down 2 marks and the other, 3 marks. This is the earliest someone can win the game.
        '   That is when checkForWin() begins checking for one of the 8 win conditions.
        '   If the game has reached 10 turns and there is not winner, a stalemate is declared.
        'This if checks to see if someone has won the game and if a win condition has been found, this sub terminates and the board is reset.
        If turnCounter > 4 Then
            If checkForWin(btnTopLeft, btnMidLeft, btnLowLeft) Then
                Exit Sub
            ElseIf checkForWin(btnTopCenter, btnMidCenter, btnLowCenter) Then
                Exit Sub
            ElseIf checkForWin(btnTopRight, btnMidRight, btnLowRight) Then
                Exit Sub
            ElseIf checkForWin(btnTopLeft, btnTopCenter, btnTopRight) Then
                Exit Sub
            ElseIf checkForWin(btnMidLeft, btnMidCenter, btnMidRight) Then
                Exit Sub
            ElseIf checkForWin(btnLowLeft, btnLowCenter, btnLowRight) Then
                Exit Sub
            ElseIf checkForWin(btnTopLeft, btnMidCenter, btnLowRight) Then
                Exit Sub
            ElseIf checkForWin(btnTopRight, btnMidCenter, btnLowLeft) Then
                Exit Sub
            End If
        End If
        'This if statement checks to see if the game is a stalemate and then resets the game board.
        If turnCounter = 10 Then
            MsgBox("Stalemate", MsgBoxStyle.OkOnly)
            resetBoard()
            Exit Sub
        End If

        'This if statement determines whos turn it is next.
        If turn = "Red" Then
            lblTurn.Text = "Player 2's Turn"
            turn = "Blue"
        ElseIf turn = "Blue" Then
            lblTurn.Text = "Player 1's Turn"
            turn = "Red"
        End If

    End Sub

    'This sub resets the board.
    Private Sub resetBoard()

        'The sub toggleButtons() in this case completely resets the buttons and turnCounter.
        toggleButtons(True, True)
        turnCounter = 1

    End Sub

    'This sub runs resetBoard.
    Private Sub btnClearBoard_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClearBoard.Click

        resetBoard()

    End Sub

    'This sub resets the game to the beginning. All tic tac toe buttons are disabled, the labels are hidden; the only buttons clickable are the game-mode buttons.
    Private Sub reset()

        btnSinglePlayer.Visible = True
        btnMultiPlayer.Visible = True
        btnReset.Visible = False
        btnClearBoard.Visible = False
        lblPlayer1Score.Visible = False
        lblPlayer1Score.Text = "Player 1's Score: 0"
        lblPlayer2Score.Visible = False
        lblPlayer2Score.Text = "Player 2's Score: 0"
        lblTurn.Visible = False
        lblTurn.Text = "Player X's Turn"
        toggleButtons(False, True)
        turnCounter = 1
        playerScores(0) = 0
        playerScores(1) = 0
        gameMode = 0

    End Sub

    'This sub determines the turn for the buttons, sets the button's text to X or O, disables the button for future selection, and calls the nexTurn() sub procedure.
    Private Sub determineTurn(ByVal button As Object)

        If turn = "Red" Then
            button.text = "X"
        ElseIf turn = "Blue" Then
            button.text = "O"
        End If
        button.Enabled = False
        determineNextTurn()
        If gameMode = 2 Then
            computerDecision()
            determineNextTurn()
            turn = "Red"
        End If

    End Sub

    'This sub will toggle the buttons on or off depending on the parameter and resets the text within the boxes depending on the second parameter.
    Private Sub toggleButtons(ByVal toggle As Boolean, ByVal resetText As Boolean)

        btnTopLeft.Enabled = toggle
        btnTopCenter.Enabled = toggle
        btnTopRight.Enabled = toggle
        btnMidLeft.Enabled = toggle
        btnMidCenter.Enabled = toggle
        btnMidRight.Enabled = toggle
        btnLowLeft.Enabled = toggle
        btnLowCenter.Enabled = toggle
        btnLowRight.Enabled = toggle
        'This if statement checks to see if the button's text is to be reset.
        If resetText = True Then
            btnTopLeft.Text = ""
            btnTopCenter.Text = ""
            btnTopRight.Text = ""
            btnMidLeft.Text = ""
            btnMidCenter.Text = ""
            btnMidRight.Text = ""
            btnLowLeft.Text = ""
            btnLowCenter.Text = ""
            btnLowRight.Text = ""
        End If

    End Sub

    'This sub initializes a multiplayer game.
    Private Sub btnMultiPlayer_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnMultiPlayer.Click

        btnMultiPlayer.Visible = False
        btnSinglePlayer.Visible = False
        btnReset.Visible = True
        btnClearBoard.Visible = True
        lblPlayer1Score.Visible = True
        lblPlayer2Score.Visible = True
        lblTurn.Visible = True
        btnReset.Location = New Point(328, 16) '(527, 12), (328, 16)
        btnClearBoard.Location = New Point(328, 72) '(527, 73) (328, 72)
        lblTurn.Text = "Player 1's Turn"
        turn = "Red"
        toggleButtons(True, True)
        gameMode = 1

    End Sub

    'The form runs reset on reset.
    Private Sub btnReset_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReset.Click

        reset()

    End Sub

    'The form runs reset on load.
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        reset()

    End Sub

#Region "Tic Tac Toe Buttons"
    Private Sub btnTopLeft_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnTopLeft.Click

        determineTurn(btnTopLeft)

    End Sub

    Private Sub btnTopCenter_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnTopCenter.Click

        determineTurn(btnTopCenter)

    End Sub

    Private Sub btnTopRight_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnTopRight.Click

        determineTurn(btnTopRight)

    End Sub

    Private Sub btnMidLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMidLeft.Click

        determineTurn(btnMidLeft)

    End Sub

    Private Sub btnMidCenter_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnMidCenter.Click

        determineTurn(btnMidCenter)

    End Sub

    Private Sub btnMidRight_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnMidRight.Click

        determineTurn(btnMidRight)

    End Sub

    Private Sub btnLowLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLowLeft.Click

        determineTurn(btnLowLeft)

    End Sub

    Private Sub btnLowCenter_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLowCenter.Click

        determineTurn(btnLowCenter)

    End Sub

    Private Sub btnLowRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLowRight.Click

        determineTurn(btnLowRight)

    End Sub
#End Region 'This region handles the 9 squares. They call the determineTurn subprocedure.


    'Below are singleplayer specific functions.


    'This sub initializes a single player game.
    Private Sub btnSinglePlayer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSinglePlayer.Click

        btnMultiPlayer.Visible = False
        btnSinglePlayer.Visible = False
        btnReset.Visible = True
        lblPlayer1Score.Visible = True
        lblPlayer2Score.Visible = True
        lblPlayer2Score.Text = "Computer: 0"
        btnReset.Location = New Point(328, 16) '(527, 12), (328, 16)
        toggleButtons(True, True)
        turn = "Red"
        gameMode = 2

    End Sub

    'This sub calculates the computer's decision for its turn. It will check the board to see if the user is in an advantage and attempt to stop the user's win.
    'Basically, it tries to not lose rather than to win.
    Private Sub computerDecision()

        'This long strand of ifs check for all possible opponent win possibilities. If one is found, the computer will mark the board in a strategic place and terminate
        '   this function.
        If checkButton(btnTopLeft, btnTopRight, btnTopCenter) Then '1
            Exit Sub
        ElseIf checkButton(btnTopLeft, btnTopCenter, btnTopRight) Then '2
            Exit Sub
        ElseIf checkButton(btnTopCenter, btnTopRight, btnTopLeft) Then '3
            Exit Sub
        ElseIf checkButton(btnMidLeft, btnMidRight, btnMidCenter) Then '4
            Exit Sub
        ElseIf checkButton(btnMidLeft, btnMidCenter, btnMidRight) Then '5
            Exit Sub
        ElseIf checkButton(btnMidCenter, btnMidRight, btnMidLeft) Then '6
            Exit Sub
        ElseIf checkButton(btnLowLeft, btnLowRight, btnLowCenter) Then '7
            Exit Sub
        ElseIf checkButton(btnLowLeft, btnLowCenter, btnLowRight) Then '8
            Exit Sub
        ElseIf checkButton(btnLowCenter, btnLowRight, btnLowLeft) Then '9
            Exit Sub
        ElseIf checkButton(btnTopLeft, btnLowRight, btnMidCenter) Then '10
            Exit Sub
        ElseIf checkButton(btnTopLeft, btnMidCenter, btnLowRight) Then '11
            Exit Sub
        ElseIf checkButton(btnMidCenter, btnLowRight, btnTopLeft) Then '12
            Exit Sub
        ElseIf checkButton(btnTopRight, btnLowLeft, btnMidCenter) Then '13
            Exit Sub
        ElseIf checkButton(btnTopRight, btnMidCenter, btnLowLeft) Then '14
            Exit Sub
        ElseIf checkButton(btnLowLeft, btnMidCenter, btnTopRight) Then '15
            Exit Sub
        ElseIf checkButton(btnTopLeft, btnLowLeft, btnMidLeft) Then '16
            Exit Sub
        ElseIf checkButton(btnTopLeft, btnMidLeft, btnLowLeft) Then '17
            Exit Sub
        ElseIf checkButton(btnLowLeft, btnMidLeft, btnTopLeft) Then '18
            Exit Sub
        ElseIf checkButton(btnTopCenter, btnLowCenter, btnMidCenter) Then '19
            Exit Sub
        ElseIf checkButton(btnTopCenter, btnMidCenter, btnLowCenter) Then '20
            Exit Sub
        ElseIf checkButton(btnLowCenter, btnMidCenter, btnTopCenter) Then '21
            Exit Sub
        ElseIf checkButton(btnTopRight, btnLowRight, btnMidRight) Then '22
            Exit Sub
        ElseIf checkButton(btnTopRight, btnMidRight, btnLowRight) Then '23
            Exit Sub
        ElseIf checkButton(btnLowRight, btnMidRight, btnTopRight) Then '24
            Exit Sub
        Else
            'In the event that no win possibilites have been found, this code will select a random open box on the board.
            Dim buttons As Object() = {btnTopLeft, btnTopCenter, btnTopRight, btnMidLeft, btnMidCenter, btnMidRight, btnLowLeft, btnLowCenter, btnLowRight}
            Dim foundButtonFlag As Boolean
            Do Until foundButtonFlag = True
                Dim randomNum As Integer = CInt(Math.Floor((8 + 1) * Rnd()))
                If buttons(randomNum).Text = "" And buttons(randomNum).enabled = True Then
                    buttons(randomNum).Text = "O"
                    buttons(randomNum).Enabled = False
                    Exit Sub
                End If
            Loop
        End If

    End Sub

    'This function checks to see if two arbitrary buttons' values are true and selects a third for the computer's selection.
    Private Function checkButton(ByVal button1 As Object, ByVal button2 As Object, ByVal button3 As Object)

        If button1.text = "X" And button2.Text = "X" And button3.Enabled = True Then
            button3.Text = "O"
            button3.Enabled = False
            Return True
        End If
        Return False

    End Function

End Class
