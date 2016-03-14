Public Class frmMain

    Dim turn As String = "Red" 'This variable logs which person's turn it currently is. 'Red' = Player 1 'Blue" = Player 2
    Dim turnCounter As Integer = 1 'This variable keeps track of the turns in a game.
    Dim playerScores As Integer() = {0, 0} 'This integer array keeps track of each player's score.


    'This sub checks to see if someone has won the game, congratulates that player, and adds 1 to the player who won the game's score.
    Private Function checkForWin(button1 As Object, button2 As Object, button3 As Object)

        If button1.text = "X" And button2.text = "X" And button3.text = "X" Then
            MsgBox("Player 1 has won!", MsgBoxStyle.OkOnly)
            toggleButtons(False, False)
            playerScores(0) += 1
            lblPlayer1Score.Text = "Player 1's Score: " & playerScores(0)
            Return True
        ElseIf button1.text = "O" And button2.text = "O" And button3.text = "O" Then
            MsgBox("Player 2 has won!", MsgBoxStyle.OkOnly)
            toggleButtons(False, False)
            playerScores(1) += 1
            lblPlayer2Score.Text = "Player 2's Score: " & playerScores(1)
            Return True
        End If
        Return False

    End Function


    'This sub sets the text boxes' new text and decides which turn it is.
    Private Sub determineNextTurn()

        'This if statement detemine's whos turn it is.
        If lblTurn.Text = "Player 1's Turn" And turn = "Red" Then
            lblTurn.Text = "Player 2's Turn"
            turn = "Blue"
        ElseIf lblTurn.Text = "Player 2's Turn" And turn = "Blue" Then
            lblTurn.Text = "Player 1's Turn"
            turn = "Red"
        End If

        'The code below handles a win and stalemate scenerio.
        turnCounter += 1 'This adds 1 to the current value turnCounter
        'This if checks to see if someone has won the game and then resets the game board.
        If turnCounter > 4 Then
            If checkForWin(btnTopLeft, btnMidLeft, btnLowLeft) Or
                checkForWin(btnTopCenter, btnMidCenter, btnLowCenter) Or
                checkForWin(btnTopRight, btnMidRight, btnLowRight) Or
                checkForWin(btnTopLeft, btnTopCenter, btnTopRight) Or
                checkForWin(btnMidLeft, btnMidCenter, btnMidRight) Or
                checkForWin(btnLowLeft, btnLowCenter, btnLowRight) Or
                checkForWin(btnTopLeft, btnMidCenter, btnLowRight) Or
                checkForWin(btnTopRight, btnMidCenter, btnLowLeft) = True Then
                resetBoard()
                Exit Sub
            End If
        End If
        'This if statement checks to see if the game is a stalemate and then resets the game board.
        If turnCounter = 10 Then
            MsgBox("Stalemate", MsgBoxStyle.OkOnly)
            resetBoard()
        End If

    End Sub

    'This sub resets the board.
    Private Sub resetBoard()

        toggleButtons(True, True)
        turnCounter = 1

    End Sub

    'This sub runs resetBoard.
    Private Sub btnClearBoard_Click(sender As Object, e As EventArgs) Handles btnClearBoard.Click

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

    End Sub

    'This sub determines the turn for the buttons, sets the button's text to X or O, disables the button for future selection, and calls the nexTurn() sub procedure.
    Private Sub determineTurn(button As Object)

        If turn = "Red" Then
            button.text = "X"
        ElseIf turn = "Blue" Then
            button.text = "O"
        End If
        button.Enabled = False
        determineNextTurn()

    End Sub

    'This sub will toggle the buttons on or off depending on the parameter and resets the text within the boxes depending on the second parameter.
    Private Sub toggleButtons(toggle As Boolean, resetText As Boolean)

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


    'This sub sets up a multiplayer game.
    Private Sub btnMultiPlayer_Click(sender As Object, e As EventArgs) Handles btnMultiPlayer.Click

        btnMultiPlayer.Visible = False
        btnSinglePlayer.Visible = False
        btnReset.Visible = True
        btnClearBoard.Visible = True
        lblPlayer1Score.Visible = True
        lblPlayer2Score.Visible = True
        lblTurn.Visible = True
        btnReset.Location = New Point(527, 12)
        btnClearBoard.Location = New Point(527, 73)
        lblTurn.Text = "Player 1's Turn"
        turn = "Red"
        toggleButtons(True, True)

    End Sub

    'The form runs reset on reset.
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click

        reset()

    End Sub

    'The form runs reset on load.
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        reset()

    End Sub

#Region "Tic Tac Toe Buttons"
    Private Sub btnTopLeft_Click(sender As Object, e As EventArgs) Handles btnTopLeft.Click

        determineTurn(btnTopLeft)

    End Sub

    Private Sub btnTopCenter_Click(sender As Object, e As EventArgs) Handles btnTopCenter.Click

        determineTurn(btnTopCenter)

    End Sub

    Private Sub btnTopRight_Click(sender As Object, e As EventArgs) Handles btnTopRight.Click

        determineTurn(btnTopRight)

    End Sub

    Private Sub btnMidLeft_Click(sender As Object, e As EventArgs) Handles btnMidLeft.Click

        determineTurn(btnMidLeft)

    End Sub

    Private Sub btnMidCenter_Click(sender As Object, e As EventArgs) Handles btnMidCenter.Click

        determineTurn(btnMidCenter)

    End Sub

    Private Sub btnMidRight_Click(sender As Object, e As EventArgs) Handles btnMidRight.Click

        determineTurn(btnMidRight)

    End Sub

    Private Sub btnLowLeft_Click(sender As Object, e As EventArgs) Handles btnLowLeft.Click

        determineTurn(btnLowLeft)

    End Sub

    Private Sub btnLowCenter_Click(sender As Object, e As EventArgs) Handles btnLowCenter.Click

        determineTurn(btnLowCenter)

    End Sub

    Private Sub btnLowRight_Click(sender As Object, e As EventArgs) Handles btnLowRight.Click

        determineTurn(btnLowRight)

    End Sub
#End Region 'This region handles the 9 squares. They call the determineTurn subprocedure.

End Class
