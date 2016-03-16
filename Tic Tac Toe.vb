    Public Class frmMain
    
    Dim turn As String = "Red" 'This variable logs which person's turn it currently is. 'Red' = Player 1 'Blue" = Player 2
    Dim turnCounter As Integer = 1 'This variable keeps track of the turns in a game.
    Dim gameMode As Integer = 0
    Dim playerScores As Integer() = {0, 0} 'This integer array keeps track of each player's score.


    'This sub checks to see if someone has won the game, congratulates that player, and adds 1 to the player who won the game's score.
    Private Function checkForWin(ByVal button1 As Object, ByVal button2 As Object, ByVal button3 As Object)

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
            If checkForWin(btnTopLeft, btnMidLeft, btnLowLeft) Or _
                checkForWin(btnTopCenter, btnMidCenter, btnLowCenter) Or _
                checkForWin(btnTopRight, btnMidRight, btnLowRight) Or _
                checkForWin(btnTopLeft, btnTopCenter, btnTopRight) Or _
                checkForWin(btnMidLeft, btnMidCenter, btnMidRight) Or _
                checkForWin(btnLowLeft, btnLowCenter, btnLowRight) Or _
                checkForWin(btnTopLeft, btnMidCenter, btnLowRight) Or _
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
        If gameMode = 1 Then
            determineNextTurn()
        ElseIf gameMode = 2 Then
            turn = "Blue"
            computerDecision()
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

    'This sub sets up a multiplayer game.
    Private Sub btnMultiPlayer_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnMultiPlayer.Click

        btnMultiPlayer.Visible = False
        btnSinglePlayer.Visible = False
        btnReset.Visible = True
        btnClearBoard.Visible = True
        lblPlayer1Score.Visible = True
        lblPlayer2Score.Visible = True
        lblTurn.Visible = True
        btnReset.Location = New Point(328, 16)
        btnClearBoard.Location = New Point(328, 72)
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

    'This sub initializes the single player game.
    Private Sub btnSinglePlayer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSinglePlayer.Click

        btnMultiPlayer.Visible = False
        btnSinglePlayer.Visible = False
        btnReset.Visible = True
        lblPlayer1Score.Visible = True
        lblPlayer2Score.Visible = True
        lblPlayer2Score.Text = "Computer: 0"
        btnReset.Location = New Point(328, 16)
        toggleButtons(True, True)
        turn = "Red"
        gameMode = 2

    End Sub

    Private Sub computerDecision()

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
            Dim buttons As Object() = {btnTopLeft, btnTopCenter, btnTopRight, btnMidLeft, btnMidCenter, btnMidRight, btnLowLeft, btnLowCenter, btnLowRight}
            For counter As Integer = 0 To 8
                If buttons(counter).Text = "" And buttons(counter).enabled = True Then
                    buttons(counter).Text = "O"
                    buttons(counter).Enabled = False
                    turn = "Red"
                    Exit Sub
                End If
            Next
        End If

    End Sub

    Private Function checkButton(ByVal button1 As Object, ByVal button2 As Object, ByVal button3 As Object)

        If button1.text = "X" And button2.Text = "X" And button3.Enabled = True Then
            button3.Text = "O"
            button3.Enabled = False
            turn = "Red"
            Return True
        End If
        Return False

    End Function

End Class
