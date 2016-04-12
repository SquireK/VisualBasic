Public Class frmMain

    Dim cardCount As Integer = 0
    Dim cardSelection(1) As String
    Dim playerScores(1) As Integer
    Dim cardCovers(1) As Label
    Dim imageLocations(23) As String
    Dim running As Boolean = False

    Private Function assignPictureName(ByVal picturebox As PictureBox)
        Select Case picturebox.Name
            Case "img1A"
                Return imageLocations(0)
            Case "img1B"
                Return imageLocations(1)
            Case "img1C"
                Return imageLocations(2)
            Case "img1D"
                Return imageLocations(3)
            Case "img1E"
                Return imageLocations(4)
            Case "img1F"
                Return imageLocations(5)
            Case "img2A"
                Return imageLocations(6)
            Case "img2B"
                Return imageLocations(7)
            Case "img2C"
                Return imageLocations(8)
            Case "img2D"
                Return imageLocations(9)
            Case "img2E"
                Return imageLocations(10)
            Case "img2F"
                Return imageLocations(11)
            Case "img3A"
                Return imageLocations(12)
            Case "img3B"
                Return imageLocations(13)
            Case "img3C"
                Return imageLocations(14)
            Case "img3D"
                Return imageLocations(15)
            Case "img3E"
                Return imageLocations(16)
            Case "img3F"
                Return imageLocations(17)
            Case "img4A"
                Return imageLocations(18)
            Case "img4B"
                Return imageLocations(19)
            Case "img4C"
                Return imageLocations(20)
            Case "img4D"
                Return imageLocations(21)
            Case "img4E"
                Return imageLocations(22)
            Case "img4F"
                Return imageLocations(23)
        End Select
    End Function

    Private Sub reset()

        lblPlayer1Score.Text = "Player 1 Score: 0"
        lblPlayer2Score.Text = "Player 2 Score: 0"
        playerScores(0) = 0
        playerScores(1) = 0
        lblTurn.Text = "Player 1's Turn"
        img1ACover.Visible = True
        img1BCover.Visible = True
        img1CCover.Visible = True
        img1DCover.Visible = True
        img1ECover.Visible = True
        img1FCover.Visible = True
        img2ACover.Visible = True
        img2BCover.Visible = True
        img2CCover.Visible = True
        img2DCover.Visible = True
        img2ECover.Visible = True
        img2FCover.Visible = True
        img3ACover.Visible = True
        img3BCover.Visible = True
        img3CCover.Visible = True
        img3DCover.Visible = True
        img3ECover.Visible = True
        img3FCover.Visible = True
        img4ACover.Visible = True
        img4BCover.Visible = True
        img4CCover.Visible = True
        img4DCover.Visible = True
        img4ECover.Visible = True
        img4FCover.Visible = True
        randomizeImages()

    End Sub

    'This subprocedure randomizes the image locations on the board.
    Private Sub randomizeImages()

        'This array contains boolean values that store information about the cards and the images assigned to them.
        Dim cards() As Boolean = {False, False, False, False, False, False, _
                                    False, False, False, False, False, False, _
                                    False, False, False, False, False, False, _
                                    False, False, False, False, False, False}
        'This array contains the name values of the images.
        Dim images() As String = {"image1.png", "image2.png", "image3.png", "image4.png", "image5.png", "image6.png", _
                                  "image7.png", "image8.png", "image9.png", "image10.png", "image11.png", "image12.png", _
                                  "image1.png", "image2.png", "image3.png", "image4.png", "image5.png", "image6.png", _
                                  "image7.png", "image8.png", "image9.png", "image10.png", "image11.png", "image12.png"}
        'This array contains the picture boxes of the program.
        Dim boardCards() As PictureBox = {Me.img1A, Me.img1B, Me.img1C, Me.img1D, Me.img1E, Me.img1F, _
                                          Me.img2A, Me.img2B, Me.img2C, Me.img2D, Me.img2E, Me.img2F, _
                                          Me.img3A, Me.img3B, Me.img3C, Me.img3D, Me.img3E, Me.img3F, _
                                          Me.img4A, Me.img4B, Me.img4C, Me.img4D, Me.img4E, Me.img4F}
        'This variable is the counter for the For statement below.
        Dim cardId As Integer

        'This for statement assigns a random image to each of the picture boxes.
        For cardId = 0 To 23
generateNewRandNum:
            Dim randNum As Integer = CInt(Math.Floor((23 + 1) * Rnd()))

            If cards(randNum) Then
                GoTo generateNewRandNum
            Else
                boardCards(randNum).Image = Image.FromFile(images(cardId))
                imageLocations(randNum) = images(cardId)
                cards(randNum) = True
            End If
        Next

    End Sub

    Private Sub determineNextTurn(ByVal card As PictureBox, ByVal cardCover As Label)

        If running = True Then
            If cardCount = 2 Then
                cardCount = 0
            End If
            cardCount += 1
            If cardCount = 1 Then
                cardSelection(0) = assignPictureName(card)
                cardCovers(0) = cardCover
                cardCovers(0).Visible = False
            ElseIf cardCount = 2 Then
                cardSelection(1) = assignPictureName(card)
                cardCovers(1) = cardCover
                cardCovers(1).Visible = False

                If cardSelection(0) = cardSelection(1) Then
                    If lblTurn.Text = "Player 1's Turn" Then
                        playerScores(0) += 1
                        lblPlayer1Score.Text = "Player 1 Score: " & playerScores(0)
                        lblTurn.Text = "Player 2's Turn"
                        cardCount = 0
                    Else
                        playerScores(1) += 1
                        lblPlayer2Score.Text = "Player 2 Score: " & playerScores(1)
                        lblTurn.Text = "Player 1's Turn"
                        cardCount = 0
                    End If
                Else
                    cardCovers(0).Visible = True
                    cardCovers(1).Visible = True
                    If lblTurn.Text = "Player 1's Turn" Then
                        If btnHardMode.Checked Then
                            playerScores(0) -= 1
                            lblPlayer1Score.Text = "Player 1 Score: " & playerScores(0)
                            lblTurn.Text = "Player 2's Turn"
                            cardCount = 0
                        Else
                            lblTurn.Text = "Player 2's Turn"
                            cardCount = 0
                        End If
                    ElseIf lblTurn.Text = "Player 2's Turn" Then
                        If btnHardMode.Checked Then
                            playerScores(1) -= 1
                            lblPlayer2Score.Text = "Player 2 Score: " & playerScores(1)
                            lblTurn.Text = "Player 1's Turn"
                            cardCount = 0
                        Else
                            lblTurn.Text = "Player 1's Turn"
                            cardCount = 0
                        End If
                    End If
                End If
            End If
            If playerScores(0) + playerScores(1) = 12 Then
                btnToggle.Text = "Replay"
            End If
        End If

    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        reset()
        lblTurn.Visible = False

    End Sub

    Private Sub btnToggle_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnToggle.Click

        reset()
        lblTurn.Visible = True
        btnToggle.Text = "Reset"
        running = True

    End Sub

#Region " Image Covers "
    Private Sub img1ACover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img1ACover.Click
        determineNextTurn(img1A, img1ACover)
    End Sub

    Private Sub img1BCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img1BCover.Click
        determineNextTurn(img1B, img1BCover)
    End Sub

    Private Sub img1CCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img1CCover.Click
        determineNextTurn(img1C, img1CCover)
    End Sub

    Private Sub img1DCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img1DCover.Click
        determineNextTurn(img1D, img1DCover)
    End Sub

    Private Sub img1ECover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img1ECover.Click
        determineNextTurn(img1E, img1ECover)
    End Sub

    Private Sub img1FCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img1FCover.Click
        determineNextTurn(img1F, img1FCover)
    End Sub

    Private Sub img2ACover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img2ACover.Click
        determineNextTurn(img2A, img2ACover)
    End Sub

    Private Sub img2BCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img2BCover.Click
        determineNextTurn(img2B, img2BCover)
    End Sub

    Private Sub img2CCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img2CCover.Click
        determineNextTurn(img2C, img2CCover)
    End Sub

    Private Sub img2DCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img2DCover.Click
        determineNextTurn(img2D, img2DCover)
    End Sub

    Private Sub img2ECover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img2ECover.Click
        determineNextTurn(img2E, img2ECover)
    End Sub

    Private Sub img2FCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img2FCover.Click
        determineNextTurn(img2F, img2FCover)
    End Sub

    Private Sub img3ACover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img3ACover.Click
        determineNextTurn(img3A, img3ACover)
    End Sub

    Private Sub img3BCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img3BCover.Click
        determineNextTurn(img3B, img3BCover)
    End Sub

    Private Sub img3CCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img3CCover.Click
        determineNextTurn(img3C, img3CCover)
    End Sub

    Private Sub img3DCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img3DCover.Click
        determineNextTurn(img3D, img3DCover)
    End Sub

    Private Sub img3ECover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img3ECover.Click
        determineNextTurn(img3E, img3ECover)
    End Sub

    Private Sub img3FCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img3FCover.Click
        determineNextTurn(img3F, img3FCover)
    End Sub

    Private Sub img4ACover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img4ACover.Click
        determineNextTurn(img4A, img4ACover)
    End Sub

    Private Sub img4BCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img4BCover.Click
        determineNextTurn(img4B, img4BCover)
    End Sub

    Private Sub img4CCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img4CCover.Click
        determineNextTurn(img4C, img4CCover)
    End Sub

    Private Sub img4DCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img4DCover.Click
        determineNextTurn(img4D, img4DCover)
    End Sub

    Private Sub img4ECover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img4ECover.Click
        determineNextTurn(img4E, img4ECover)
    End Sub

    Private Sub img4FCover_Click(ByVal sender As Object, ByVal e As EventArgs) Handles img4FCover.Click
        determineNextTurn(img4F, img4FCover)
    End Sub
#End Region

End Class
