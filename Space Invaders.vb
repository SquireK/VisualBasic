Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Global Variables "

    Dim edgeDetector As Boolean = False 'This variable checks to see if the moving enemies have hit the edge.
    Dim direction As String = "left" 'This variable determines the direction of the enemy movement.
    Dim score As Integer = 0 'This variable keeps the score.
    Dim lives As Integer = 3 'This variable keeps the number of lives left.
    Dim onMenu As Boolean = True 'This variable keeps track of the game state: on the menu or playing the game.
    Dim loadingFinished As Boolean = False 'This variable keeps track of whether or not the level has finished loading.
    Dim houseDestruction(3) As Integer 'This variable keeps the destruction level of each of the houses.
    Dim enemies As PictureBox() = {enemy00, enemy10, enemy20, enemy30, enemy40, enemy50, enemy60, enemy70, enemy80, enemy90, enemyA0, _
                                   enemy01, enemy11, enemy21, enemy31, enemy41, enemy51, enemy61, enemy71, enemy81, enemy91, enemyA1, _
                                   enemy02, enemy12, enemy22, enemy32, enemy42, enemy52, enemy62, enemy72, enemy82, enemy92, enemyA2, _
                                   enemy03, enemy13, enemy23, enemy33, enemy43, enemy53, enemy63, enemy73, enemy83, enemy93, enemyA3, _
                                   enemy04, enemy14, enemy24, enemy34, enemy44, enemy54, enemy64, enemy74, enemy84, enemy94, enemyA4, _
                                   enemy05, enemy15, enemy25, enemy35, enemy45, enemy55, enemy65, enemy75, enemy85, enemy95, enemyA5, enemyBoss}

#End Region

    'The menu will run loadMenu() on program run.
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        loadMenu(True)

    End Sub

    'This function loads or unloads all of the components of the menu and the base components of the game.
    Private Sub loadMenu(ByVal load As Boolean)

        If load Then
            'menu stuff
            lblTitle0.Visible = True
            lblTitle1.Visible = True
            imgPnts0.Visible = True
            imgPnts1.Visible = True
            imgPnts2.Visible = True
            imgPnts3.Visible = True
            lblPnts0.Visible = True
            lblPnts1.Visible = True
            lblPnts2.Visible = True
            lblPnts3.Visible = True
            lblEndGame.Visible = False
            btnBegin.Visible = True
            btnBegin.Text = "Begin"
            btnBegin.Font = New Font("Lucida Console", 24)
            btnQuit.Visible = False
            btnQuit.Size = New Size(128, 48)
            btnQuit.Font = New Font("Lucida Console", 24)
            btnQuit.Location = New Point(376, 408)
            btnQuit.Visible = True
            onMenu = True
            'game stuff
            lblScore.Visible = False
            lblLives.Visible = False
            imgLives0.Visible = False
            imgLives1.Visible = False
            imgLives2.Visible = False
            setupPlayingField(False)
        Else
            'menu stuff
            lblTitle0.Visible = False
            lblTitle1.Visible = False
            imgPnts0.Visible = False
            imgPnts1.Visible = False
            imgPnts2.Visible = False
            imgPnts3.Visible = False
            lblPnts0.Visible = False
            lblPnts1.Visible = False
            lblPnts2.Visible = False
            lblPnts3.Visible = False
            lblEndGame.Visible = False
            btnBegin.Visible = False
            btnQuit.Visible = False
            btnQuit.Font = New Font("Lucida Console", 10)
            btnQuit.Size = New Size(80, 24)
            btnQuit.Location = New Point(24, 448)
            btnQuit.Visible = True
            onMenu = False
            'game stuff
            lblScore.Visible = True
            lblLives.Visible = True
            imgLives0.Visible = True
            imgLives1.Visible = True
            imgLives2.Visible = True
            setupPlayingField(True)
        End If

    End Sub

    'This sub sets up and begins the game.
    Private Sub btnBegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBegin.Click

        loadMenu(False)
        score = 0
        lives = 3

    End Sub

    'This subprocedure runs menuFunctions().
    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuit.Click

        menuFunctions()

    End Sub

    'This subprocedure checks to see if the user is on the menu, then quits the game. If the user is in the game, then the menu will load.
    Private Sub menuFunctions()

        If onMenu Then
            Me.Close()
        Else
            setupPlayingField(False)
            loadMenu(True)
        End If

    End Sub

    'This subprocedure will check and update the number of lives left the player has.
    Private Sub checkLives()

        Select Case lives
            Case 3
                imgLives0.Visible = True
                imgLives1.Visible = True
                imgLives2.Visible = True
            Case 2
                imgLives0.Visible = True
                imgLives1.Visible = True
                imgLives2.Visible = False
            Case 1
                imgLives0.Visible = True
                imgLives1.Visible = False
                imgLives2.Visible = False
            Case 0
                imgLives0.Visible = False
                imgLives1.Visible = False
                imgLives2.Visible = False
        End Select

    End Sub

    'This subprocedure will check the destruction level of the specified house and update the image.
    Private Sub checkDestructionLevels(ByVal house As Integer, ByVal load As Boolean)

        Dim houseList As PictureBox() = {Me.house0, Me.house1, Me.house2, Me.house3}

        Select Case houseDestruction(house)
            Case 0
                houseList(house).Image.FromFile("house0.png")
                If load Then
                    houseList(house).Visible = True
                Else
                    houseList(house).Visible = False
                End If
            Case 1
                houseList(house).Image.FromFile("house1.png")
            Case 2
                houseList(house).Image.FromFile("house2.png")
            Case 3
                houseList(house).Image.FromFile("house3.png")
            Case Is >= 4
                houseList(house).Visible = False
                houseList(house).Image.FromFile("house0.png")
        End Select

    End Sub

    'This subprocedure will setup or unload the playing field.
    Private Sub setupPlayingField(ByVal load As Boolean)

        If load Then
            For c As Integer = 0 To 3 'This for statement loads the houses.
                houseDestruction(c) = 0
                checkDestructionLevels(c, True)
            Next
            imgPlayer.Visible = True 'This makes the player visible.
            imgPlayer.Location = New Point(336, 384)
            tmrEnemyMovement.Start() 'This begins the enemy movement.
        Else
            For c As Integer = 0 To 3 'This resets the houses and makes them dissapear.
                houseDestruction(c) = 0
                checkDestructionLevels(c, False)
            Next
            imgPlayer.Visible = False 'This makes the player invisible.
            imgPlayer.Location = New Point(336, 384)
            tmrEnemyMovement.Stop() 'This stops enemy movement.
            resetEnemies() 'This resets the enemies' positions to off the screen.
        End If


    End Sub

    'This subprocedure will reset the enemy's positions to outside the screen.
    Private Sub resetEnemies()

        Dim loadEnemies As PictureBox() = {enemy00, enemy10, enemy20, enemy30, enemy40, enemy50, enemy60, enemy70, enemy80, enemy90, enemyA0, _
                                           enemy01, enemy11, enemy21, enemy31, enemy41, enemy51, enemy61, enemy71, enemy81, enemy91, enemyA1, _
                                           enemy02, enemy12, enemy22, enemy32, enemy42, enemy52, enemy62, enemy72, enemy82, enemy92, enemyA2, _
                                           enemy03, enemy13, enemy23, enemy33, enemy43, enemy53, enemy63, enemy73, enemy83, enemy93, enemyA3, _
                                           enemy04, enemy14, enemy24, enemy34, enemy44, enemy54, enemy64, enemy74, enemy84, enemy94, enemyA4, _
                                           enemy05, enemy15, enemy25, enemy35, enemy45, enemy55, enemy65, enemy75, enemy85, enemy95, enemyA5}
        Dim enemyOrigins As Integer() = {728, 238, 776, 238, 824, 238, 872, 238, 920, 238, 968, 238, 1016, 238, 1064, 238, 1112, 238, 1160, 238, 1208, 238, _
                                         728, 208, 776, 208, 824, 208, 872, 208, 920, 208, 968, 208, 1016, 208, 1064, 208, 1112, 208, 1160, 208, 1208, 208, _
                                         728, 176, 776, 176, 824, 176, 872, 176, 920, 176, 968, 176, 1016, 176, 1064, 176, 1112, 176, 1160, 176, 1208, 176, _
                                         728, 152, 776, 152, 824, 152, 872, 152, 920, 152, 968, 152, 1016, 152, 1064, 152, 1112, 152, 1160, 152, 1208, 152, _
                                         728, 120, 776, 120, 824, 120, 872, 120, 920, 120, 968, 120, 1016, 120, 1064, 120, 1112, 120, 1160, 120, 1208, 120, _
                                         728, 96, 776, 96, 824, 96, 872, 96, 920, 96, 968, 96, 1016, 96, 1064, 96, 1112, 96, 1160, 96, 1208, 96}

        For c As Integer = 0 To 130 Step 2
            loadEnemies(c / 2).Location = New Point(enemyOrigins(c), enemyOrigins(c + 1))
            loadEnemies(c / 2).Visible = True
        Next


    End Sub

    'This subprocdure holds code to unload the playing field and tell the player whether they won or lost.
    'This is fine without the shooting code, but it will need to be changed later.
    Private Sub endGame(ByVal condition As Integer)

        If condition = 1 Then
            menuFunctions()
            imgPnts0.Visible = False
            imgPnts1.Visible = False
            imgPnts2.Visible = False
            imgPnts3.Visible = False
            lblPnts0.Visible = False
            lblPnts1.Visible = False
            lblPnts2.Visible = False
            lblPnts3.Visible = False
            lblEndGame.Text = "You Lost!"
            lblEndGame.Visible = True
            btnBegin.Font = New Font("Lucida Console", 20)
            btnBegin.Text = "Replay"
        End If

    End Sub

    'This subprocedure handles the movement of the enemies.TO DO: movement along the Y axis
    Private Sub tmrEnemyMovement_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrEnemyMovement.Tick

        Dim loadEnemies As PictureBox() = {enemy00, enemy10, enemy20, enemy30, enemy40, enemy50, enemy60, enemy70, enemy80, enemy90, enemyA0, _
                                           enemy01, enemy11, enemy21, enemy31, enemy41, enemy51, enemy61, enemy71, enemy81, enemy91, enemyA1, _
                                           enemy02, enemy12, enemy22, enemy32, enemy42, enemy52, enemy62, enemy72, enemy82, enemy92, enemyA2, _
                                           enemy03, enemy13, enemy23, enemy33, enemy43, enemy53, enemy63, enemy73, enemy83, enemy93, enemyA3, _
                                           enemy04, enemy14, enemy24, enemy34, enemy44, enemy54, enemy64, enemy74, enemy84, enemy94, enemyA4, _
                                           enemy05, enemy15, enemy25, enemy35, enemy45, enemy55, enemy65, enemy75, enemy85, enemy95, enemyA5}

        If Not edgeDetector Then
            'If the enemy is moving left, keep it moving left.
            If direction = "left" Then
                For c As Integer = 0 To 65
                    loadEnemies(c).Location = New Point(loadEnemies(c).Location.X - 15, loadEnemies(c).Location.Y)
                Next
                'If the enemy gets to the left constant limit, have it turn around.
                If enemyA5.Location.X <= 520 Then
                    direction = "right"
                    edgeDetector = True
                End If
                'If the enemy is moving right, keeep it moving right.
            ElseIf direction = "right" Then
                For c As Integer = 0 To 65
                    loadEnemies(c).Location = New Point(loadEnemies(c).Location.X + 15, loadEnemies(c).Location.Y)
                Next
                'If the enemy gets to the right constant limit, have it turn around.
                If enemyA5.Location.X >= 632 Then
                    direction = "left"
                    edgeDetector = True
                End If
            End If
        ElseIf edgeDetector Then

            For c As Integer = 0 To 65
                loadEnemies(c).Location = New Point(loadEnemies(c).Location.X, loadEnemies(c).Location.Y + 5)
            Next
            edgeDetector = False

        End If
        If enemyA5.Location.Y >= 168 Then
            endGame(1)
        End If

    End Sub

    Private Sub detectKeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        If e.KeyCode = Keys.A Then
            imgPlayer.Location = New Point(imgPlayer.Location.X + 50, imgPlayer.Location.Y)
        ElseIf e.KeyCode = Keys.D Then
            imgPlayer.Location = New Point(imgPlayer.Location.X - 50, imgPlayer.Location.Y)
        End If

    End Sub
End Class
