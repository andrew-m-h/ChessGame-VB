Public Class frmGame
#Region "dynamic variables"
    'These are objects are used in the form creation
    Dim picbox1, picbox2, picbox3, picbox4, picbox5, picbox6, picbox7, picbox8, picbox9, picbox10, picbox11, picbox12, picbox13, picbox14, picbox15, picbox16, picbox17, picbox18, picbox19, picbox20, _
picbox21, picbox22, picbox23, picbox24, picbox25, picbox26, picbox27, picbox28, picbox29, picbox30, picbox31, picbox32, picbox33, picbox34, picbox35, picbox36, picbox37, picbox38, picbox39, picbox40, _
picbox41, picbox42, picbox43, picbox44, picbox45, picbox46, picbox47, picbox48, picbox49, picbox50, picbox51, picbox52, picbox53, picbox54, picbox55, picbox56, picbox57, picbox58, picbox59, picbox60, _
picbox61, picbox62, picbox63, picbox64 As New PictureBox 'used when dynamically creating the board
    Dim lstScores As New ListBox 'these list boxes display a list of the pieces that have been lost by each player
    Dim lstP1Grave, lstP2Grave As New ListBox
    Dim btnRestart, btnQuit, btnHelp, btnSave, btnClear As New Button 'these buttons are used to perform various tasks
    Dim lblA, lblB, lblC, lblD, lblE, lblF, lblG, lblH, lbl1, lbl2, lbl3, lbl4, lbl5, lbl6, lbl7, lbl8 As New Label 'labels used to display co-ordinates
    Dim lblP1Score, lblP2Score As New Label 'display the scores of two players
    Dim lblTurn As New Label 'tells the players who’s turn it is
    Dim lblP1Heading, lblP2heading As New Label
    Dim lblScoresHeading As New Label
#End Region

#Region "global variables"
    'These variables are global throughout this form
    Dim origin, destination, winner As String
    Dim click_counter As Integer
    Dim title As String
    Dim picbox_array(8, 8) As PictureBox
    Dim label_array(8, 2) As Label
    Dim game_ended As Boolean = False
#End Region

#Region "public Variables"
    'These variables are public throughout the program
    Public board(8, 8) As String
    Public turn As String
    Public p1score, p2score As Integer
    Public p1name, p2name As String
#End Region

    Private Structure people
        'This record holds the name and score of one previous player
        'It will be part of an array of records to hold the scores and names of all previous players
        Dim score As Integer
        Dim name As String
    End Structure

    Private Sub frmGame_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        'This activates when the form loads
        'Builds the form and the controls, fills the board array and outputs the board onto the screen
        fill_control_arrays()                   'Fills the control arrays that are used to group screen elements
        create_board()                          'Dynamically creates the board for ease of manipulation
        title = frmNewGame.game_title
        If Not frmLoad.game_loaded Then
            fill_board()                        'Fills the board array with the starting layout
            save_game(title)                    'This saves the new game to a game file in its starting position
            turn = "p1"                         'sets the starting value of turn
        Else
            If frmLoad.p1grave.Length <> 0 Then 'This only activates if the public grave array is not empty
                For Each i In frmLoad.p1grave   'This cycles through the array and puts the contents into the appropriate list box
                    update_grave(i)
                Next
            End If
            If frmLoad.p2grave.Length <> 0 Then
                For Each i In frmLoad.p2grave
                    update_grave(i)
                Next
                lblP1Score.Text = p1name & " : " & p1score
                lblP2Score.Text = p2name & " : " & p2score

                'lblTurn.Text = "Turn : Player " & Mid(turn, 2, 1)
            End If
        End If
        Select Case turn
            Case "p1" : lblTurn.Text = "Turn : " & p1name
            Case "p2" : lblTurn.Text = "Turn : " & p2name
        End Select
        display_board()                         'outputs the board in its starting position
    End Sub

    Private Sub Action_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This sub is activated by clicking any of the picture boxes that make up the board
        Dim clicked As PictureBox                   'clicked will be the picbox that was clicked
        clicked = sender                            'assigns clicked to the picture box that activated the sub
        clicked.BorderStyle = BorderStyle.Fixed3D   'highlights the selected tile
        click_counter = click_counter + 1
        If click_counter = 1 Then
            origin = clicked.Name                   'This sets the origin variable to the value of the first tile clicked click
            get_all_moves(origin)
        End If
        If click_counter = 2 Then
            destination = clicked.Name
            make_move()                             'makes the move given both the origin and destination picboxes
            click_counter = 0                       'resets click counter to 0
            swap(turn)                              'switches the players turn
            If game_ended Then
                restart_game()                      'This resets the board and deals with restarts
            End If
        End If
    End Sub

    Private Sub Save_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        save_game(title)            'This re-writes the game file with the new data
        MsgBox("Game Saved")        'This outputs a text box that alerts the player that the game has been saved
    End Sub

    Private Sub Restart_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        restart_game()              'This will reset the game board and the game
        save_game(title)            'This saves the game in its restarted mode
    End Sub

    Private Sub Help_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        MsgBox("This'll learn ya!") '
    End Sub

    Private Sub Quit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This redirects the user to the main form instead of actually quitting
        frmLoad.game_loaded = False 'This sets the game_loaded boolean to false for later use
        frmMain.Show()
        Me.Close()
    End Sub

    Private Sub Clear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This button is used to clear the high scores list box
        lstScores.Items.Clear()
    End Sub

    Private Sub swap(ByRef turn)
        'This sub is used to change the turn variable
        If turn = "p1" Then
            turn = "p2"
        Else
            turn = "p1"
        End If
    End Sub

    Private Sub create_board()
        'This sub programmatically creates the 8*8 square board of picture boxes
        'This is quicker and easier than manual creation
        'Allows for easy bulk manipulation of the picture boxes

        'The nature of this strategy means that screen resolutions are a concern with apps designed on a 1080x1920 screen occupying more space on a 720x1280 scree
        'This problem is resolved through the use of ratios in the x and y axis
        'These ratios are then multiplied by all the constant terms to adjust the screen size for screens other than 1080x1920 pixles
        Dim x_ratio As Single = Screen.PrimaryScreen.Bounds.Width       'get the width of the screen in pixels
        Dim y_ratio As Single = Screen.PrimaryScreen.Bounds.Height      'get the height of the screen in pixels
        Me.Size = New System.Drawing.Size(0.68 * x_ratio, 0.96 * y_ratio) 'sets the size of the form in relation to the size of the screen  
        Me.Location = New System.Drawing.Point(0.15 * x_ratio, 0)       'sets the location on the screen where the form loads
        x_ratio = x_ratio / 1920                                        'set the ratios size of screen : size of 1080x1920 screen
        y_ratio = y_ratio / 1080
        create_tiles(x_ratio, y_ratio)                                  'creates the tiles that make up the board
        create_buttons(x_ratio, y_ratio)                                'creates the buttons that are used on the form
        create_labels(x_ratio, y_ratio)                                 'creates the labels that run along the edge of the board
        create_listboxes(x_ratio, y_ratio)                              'creates the list boxes that are used on the form
    End Sub

    Private Sub create_tiles(x_ratio, y_ratio)
        'Create the 8 * 8 board of picture boxes that are used to display the pieces
        'The boxes are in a picturebox array for ease of manipulation

        'All constant terms must be multiplied by the appropriate ratio
        Dim x_coordinate, y_coordinate As Integer
        For y = 1 To 8                                                  'Cycles through the two dimensional array "picbox_array" and defines the objects within
            For x = 1 To 8
                x_coordinate = 5 * x_ratio * (19 * x + 35)              'This separates the picture boxes by 95 pixels in the x axis
                y_coordinate = 5 * y_ratio * (19 * y + 3)               'This separates the picture boxes by 95 pixels in the y axis
                With picbox_array(x, y)                                 'for some reason, VB uses (y, x) instead of (x, y)
                    .Name = "picTile" & (8 * y - 8 + x)                 'This will name the picture box "picTile" + n where n is an integer
                    .Size = New System.Drawing.Point(95 * x_ratio, 95 * y_ratio)    'The size will be all the same, almost square
                    .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
                    .SizeMode = PictureBoxSizeMode.StretchImage
                End With
                Me.Controls.Add(picbox_array(x, y))                     'Add picbox to the controls
                AddHandler picbox_array(x, y).Click, AddressOf Action_Click 'link button click to action click sub
            Next x
        Next y
    End Sub

    Private Sub create_buttons(x_ratio, y_ratio)
        'Creates the buttons that are used on the form
        'This sub must add click events as well as define the buttons features

        'All constant terms must be multiplied by the appropriate ratio
        Dim x_coordinate, y_coordinate As Integer
        y_coordinate = 900 * y_ratio                                        'The y co-ordinate is constant for all the buttons
        x_coordinate = 955 * x_ratio
        With btnSave                                                        'This defines the button btnSave
            .Size = New System.Drawing.Size(75 * x_ratio, 35 * y_ratio)
            .Text = "Save"
            .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
            .Font = New System.Drawing.Font("standard", Int(11 * x_ratio), FontStyle.Regular)
        End With
        x_coordinate = x_coordinate + 80 * x_ratio                          'The x_coordinate is incremented by 80 to allow for a 5 pixle gap
        With btnHelp                                                        'This defines the button btnHelp
            .Size = New System.Drawing.Size(75 * x_ratio, 35 * y_ratio)
            .Text = "Help"
            .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
            .Font = New System.Drawing.Font("standard", Int(11 * x_ratio), FontStyle.Regular)
        End With
        x_coordinate = x_coordinate + 80 * x_ratio
        With btnRestart                                                     'This defines the button btnRestart
            .Size = New System.Drawing.Size(75 * x_ratio, 35 * y_ratio)
            .Text = "Restart"
            .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
            .Font = New System.Drawing.Font("standard", Int(10 * x_ratio), FontStyle.Regular)
        End With
        x_coordinate = x_coordinate + 80 * x_ratio
        With btnQuit                                                        'This defines the button btnQuit
            .Size = New System.Drawing.Size(75 * x_ratio, 35 * y_ratio)
            .Text = "Quit"
            .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
            .Font = New System.Drawing.Font("standard", Int(11 * x_ratio), FontStyle.Regular)
        End With

        'Add buttons to controls
        Me.Controls.Add(btnSave)
        Me.Controls.Add(btnHelp)
        Me.Controls.Add(btnRestart)
        Me.Controls.Add(btnQuit)
        'assign buttons to click events
        AddHandler btnSave.Click, AddressOf Save_Click
        AddHandler btnHelp.Click, AddressOf Help_Click
        AddHandler btnRestart.Click, AddressOf Restart_Click
        AddHandler btnQuit.Click, AddressOf Quit_Click

        'This defines the clear button that is used to clear the high scores list box
        With btnClear
            .Location = New System.Drawing.Point(1120 * x_ratio, 620 * y_ratio)
            .Size = New System.Drawing.Size(75 * x_ratio, 35 * y_ratio)
            .Text = "Clear"
            .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
        End With
        'Add button to controls
        Me.Controls.Add(btnClear)
        'Add click event
        AddHandler btnClear.Click, AddressOf Clear_Click
    End Sub

    Private Sub create_labels(x_ratio, y_ratio)
        'Create the labels that go across the top and down the side of the board
        'letters go across the top of the board
        Dim x_coordinate, y_coordinate As Integer
        For x = 1 To 8
            x_coordinate = 5 * x_ratio * (19 * x + 42)                                       'The x co-ordinate changes by 95 pixels
            y_coordinate = 85 * y_ratio                                                       'the y co-ordinate is stationary, this creates a straight line
            With label_array(x, 2)                                                  'define the label properties
                .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
                .Text = Chr(64 + x)                                                 'The ascii character for the capital letter relating to the number (1 = A, 2 = B)
                .Size = New System.Drawing.Size(70 * x_ratio, 30 * y_ratio)                             'The label size
                .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)  'The font and text size
            End With
            Me.Controls.Add(label_array(x, 2))                                      'Add to the controls
        Next x
        'numbers go down the left side of the board
        For y = 1 To 8
            x_coordinate = 250 * x_ratio                                                      'The x co-ordinate is stationary
            y_coordinate = 5 * y_ratio * (19 * y + 10)                                       'The y co-ordinate increments by 95, starting at 145 pixels in
            With label_array(y, 1)
                .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
                .Text = y
                .Size = New System.Drawing.Size(50 * x_ratio, 30 * y_ratio)
                .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
            End With
            Me.Controls.Add(label_array(y, 1))                                      'Add label to the controls
        Next y

        'create the labels that display the players score
        With lblP1Score
            .Text = p1name & " : 0"
            .Location = New System.Drawing.Point(70 * x_ratio, 150 * y_ratio)
            .Size = New System.Drawing.Size(130 * x_ratio, 25 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
        End With
        With lblP2Score
            .Text = p2name & " : 0"
            .Location = New System.Drawing.Point(70 * x_ratio, 400 * y_ratio)
            .Size = New System.Drawing.Size(130 * x_ratio, 25 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
        End With
        'Add to controls
        Me.Controls.Add(lblP1Score)
        Me.Controls.Add(lblP2Score)

        'Create the label that tells the players who's turn it is
        With lblTurn
            .Text = "Turn : "
            .Location = New System.Drawing.Point(70 * x_ratio, 105 * y_ratio)
            .Size = New System.Drawing.Size(140 * x_ratio, 25 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
        End With
        'Add to controls
        Me.Controls.Add(lblTurn)

        'Create the heading labels
        With lblP1Heading
            .Text = p1name
            .Size = New System.Drawing.Size(250 * x_ratio, 50 * y_ratio)
            .Location = New System.Drawing.Point(570 * x_ratio, 35 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(20 * x_ratio), FontStyle.Underline)
        End With
        With lblP2heading
            .Text = p2name
            .Size = New System.Drawing.Size(250 * x_ratio, 50 * y_ratio)
            .Location = New System.Drawing.Point(570 * x_ratio, 885 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(20 * x_ratio), FontStyle.Underline)
        End With
        'Add labels to controls
        Me.Controls.Add(lblP1Heading)
        Me.Controls.Add(lblP2heading)

        'Create the title of the scores listbox
        With lblScoresHeading
            .Location = New System.Drawing.Point(1075 * x_ratio, 75 * y_ratio)
            .Size = New System.Drawing.Size(170 * x_ratio, 30 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(16 * x_ratio), FontStyle.Underline)
            .Text = "High Scores"
        End With
        Me.Controls.Add(lblScoresHeading) 'add labels to controls
    End Sub

    Private Sub create_listboxes(x_ratio, y_ratio)
        'creates the list boxes that will be used in the program
        'Create the listbox that displays the scores
        With lstScores
            .Size = New System.Drawing.Size(170 * x_ratio, 500 * y_ratio)
            .Location = New System.Drawing.Point(1070 * x_ratio, 110 * y_ratio)
        End With
        'Add to controls
        Me.Controls.Add(lstScores) 'add listbox to controls

        'Create the list boxes that display the grave of each player
        With lstP1Grave
            .Size = New System.Drawing.Size(160 * x_ratio, 200 * y_ratio)
            .Location = New System.Drawing.Point(60 * x_ratio, 180 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(10 * x_ratio), FontStyle.Regular)
        End With
        With lstP2Grave
            .Size = New System.Drawing.Size(160 * x_ratio, 200 * y_ratio)
            .Location = New System.Drawing.Point(60 * x_ratio, 430 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(10 * x_ratio), FontStyle.Regular)
        End With
        'Add to controls
        Me.Controls.Add(lstP1Grave)
        Me.Controls.Add(lstP2Grave)
    End Sub

    Private Sub fill_control_arrays()
        'fills the picbox_array with the required picture boxes
        picbox_array(1, 1) = picbox1
        picbox_array(2, 1) = picbox2
        picbox_array(3, 1) = picbox3
        picbox_array(4, 1) = picbox4
        picbox_array(5, 1) = picbox5
        picbox_array(6, 1) = picbox6
        picbox_array(7, 1) = picbox7
        picbox_array(8, 1) = picbox8
        picbox_array(1, 2) = picbox9
        picbox_array(2, 2) = picbox10
        picbox_array(3, 2) = picbox11
        picbox_array(4, 2) = picbox12
        picbox_array(5, 2) = picbox13
        picbox_array(6, 2) = picbox14
        picbox_array(7, 2) = picbox15
        picbox_array(8, 2) = picbox16
        picbox_array(1, 3) = picbox17
        picbox_array(2, 3) = picbox18
        picbox_array(3, 3) = picbox19
        picbox_array(4, 3) = picbox20
        picbox_array(5, 3) = picbox21
        picbox_array(6, 3) = picbox22
        picbox_array(7, 3) = picbox23
        picbox_array(8, 3) = picbox24
        picbox_array(1, 4) = picbox25
        picbox_array(2, 4) = picbox26
        picbox_array(3, 4) = picbox27
        picbox_array(4, 4) = picbox28
        picbox_array(5, 4) = picbox29
        picbox_array(6, 4) = picbox30
        picbox_array(7, 4) = picbox31
        picbox_array(8, 4) = picbox32
        picbox_array(1, 5) = picbox33
        picbox_array(2, 5) = picbox34
        picbox_array(3, 5) = picbox35
        picbox_array(4, 5) = picbox36
        picbox_array(5, 5) = picbox37
        picbox_array(6, 5) = picbox38
        picbox_array(7, 5) = picbox39
        picbox_array(8, 5) = picbox40
        picbox_array(1, 6) = picbox41
        picbox_array(2, 6) = picbox42
        picbox_array(3, 6) = picbox43
        picbox_array(4, 6) = picbox44
        picbox_array(5, 6) = picbox45
        picbox_array(6, 6) = picbox46
        picbox_array(7, 6) = picbox47
        picbox_array(8, 6) = picbox48
        picbox_array(1, 7) = picbox49
        picbox_array(2, 7) = picbox50
        picbox_array(3, 7) = picbox51
        picbox_array(4, 7) = picbox52
        picbox_array(5, 7) = picbox53
        picbox_array(6, 7) = picbox54
        picbox_array(7, 7) = picbox55
        picbox_array(8, 7) = picbox56
        picbox_array(1, 8) = picbox57
        picbox_array(2, 8) = picbox58
        picbox_array(3, 8) = picbox59
        picbox_array(4, 8) = picbox60
        picbox_array(5, 8) = picbox61
        picbox_array(6, 8) = picbox62
        picbox_array(7, 8) = picbox63
        picbox_array(8, 8) = picbox64

        'Fill label array
        'lbl(n, 1) is the numbered labels, lbl(n, 2) is the lettered ones
        label_array(1, 1) = lbl1
        label_array(2, 1) = lbl2
        label_array(3, 1) = lbl3
        label_array(4, 1) = lbl4
        label_array(5, 1) = lbl5
        label_array(6, 1) = lbl6
        label_array(7, 1) = lbl7
        label_array(8, 1) = lbl8
        label_array(1, 2) = lblA
        label_array(2, 2) = lblB
        label_array(3, 2) = lblC
        label_array(4, 2) = lblD
        label_array(5, 2) = lblE
        label_array(6, 2) = lblF
        label_array(7, 2) = lblG
        label_array(8, 2) = lblH
    End Sub

    Private Sub fill_board()
        'Fills the board array in its starting position
        'Follows the rules of chess in its layout
        'places the pawns in their positions
        For x = 1 To 8
            board(x, 2) = "wp"
            board(x, 7) = "bp"
        Next
        'Place the other pieces in their correct positions
        board(1, 1) = "wr"
        board(2, 1) = "wh"
        board(3, 1) = "wb"
        board(4, 1) = "wq"
        board(5, 1) = "wk"
        board(6, 1) = "wb"
        board(7, 1) = "wh"
        board(8, 1) = "wr"
        board(1, 8) = "br"
        board(2, 8) = "bh"
        board(3, 8) = "bb"
        board(4, 8) = "bq"
        board(5, 8) = "bk"
        board(6, 8) = "bb"
        board(7, 8) = "bh"
        board(8, 8) = "br"

        'this fills the rest of the array with "" values
        'this is instead of Nothing values
        For x = 1 To 8
            For y = 3 To 6
                board(x, y) = ""
            Next y
        Next x
    End Sub

    Private Function get_background(x As Integer, y As Integer) As Color
        'These If statements display the checked background that chess boards have
        If x Mod 2 = 0 And y Mod 2 = 0 Then
            Return Color.White
        End If
        If x Mod 2 <> 0 And y Mod 2 = 0 Then
            Return Color.DarkGray
        End If
        If x Mod 2 = 0 And y Mod 2 <> 0 Then
            Return Color.DarkGray
        End If
        If x Mod 2 <> 0 And y Mod 2 <> 0 Then
            Return Color.White
        End If
    End Function
    Private Sub display_board()
        'This sub takes the board array and outputs it onto the screen
        Dim pic As String = ""          'stores the acronym for the picture
        For x = 1 To 8
            For y = 1 To 8
                picbox_array(x, y).BackColor = get_background(x, y)
                'display the pictures of the pieces on the board
                pic = board(x, y)
                If pic = "" Then        'if pic = "", then an error will occur because no picture found
                    picbox_array(x, y).Image = Nothing
                Else
                    picbox_array(x, y).Image = Image.FromFile(frmMain.bin & "\chess pieces\" & pic & ".png")
                End If
                picbox_array(x, y).BorderStyle = BorderStyle.None
            Next y
        Next x
    End Sub

    Private Sub restart_game()
        'This sub deals with restarting the game and resetting the board
        fill_board()            're fills the board array to the opening format
        display_board()         'outputs the re filled board onto the screen in its starting layout
        lblP1Heading.Text = p1name
        lblP2heading.Text = p2name
        lstP1Grave.Items.Clear()
        lstP2Grave.Items.Clear()
        game_ended = False      'begins the new game
        click_counter = 0
        turn = "p1"
        lblTurn.Text = "Turn : " & p1name
        p1score = 0
        p2score = 0
        lblP1Score.Text = p1name & " : " & p1score
        lblP2Score.Text = p2name & " : " & p2score
    End Sub

    Private Sub make_move()
        'This sub deals with the player making a move.
        'It controls the flow of logic
        Dim move(2, 2) As Integer                                           'A 2D array storing [[origin x, origin y], [destination x, destination y]]
        Dim attacking, piece As String                                      'Attacking stores the acronym of the piece (if any) that is being taken
        Dim valid As Boolean = False                                        'This is used to validate moves
        Dim castle As Boolean = False                                       'This is used to validate a castle move
        Dim side As String
        get_move(move, piece)                                               'This sub gets the raw data (origin, destination) and fills the move array
        check_move(piece, move, valid)                                      'This sub checks the move using global variables (board) as well as the move array
        castle = check_castle(piece, move)
        attacking = board(move(2, 1), move(2, 2))                           'This gets the piece that the player is taking
        For x = 1 To 8
            For y = 1 To 8
                picbox_array(x, y).BorderStyle = BorderStyle.None           'This resets the borderstyles of all the tiles
                picbox_array(x, y).BackColor = get_background(x, y)         'This resets the backgrounds of all of the tiles
            Next
        Next
        If Not valid And Not castle Then                                    'This accounts for invalid moves
            click_counter = 0                                               'Reset click_counter
            swap(turn)                                                      'Swapping the turn here means that when it is swapped later, it is back to it's original value
        End If
        If valid Then
            execute(move, piece, attacking)                                 'If the move is valid, execute it on the board
        End If
        If castle Then                                                      'This executes if the player has selected to castle
            If move(2, 1) = 3 Then                                          'This identifies the type of castle move (king or queen side)
                side = "queen"
            Else
                side = "king"
            End If
            update_castle(move, piece, side)                                'Update the board with the castle move
        End If
        If valid Or castle Then
            Select Case turn                'Toggle the turn lable
                Case "p2" : lblTurn.Text = "Turn : " & p1name
                Case "p1" : lblTurn.Text = "Turn : " & p2name
            End Select
        End If
    End Sub

    Private Sub save_game(game_title)
        'This sub must save games to a given file that is defined by the games title
        'It works by resetting the file to have nothing in it
        'And then re writing all the data into it in a specific order so that it can be read by the computer later
        Dim path As String = frmMain.bin & "\saved games\" & game_title & ".txt"    'Path is set here to avoid long, unweidly scentences
        System.IO.File.Open(path, System.IO.FileMode.Truncate).Close()                          'This code resets the contents of the file
        FileSystem.FileOpen(1, path, OpenMode.Append)                                           'Open the file in append mode
        FileSystem.WriteLine(1, p1name)                                                         'append P1name to file
        FileSystem.WriteLine(1, p1score)                                                        'append P1score to file
        FileSystem.WriteLine(1, p2name)                                                         'append P2name to file
        FileSystem.WriteLine(1, p2score)                                                        'append p2score to file
        FileSystem.WriteLine(1, turn)                                                           'append turn to file
        For y = 1 To 8
            For x = 1 To 8
                FileSystem.WriteLine(1, board(x, y))                                            'append the board array to the file in a specific order (across then down)
            Next
        Next
        FileSystem.FileClose(1)                                                                 'close the file
    End Sub

    Private Sub get_move(ByRef move, ByRef piece)
        'uses board, origin and destination as its global input
        'This sub must determine the players move given the tiles that were clicked
        'as well as the piece on that tile
        Dim length, temp As Integer
        length = origin.Length
        temp = Int(Mid(origin, 8))              'This is the number of the picbox clicked
        'e.g. if origin = picTile34, temp = 34
        'This gets the x, y co-ordinates of the origin tile
        move(1, 1) = ((temp - 1) Mod 8) + 1     'x co-ordinate origin
        move(1, 2) = Int((temp - 1) / 8) + 1    'y co-ordinate origin

        'this gets the destination co-ordinates
        length = destination.Length
        temp = Int(Mid(destination, 8))
        move(2, 1) = ((temp - 1) Mod 8) + 1
        move(2, 2) = Int((temp - 1) / 8) + 1
        piece = board(move(1, 1), move(1, 2))   'this is the piece being moved
        'move now contains the two co-ordinates of the move
        'e.g. if move contains [(3, 4), (4, 5)] it would represent a move origin point (3, 4) to point (4, 5). 
        'piece now contains the acronym of the piece moving
    End Sub

    Private Sub check_move(ByVal piece, ByVal move, ByRef valid)
        'This sub takes the piece and move array and assigns functions to check the move. It then sets valid to an appropriate boolean value
        'This does not do any of the calculations itself, it merely assigns functions to do that
        valid = True                                                'Assume move is true
        If (turn = "p1" And Mid(piece, 1, 1) <> "w") Or (turn = "p2" And Mid(piece, 1, 1) <> "b") Then
            valid = False                                           'checks that the player is moving the correct piece
        End If
        If move(1, 1) = move(2, 1) And move(1, 2) = move(2, 2) Then
            valid = False                                           'This is the case where the player hasn't moved
        End If
        If valid = True Then                                        'This assigns an appropriate function
            Select Case Mid(piece, 2, 1)                            'This gets the right character of the piece string (the piece being moved)
                Case "p" : valid = check_pawn(piece, move)          'Case pawn is moved
                Case "r" : valid = check_rook(piece, move)          'Case rook is moved
                Case "h" : valid = check_horse(piece, move)         'Case knight is moved (h is used for horse to avoid conflict with the king which also uses k)
                Case "b" : valid = check_bishop(piece, move)        'Case bishop is moved
                Case "q" : valid = check_queen(piece, move)         'Case queen is moved
                Case "k" : valid = check_king(piece, move)          'Case king is moved
            End Select
        End If
    End Sub

    Private Sub get_all_moves(origin)
        'This scans the whole board and gets all possible moves that the player can make with a particular clicked piece
        'It uses the check_piece sub to get a boolean value for the validity of every possible move
        'The local move data structure is used here to simulate the data structure of the game
        'All possible destination squares are highlighted blue
        'This has to be reset once the player has made their second click
        Dim valid As Boolean = False                'Used to check each destination
        Dim castle As Boolean = False
        Dim move(2, 2) As Integer
        Dim length, temp, x_coordinate, y_coordinate As Integer
        Dim piece As String
        length = origin.Length
        temp = Int(Mid(origin, 8))                  'This is the number of the picbox clicked
        'e.g. if origin = picTile34, temp = 34
        'This gets the x, y co-ordinates of the origin tile
        x_coordinate = ((temp - 1) Mod 8) + 1       'x co-ordinate origin
        y_coordinate = Int((temp - 1) / 8) + 1      'y co-ordinate origin
        piece = board(x_coordinate, y_coordinate)   'Get the piece that has been clicked
        move(1, 1) = x_coordinate                   'set the x, y co_ordinates of the origin
        move(1, 2) = y_coordinate
        For y = 1 To 8                              'Cycle through the array and check each tile
            For x = 1 To 8
                move(2, 1) = x                      'set x and y
                move(2, 2) = y
                check_move(piece, move, valid)      'Check the tile
                If valid Then                       'If valid, change the background and borderstyle
                    'This turns the boxes light blue to distinguish them from normal tiles
                    picbox_array(move(2, 1), move(2, 2)).BackColor = Color.LightBlue
                    picbox_array(move(2, 1), move(2, 2)).BorderStyle = BorderStyle.FixedSingle
                End If
                valid = False                       'reset valid
                castle = check_castle(piece, move)
                If castle Then                      'This deals with all valid castle moves (king and queen side)
                    'This turns the boxes salmon to distinguish them from normal moves
                    picbox_array(move(2, 1), move(2, 2)).BackColor = Color.LightSalmon
                    picbox_array(move(2, 1), move(2, 2)).BorderStyle = BorderStyle.FixedSingle
                End If
            Next x
        Next y
    End Sub

    Private Sub execute(move, piece, attacking)
        If attacking <> "" Then             'This will be "" if not attacking
            update_score(attacking)         'If activated, the score will be updated according to the rules of chess
            update_grave(attacking)
        End If
        If Mid(attacking, 2, 1) = "k" Then  'This checks if the king piece has been taken
            game_ended = True               'if the king has been taken, the game is over
            If Mid(piece, 1, 1) = "w" Then  'this determines the winner of the game
                winner = "p2"
            Else
                winner = "p1"
            End If
            scoring()                       'This deals the scoring when the winner is determined
        Else
            update_board(move, piece)       'This updates the board
        End If
    End Sub

    Private Sub update_score(attacking)
        'This sub must take the attacking variable and update the appropriate score with appropriate values
        Const p = 1                             'the scores of taking each piece
        Const r = 5
        Const h = 3
        Const b = 3
        Const q = 9
        Const k = 0
        If turn = "p1" Then                     'will increment p1's score
            Select Case Mid(attacking, 2, 1)
                Case "p" : p1score = p1score + p
                Case "r" : p1score = p1score + r
                Case "h" : p1score = p1score + h
                Case "b" : p1score = p1score + b
                Case "q" : p1score = p1score + q
                Case "k" : p1score = p1score + k
            End Select
        Else
            Select Case Mid(attacking, 2, 1)    'update p2's score
                Case "p" : p2score = p2score + p
                Case "r" : p2score = p2score + r
                Case "h" : p2score = p2score + h
                Case "b" : p2score = p2score + b
                Case "q" : p2score = p2score + q
                Case "k" : p2score = p2score + k
            End Select
        End If
        lblP1Score.Text = "Player 1 : " & p1score 'update the labels that display the scores
        lblP2Score.Text = "Player 2 : " & p2score
    End Sub

    Private Sub update_castle(move, piece, side)
        'This sub updates the board when a valid castle move has been submitted
        'There are two different types of castle, queen side and king side
        'That is why there is a different protocol for the king and queen side castle

        'The king side castle moves the king from its starting position two places right and moves the rook two places left
        If side = "king" Then
            board(move(1, 1), move(1, 2)) = ""                      'update the board with the kings move (two places right)
            board(move(2, 1), move(2, 2)) = piece
            picbox_array(move(1, 1), move(1, 2)).Image = Nothing    'Update the pictures with the kings move
            picbox_array(move(2, 1), move(2, 2)).Image = Image.FromFile(frmMain.bin & "\chess pieces\" & piece & ".png")
            board(8, move(2, 2)) = ""                               'Update the board with the rooks move (two places left)
            board(6, move(2, 2)) = Mid(piece, 1, 1) & "r"
            picbox_array(8, move(2, 2)).Image = Nothing             'Update the pictures with the rooks move
            picbox_array(6, move(2, 2)).Image = Image.FromFile(frmMain.bin & "\chess pieces\" & Mid(piece, 1, 1) & "r.png")
        Else    'The queen side castle moves the king two places left and the rook three places right
            board(move(1, 1), move(1, 2)) = ""                      'Update the board with the kings move (two places left)
            board(move(2, 1), move(2, 2)) = piece
            picbox_array(move(1, 1), move(1, 2)).Image = Nothing    'Update the pictures with the kings move
            picbox_array(move(2, 1), move(2, 2)).Image = Image.FromFile(frmMain.bin & "\chess pieces\" & piece & ".png")
            board(1, move(2, 2)) = ""                               'Update the board with the rooks move (Three places right)
            board(4, move(2, 2)) = Mid(piece, 1, 1) & "r"
            picbox_array(1, move(2, 2)).Image = Nothing             'Update the pictures with the rooks move
            picbox_array(4, move(2, 2)).Image = Image.FromFile(frmMain.bin & "\chess pieces\" & Mid(piece, 1, 1) & "r.png")
        End If
        Select Case turn
            Case "p1" : lblTurn.Text = "Turn : " & p1name
            Case "p2" : lblTurn.Text = "Turn : " & p2name
        End Select
    End Sub

    Private Function check_castle(piece As String, move(,) As Integer) As Boolean
        'This function checks a given move and checks whether it is a valid castle move
        'A castle move is where the king and the rook both move to set locations in one turn
        'There are two types of castle, king and queen side
        'There are two criteria to castle. The king and rook must be in their starting positions, there must be no pieces in between the king and queen.
        'It is a crucial move in chess
        Dim a, b, c, d As Boolean   'These are used to outsource some of the boolean calculations involved in testing a castle move
        'This checks that the player is moving his/her piece
        If (turn = "p1" And board(move(1, 1), move(1, 2)) = "bk") Or (turn = "p2" And board(move(1, 1), move(1, 2)) = "wk") Then
            Return False
        End If
        If Mid(piece, 1, 1) = "w" Then              'This handles a white castle
            'The variables a and b deal with the king side castle
            'a checks that the move represents a valid castle move
            a = move(1, 1) = 5 And move(1, 2) = 1 And move(2, 2) = 1 And move(2, 1) = 7
            'b checks that the pieces are in their required positions and that the tiles in between are vacant
            b = board(7, 1) = "" And board(8, 1) = "wr" And board(6, 1) = "" And board(5, 1) = "wk"
            'the variables c and d deal with the queen side castle
            'c checks that the move represents a valid castle move
            c = move(1, 1) = 5 And move(1, 2) = 1 And move(2, 2) = 1 And move(2, 1) = 3
            'd checks that the pieces are in their required positions and that the tiles in between are vacant
            d = board(2, 1) = "" And board(3, 1) = "" And board(4, 1) = "" And board(1, 1) = "wr" And board(5, 1) = "wk"
            If (a And b) Or (c And d) Then      'The function returns true for both valid king and queen side castle moves
                Return True
            End If
        Else                                        'This handles a black castle
            'The variables a and b deal with the king side castle
            'a checks that the move represents a valid castle move
            a = move(1, 1) = 5 And move(1, 2) = 8 And move(2, 2) = 8 And move(2, 1) = 7
            'b checks that the pieces are in their required positions and that the tiles in between are vacant
            b = board(7, 8) = "" And board(8, 8) = "br" And board(6, 8) = "" And board(5, 8) = "bk"
            'the variables c and d deal with the queen side castle
            'c checks that the move represents a valid castle move
            c = move(1, 1) = 5 And move(1, 2) = 8 And move(2, 2) = 8 And move(2, 1) = 3
            'd checks that the pieces are in their required positions and that the tiles in between are vacant
            d = board(2, 8) = "" And board(3, 8) = "" And board(4, 8) = "" And board(1, 8) = "br" And board(5, 8) = "bk"
            If (a And b) Or (c And d) Then      'The function returns true for both valid king and queen side castle moves
                Return True
            End If
        End If
        Return False                            'return false for all invalid castle moves
    End Function

    Private Sub update_grave(attacking)
        'This sub updates the grave list that lists all hte pieces that each player has lost and their value
        Const p = 1                             'the scores of taking each piece
        Const r = 5
        Const h = 3
        Const b = 3
        Const q = 9
        Const k = 0
        Select Case attacking                   'This checks the piece that is being taken and uses its colour to place it in the correct list
            Case "wp" : lstP1Grave.Items.Add("White Pawn " & p & "pt")
            Case "wr" : lstP1Grave.Items.Add("White Rook " & r & "pts")
            Case "wh" : lstP1Grave.Items.Add("White Horse " & h & "pts")
            Case "wb" : lstP1Grave.Items.Add("White Bishop " & b & "pts")
            Case "wq" : lstP1Grave.Items.Add("White Queen " & q & "pts")
            Case "wk" : lstP1Grave.Items.Add("White King " & k & "pts")
            Case "bp" : lstP2Grave.Items.Add("Black Pawn " & p & "pt")
            Case "br" : lstP2Grave.Items.Add("Black Rook " & r & "pts")
            Case "bh" : lstP2Grave.Items.Add("Black Horse " & h & "pts")
            Case "bb" : lstP2Grave.Items.Add("Black Bishop " & b & "pts")
            Case "bq" : lstP2Grave.Items.Add("Black Queen " & q & "pts")
            Case "bk" : lstP2Grave.Items.Add("Black King " & k & "pts")
        End Select
    End Sub

    Private Sub update_board(move, piece)
        change_pawn(piece, move)
        'This sub will change the board_array and the board
        picbox_array(move(1, 1), move(1, 2)).Image = Nothing    'Update the board pictures
        picbox_array(move(2, 1), move(2, 2)).Image = Image.FromFile(frmMain.bin & "\chess pieces\" & piece & ".png")
        board(move(1, 1), move(1, 2)) = ""                      'Update the board array
        board(move(2, 1), move(2, 2)) = piece
    End Sub

    Private Sub change_pawn(ByRef piece, move)
        'This sub deals with a pawn that has reached the other players end
        'The rules of chess say that you have to exchange the pawn for a queen of the same colour
        If (piece = "bp" And move(2, 2) = 1) Or (piece = "wp" And move(2, 2) = 8) Then
            'This checks if the piece being moved is a pawn and if it is in the required position
            frmPawnSelection.ShowDialog()                                   'This is a customised dialogue box which suspends code execution until it is closed
            'It allows the player to select the piece that they wish to promote
            Select Case piece       'This swaps the white pawn for a white piece of the players choosing
                Case "bp" : piece = "b" & frmPawnSelection.promotion_selection
                Case "wp" : piece = "w" & frmPawnSelection.promotion_selection
            End Select
        End If
    End Sub

    Private Sub scoring()
        'This sub Handles the score submission process and outputs the scores onto the board
        Dim scores(1) As people
        write_file(p1name, p1score)         'write the two score, name sets to the scores.txt file
        write_file(p2name, p2score)
        get_high_scores(scores)             'read the whole file into two arrays (names and scores)
        sort_scores(scores)                 'temp is the an array of arrays that gets the output from sort_scores\
        disp_high_scores(scores)
    End Sub

    Private Sub get_high_scores(ByRef scores() As people)
        'This sub reads the scores sequential file and places the values into two arrays storing the names and corresponding scores
        Dim temp_name, temp_score As String                                 'The temporary values that are then validated before being placed in their arrays
        Dim counter As Integer
        counter = 1
        FileOpen(1, frmMain.bin & "\scores.txt", OpenMode.Input) 'Open the file
        While Not EOF(1)                                                    'Check for the end of file before attempting to read from it
            FileSystem.Input(1, temp_name)                                  'read the name value into temp_name (the file is constructed in {name : score} pairs
            FileSystem.Input(1, temp_score)                                 'read the score value into temp_score
            If validate_integer(temp_score) Then                            'This validates the integer for good measure
                scores(counter).score = CInt(temp_score)                    'this converts the now validated integer and puts it into the array
            End If
            scores(counter).name = temp_name
            counter = counter + 1
            ReDim Preserve scores(counter)                                  'increase the length of the arrays by one
        End While
        FileSystem.FileClose(1)
        ReDim Preserve scores(counter - 1)                                  'decrease the length of the arrays because the last element is always empty
    End Sub

    Private Sub disp_high_scores(scores() As people)
        'This sub displays the ordered {name : score} pairs into the listbox
        lstScores.Items.Clear()                     'Clear the scores from any previous games
        For i = 1 To scores.Length - 1              'Cycle through the parallel scores & names array and output the pairs
            lstScores.Items.Add(scores(i).name & " : " & scores(i).score)
        Next
    End Sub

    Private Sub write_file(name, score)
        'This sub appends the name and score to the file
        FileSystem.FileOpen(1, frmMain.bin & "\scores.txt", OpenMode.Append)    'open file in append mode
        FileSystem.WriteLine(1, name)                                                       'write the name to the file
        FileSystem.WriteLine(1, score)                                                      'Write the score to the file
        FileSystem.FileClose(1)                                                             'Close the file
    End Sub

    Private Function sort_scores(scores() As people) As Array
        'Sorts in decending order using a selection sort
        'much faster than bubble sort for large data sets
        'This takes an array of records (people) and sorts by score
        Dim end_unsorted, temp_int, min, i, posmax As Integer 'temp_int and temp_str are used to store temporary values of different data types
        Dim temp_str As String
        end_unsorted = scores.Length - 1    'The last element in the array scores
        temp_int = 0
        While end_unsorted > 1
            i = 1
            min = scores(i).score           'set the min to the first array element
            posmax = 1                      'set the index (posmax) to 1
            While i < end_unsorted
                i = i + 1
                If scores(i).score < min Then 'compare the ith element of scores with the current min value
                    min = scores(i).score   'if the ith element is less, set min to the ith element
                    posmax = i
                End If
            End While
            'swap scores(posmax) with scores(end_unsorted)
            temp_int = scores(posmax).score
            scores(posmax).score = scores(end_unsorted).score
            scores(end_unsorted).score = temp_int
            'swap names(posmax) with names(end_unsorted) to keep arrays parallel
            temp_str = scores(posmax).name
            scores(posmax).name = scores(end_unsorted).name
            scores(end_unsorted).name = temp_str
            end_unsorted = end_unsorted - 1
        End While
        Return scores
    End Function

    Private Function validate_integer(number As String) As Boolean
        'This Takes a candidate and returns a boolean indicating if each and every character in the candidate is an integer
        If number = "" Or IsNothing(number) Then 'this checks for null string values
            Return False
        End If
        For Each i In number                    'This increments through each character in the string
            If Asc(i) < 48 Or Asc(i) > 57 Then  'These are the ascii characters that represent the digits
                Return False
            End If
        Next
        Return True                             'return true in all other cases
    End Function

    Private Function check_pawn(piece As String, move(,) As Integer) As Boolean
        'This sub checks a pawn move given the move array
        Dim taken_piece As String
        Dim change_x, change_y As Integer
        Dim attacking, starting As Boolean
        Dim a, b, c, d As Boolean
        taken_piece = board(move(2, 1), move(2, 2))
        'will check that the piece to be taken is not being taken by the same colour (e.g. wq cannot take another wp) 
        If Mid(piece, 1, 1) = Mid(taken_piece, 1, 1) Then
            Return False
        End If
        change_x = System.Math.Abs(move(1, 1) - move(2, 1))     'the change in the x component of the move
        'abs used because direction is irrelevant
        change_y = move(2, 2) - move(1, 2)                      'a move down the screen is positive, a move up is negative
        attacking = False
        starting = False
        If board(move(2, 1), move(2, 2)) <> "" Then             'checks if the piece is attacking (if so, special rules apply)
            attacking = True
        End If
        'checks if the piece is in the starting position
        If (move(1, 2) = 2 And Mid(piece, 1, 1) = "w") Or (move(1, 2) = 7 And Mid(piece, 1, 1) = "b") Then
            starting = True
        End If
        'case where pawn is attacking
        a = attacking And Not starting And ((change_y = 1 And Mid(piece, 1, 1) = "w") Or (change_y = -1 And Mid(piece, 1, 1) = "b")) And change_x = 1
        'case where the pawn is moving 1 or 2 spaces forward from starting position, not attacking
        b = Not attacking And starting And change_x = 0 And (((change_y = 1 Or change_y = 2) And Mid(piece, 1, 1) = "w") Or ((change_y = -1 Or change_y = -2) And Mid(piece, 1, 1) = "b"))
        'case where the pawn is moving 1 space foward, not attacking, not starting
        c = Not attacking And Not starting And change_x = 0 And ((change_y = 1 And Mid(piece, 1, 1) = "w") Or (change_y = -1 And Mid(piece, 1, 1) = "b"))
        'case where the pawn is attacking on the first move
        d = attacking And starting And change_x = 1 And ((change_y = 1 And Mid(piece, 1, 1) = "w") Or (change_y = -1 And Mid(piece, 1, 1) = "b"))
        If a Or b Or c Or d Then
            Return True
        End If
        Return False
    End Function

    Private Function check_rook(piece As String, move(,) As Integer) As Boolean
        'This sub checks a rooks move
        Dim taken_piece As String
        Dim change_x, change_y As Integer
        Dim x, y As Integer
        'rooks can move vertically and horizontally for any distance
        taken_piece = board(move(2, 1), move(2, 2))
        If Mid(piece, 1, 1) = Mid(taken_piece, 1, 1) Then   'will check that the piece that is to be taken by the rook is not of the same colour as it
            Return False
        End If
        change_x = move(2, 1) - move(1, 1)                  'a negative value indicates moving Left
        change_y = move(2, 2) - move(1, 2)                  'a negative value indicates moving backwards
        If change_x <> 0 And change_y <> 0 Then             'checks that the move is not diagonal
            Return False
        End If
        'checks movement in the x co-ordinate
        If change_x > 0 Then                                'checks wether there are any pieces in between the rook and it's destination
            x = move(1, 1) + 1                              'the +1 is so that the FOR...NEXT loop does not check the tile its self
            For i = x To x + change_x - 2                   'this will cycle through the board until it reaches the specified end point
                If board(i, move(1, 2)) <> "" Then
                    Return False
                End If
            Next i
        End If
        If change_x < 0 Then                                'case: moving backwards
            x = move(1, 1) - 1                              'the -1 is so that the FOR...NEXT loop does not check the tile its self
            For i = x To x + change_x + 2 Step -1           'cycles from Right to Left through the board
                If board(i, move(1, 2)) <> "" Then
                    Return False
                End If
            Next i
        End If
        'checks movement in the y co-ordinate
        If change_y > 0 Then
            y = move(1, 2) + 1
            For i = y To y + change_y - 2
                If board(move(1, 1), i) <> "" Then
                    Return False
                End If
            Next i
        End If
        If change_y < 0 Then                                'if the piece is moving down the board
            y = move(1, 2) - 1
            For i = y To y + change_y + 2 Step -1
                If board(move(1, 1), i) <> "" Then
                    Return False
                End If
            Next i
        End If
        Return True                                         'returns true in all other cases
    End Function

    Private Function check_horse(piece As String, move(,) As Integer) As Boolean
        'This sub checks a horses move
        'for horses move see documentation
        Dim taken_piece As String
        Dim change_x, change_y As Integer
        Dim a, b As Boolean
        taken_piece = board(move(2, 1), move(2, 2))
        If Mid(piece, 1, 1) = Mid(taken_piece, 1, 1) Then   'will check that the piece that is to be taken by the rook is not of the same colour as it
            Return False
        End If
        change_x = System.Math.Abs(move(2, 1) - move(1, 1)) 'abs because +/- x becomes x
        change_y = System.Math.Abs(move(2, 2) - move(1, 2))
        a = change_x = 1 And change_y = 2                   'the horse has two different abs(x), abs(y) combinations that are acceptable
        b = change_x = 2 And change_y = 1
        If a Or b Then                                      'both case a and b are valid moves
            Return True
        End If
        Return False
    End Function

    Private Function check_bishop(piece As String, move(,) As Integer) As Boolean
        'bishops can move diagonally for any distance
        'for horses move see documentation
        Dim taken_piece As String
        Dim change_x, change_y As Integer
        Dim temp As String
        Dim x, y As Integer
        taken_piece = board(move(2, 1), move(2, 2))
        If Mid(piece, 1, 1) = Mid(taken_piece, 1, 1) Then               'will check that the piece that is to be taken by the rook is not of the same colour as it
            Return False
        End If
        change_x = move(2, 1) - move(1, 1)                              'the direction of the movment in the x and y axis is important so no abs is used
        change_y = move(2, 2) - move(1, 2)
        If System.Math.Abs(change_x) <> System.Math.Abs(change_y) Then  'this checks that the move is diagonal
            Return False
        End If
        temp = get_increment(change_x, change_y)                        'used to indicate which direction to increment the x and y values
        x = move(1, 1)
        y = move(1, 2)
        'will cycle through the tiles in between the piece and its destination
        While x < move(2, 1) - 1 Or x > move(2, 1) + 1                  'uses the fact that abs(x) = abs(y)
            If Mid(temp, 1, 1) = "+" Then
                x = x + 1
            Else
                x = x - 1
            End If
            If Mid(temp, 2, 1) = "+" Then
                y = y + 1
            Else
                y = y - 1
            End If
            If board(x, y) <> "" Then                                   'checks that the piece is not "jumping" any other pieces
                Return False
            End If
        End While
        Return True                                                     'This returns true for all other paths
    End Function

    Private Function get_increment(change_x As Integer, change_y As Integer) As String
        'used with the check_bishop function
        'this checks wether to increment or decrement x and y and Returns a string
        If change_x > 0 Then            'if change_x > 0, then the piece is moving diagonally to the Right
            If change_y > 0 Then        'if change_y > 0 then the piece is moving diagonally forward
                Return "++"
            Else
                Return "+-"
            End If
        Else
            If change_y > 0 Then
                Return "-+"
            Else
                Return "--"
            End If
        End If
    End Function

    Private Function check_queen(piece As String, move(,) As Integer) As Boolean
        'queen can move vertically, horizontally or diagonally for any distance
        'since the queen move is made up of rook and bishop moves, this uses the other functions
        Dim taken_piece As String
        Dim change_x, change_y As Integer
        taken_piece = board(move(2, 1), move(2, 2))
        If Mid(piece, 1, 1) = Mid(taken_piece, 1, 1) Then                       'will check that the piece that is to be taken by the rook is not of the same colour as it
            Return False
        End If
        change_x = move(2, 1) - move(1, 1)
        change_y = move(2, 2) - move(1, 2)
        If System.Math.Abs(change_x) = System.Math.Abs(change_y) Then           'checks a diagonal move
            Return check_bishop(piece, move)
        End If
        If (change_x = 0 And change_y <> 0) Or (change_x <> 0 And change_y = 0) Then
            Return check_rook(piece, move)
        End If
        Return False
    End Function

    Private Function check_king(piece As String, move(,) As Integer) As Boolean
        'king can move vertically, horizontally or diagonally 1 space
        'similarly to a queen, a king move is made up of rook and bishop moves
        Dim taken_piece As String
        Dim change_x, change_y As Integer
        taken_piece = board(move(2, 1), move(2, 2))
        If Mid(piece, 1, 1) = Mid(taken_piece, 1, 1) Then                                   'will check that the piece that is to be taken by the rook is not of the same colour as it
            Return False
        End If
        change_x = move(2, 1) - move(1, 1)
        change_y = move(2, 2) - move(1, 2)
        If (change_x >= -1 And change_x <= 1) And (change_y >= -1 And change_y <= 1) Then   'This checks that the move is only one square
            Return True
        End If
        Return False
    End Function

End Class
