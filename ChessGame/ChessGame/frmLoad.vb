Public Class frmLoad
#Region "public variables"
    'These variables are public throughout the program
    Public game_loaded As Boolean
    Public p1grave(0), p2grave(0) As String
#End Region

#Region "dynamic variables"
    'These objects are used in the creation of the form
    Dim games As New ArrayList()
    Dim lstGames As New ListBox
    Dim lblHeading As New Label
    Dim btnPlay, btnCancel As New Button
#End Region

    Private Sub frmLoad_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        'This creates the board and gets a list of the games that are available
        game_loaded = False
        create_board()
        get_games()
    End Sub

    Private Sub Cancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This closes the form
        frmMain.Show()  'This quits the form
        Me.Close()
    End Sub

    Private Sub create_board()
        'This creates the board which is a simple list box, label and two buttons
        Dim x_ratio As Single = Screen.PrimaryScreen.Bounds.Width               'Get the width of the screen in pixels
        Dim y_ratio As Single = Screen.PrimaryScreen.Bounds.Height              'Get the height of the screen in pixels
        Me.Size = New System.Drawing.Size(x_ratio * 0.15, y_ratio * 0.4)        'set the size of the form in relation to the screen
        Me.Location = New System.Drawing.Point(0.42 * x_ratio, 0.25 * y_ratio)  'set the location of the form in relation to the screen
        x_ratio = x_ratio / 1920                                                'set the x and y ratios in comparison to a 1080p screen
        y_ratio = y_ratio / 1080
        Me.Text = "Load Game"                                                   'Set the text of the form
        With lstGames                                                           'Describe the list box
            .Size = New System.Drawing.Size(140 * x_ratio, 250 * y_ratio)
            .Location = New System.Drawing.Point(50 * x_ratio, 40 * y_ratio)
        End With
        'Add to controls
        Me.Controls.Add(lstGames)

        With lblHeading                                                         'Describe the heading label
            .Size = New System.Drawing.Size(200 * x_ratio, 30 * y_ratio)
            .Location = New System.Drawing.Point(35 * x_ratio, 10 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(13 * x_ratio), FontStyle.Underline)
            .Text = "Availabel Games"
        End With
        'Add to controls
        Me.Controls.Add(lblHeading)

        With btnCancel                                                          'Describe the cancel button
            .Location = New System.Drawing.Point(10 * x_ratio, 320 * y_ratio)
            .Size = New System.Drawing.Size(115 * x_ratio, 35 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
            .Text = "Cancel"
        End With
        'Add to controls
        Me.Controls.Add(btnCancel)
        'Add Click Event
        AddHandler btnCancel.Click, AddressOf Cancel_Click

        With btnPlay                                                            'Describe the Delete button
            .Location = New System.Drawing.Point(130 * x_ratio, 320 * y_ratio)
            .Size = New System.Drawing.Size(115 * x_ratio, 35 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
            .Text = "Play"
        End With
        'Add to controls
        Me.Controls.Add(btnPlay)
        'Add Click Event
        AddHandler btnPlay.Click, AddressOf Play_Click
    End Sub

    Private Sub Play_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This activates when the delete button is clicked
        'It gets the selected item in the list box and deletes the appropriate file
        Dim index As Integer
        Dim path As String = frmMain.bin & "\saved games\"      'This is the path that all the game files are located in
        index = lstGames.SelectedIndex                                      'This is the index of the selected item in the list box
        If index <> -1 Then                                                 'index is -1 if nothing is selected
            frmNewGame.game_title = games(index)
            game_loaded = True                                              'Boolean indicating that the game has been loaded instead of a new game
            get_saved_game(frmNewGame.game_title)                           'Gets the data from the sequential file
            get_grave()                                                     'Searches the board array and finds any missing pieces and places them in the grave
            frmGame.Show()
            Me.Close()
        End If
    End Sub

    Private Sub get_saved_game(file)
        'This sub reads a sequential file and puts the information into various variables and arrays
        'The files are all written in a certain way to allow for this to happen
        'p1name, p1score, p2name, p2score, turn, each element in the board from left to right and top to bottom
        FileSystem.FileOpen(1, frmMain.bin & "\saved games\" & file & ".txt", OpenMode.Input)
        FileSystem.Input(1, frmGame.p1name)         'read p1name
        FileSystem.Input(1, frmGame.p1score)        'read p1score
        FileSystem.Input(1, frmGame.p2name)         'read p2name
        FileSystem.Input(1, frmGame.p2score)        'read p2score
        FileSystem.Input(1, frmGame.turn)           'read turn
        For y = 1 To 8                              'is written left to right and top to bottom so y is on the outside
            For x = 1 To 8
                FileSystem.Input(1, frmGame.board(x, y)) 'fill the board array according to the file
            Next x
        Next y
        FileSystem.FileClose(1)                     'close the file
    End Sub

    Private Sub get_grave()
        'This sub must search the board array and find any missing pieces, these pieces will represent pieces which have been taken by the other players
        'This requires a counting algorithm which counts the number of objects in an array
        Dim p1counter, p2counter As Integer
        Dim counter As Integer
        Dim pieces(2) As String                                 'This array holds all the pieces which are in pairs (knights, rooks and bishops)
        pieces(0) = "r"
        pieces(1) = "h"
        pieces(2) = "b"
        p1counter = 0                                           'The p1grave() and p2grave() are both zero indexed
        p2counter = 0
        counter = count_objects(frmGame.board, "wp")            'count the number of white pawns on the board (there are supposed to be 8)
        For i = 1 To 8 - counter                                'This runs if less than 8 pawns are found and it runs for the number of pawns less than 8 that are found
            ReDim p1grave(p1counter)
            p1grave(p1counter) = "wp"
            p1counter = p1counter + 1
        Next
        counter = count_objects(frmGame.board, "bp")            'This does the same process for black pawns
        For i = 1 To 8 - counter
            ReDim p2grave(p2counter)
            p2grave(p2counter) = "bp"
            p2counter = p2counter + 1
        Next
        For Each i In pieces                                    'This takes care of the main pieces (knights, rooks and bishops)
            counter = count_objects(frmGame.board, "w" & i)     'This counts the white pieces
            For r = 1 To 2 - counter                            'appends piece to grave array if not found
                ReDim p1grave(p1counter)
                p1grave(p1counter) = "w" & i
                p1counter = p1counter + 1
            Next
            counter = count_objects(frmGame.board, "b" & i)     'This counts the black pieces
            For r = 1 To 2 - counter
                ReDim p2grave(p2counter)
                p2grave(p2counter) = "b" & i
                p2counter = p2counter + 1
            Next
        Next
        counter = count_objects(frmGame.board, "wq")            'This deals with the white queen (of which there is only one)
        If counter <> 1 Then
            ReDim Preserve p1grave(p1counter)
            p1grave(p1counter) = "wq"
        End If
        counter = count_objects(frmGame.board, "bq")            'This deals with the black queen (of which there is only one)
        If counter <> 1 Then
            ReDim Preserve p2grave(p2counter)
            p2grave(p2counter) = "bq"
        End If
        'The graves should now contain the disparity between the starting board and the saved board.
        'this means that the score of all the pieces in p1grave should be equal to p2score and similarly for p2grave and p1score
    End Sub

    Private Sub get_games()
        'This sub looks in the saved games folder and makes a list of all the items there
        'It adds these items to the list box
        Dim path As String = frmMain.bin & "\saved games\"          'This is the folder that holds the score files
        Dim name As String
        Dim start As Integer                                                    'This is used to get just the text file name by cutting out the file path
        start = path.Length + 1
        For Each foundfile As String In My.Computer.FileSystem.GetFiles(path)   'This iterates through the files in the folder
            name = Mid(foundfile, start, foundfile.Length - start - 3)          'This selects just the Filename (without ".txt" on the end)
            games.Add(name)                                                     'Arraylist's have a .Add method which is useful
            lstGames.Items.Add(name)                                            'Add name to the list box
        Next foundfile
    End Sub

    Private Function count_objects(list(,) As String, search_item As String) As Integer
        'This sub takes a two dimensional array and a search_item and counts the number of occurrences of that item in the array
        'It returns an integer representing this number
        'It uses a linear search mechanism to search through the 2d array
        Dim counter As Integer                      'This is used to count the number of occurrences
        counter = 0
        For x = 0 To list.GetUpperBound(0)          'This is the upper bound of the first dimension (vb is zero indexed)
            For y = 0 To list.GetUpperBound(1)      'This is the upper bound of the second dimension
                If list(x, y) = search_item Then    'compare the element with the search item
                    counter = counter + 1           'If they are the same, increment counter
                End If
            Next y
        Next x
        Return counter
    End Function

End Class