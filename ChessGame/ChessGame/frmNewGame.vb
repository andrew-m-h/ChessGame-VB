Imports System.IO
Public Class frmNewGame
    Public game_title As String

#Region "dynamic variables"
    Dim txtP1name, txtP2name, txtGameName As New TextBox
    Dim lblP1name, lblP2name, lblGameName As New Label
    Dim txt_array(3) As TextBox
    Dim lbl_array(3) As Label
    Dim btnCreate, btnCancel, btnHelp As New Button
#End Region

    Private Sub frmNewGame_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        'This creates the form on start-up
        fill_control_arrays()   'Fill the control array
        create_board()          'Output the controls onto the form
    End Sub

    Private Sub Help_Click(sender As Object, e As System.EventArgs)
        MsgBox("The Game Name must contain only Capital letters, lower case letters or numbers and be at least one character in length. No spaces or special characters." & vbNewLine & _
               "The Initials of each player must be capital letters only, no spaces, numbers or lower case letter. They must be at least one character in length and less than six.")
    End Sub

    Private Sub Cancel_Click(sender As Object, e As System.EventArgs)
        'Return to the main screen
        frmMain.Show()
        Me.Close()
    End Sub

    Private Sub Create_Click(sender As Object, e As System.EventArgs)
        'This activates when the create new game button is clicked
        'It gets the values of the textboxes, validates them and creates a new save file
        Dim valid As Boolean = True                 'Assume true
        Dim error1, error2, error3 As String        'These store text for the three errors that can result from incorrect data entry
        get_details(valid, error1, error2, error3)  'These are all ByRef so they will be changed by the get_details sub
        If valid Then                               'Valid means that all three data values are correct
            save_to_file()                          'Create a new save file with the title of the game
            frmGame.Show()                          'Show the game form
            Me.Close()                              'Close this form
        Else
            MsgBox(error3 & vbNewLine & error1 & vbNewLine & error2) 'Return any of the errors that occurred with data entry
        End If
    End Sub

    Private Sub get_details(ByRef valid, ByRef error1, ByRef error2, ByRef error3)
        'This sub must get the values from the three text boxes and validate them
        Dim temp_p1name, temp_p2name, temp_gametitle As String      'These store temporary output from the text boxes
        temp_p1name = txtP1name.Text                                'Get the output
        temp_p2name = txtP2name.Text
        temp_gametitle = txtGameName.Text
        If validate_initials(temp_p1name) Then                      'This validates the initials that were entered
            frmGame.p1name = CStr(temp_p1name)                      'This sets the public p1name to the value entered
        Else
            error1 = "The Initials entered for Player 1 are invalid" 'Set error1 to this text because the data was entered incorrectly
            txtP1name.Clear()                                       'Clear the textbox
            valid = False                                           'Valid is set to false to indicate that at least one value was wrong
        End If
        If validate_initials(temp_p2name) Then                      'Validate the initials entered for player 2
            frmGame.p2name = CStr(temp_p2name)
        Else
            error2 = "The Initials entered for Player 2 are invalid"
            txtP2name.Clear()
            valid = False
        End If
        If validate_game(temp_gametitle) Then                       'Validate the game name that was entered against different criteria
            game_title = CStr(temp_gametitle)
        Else
            error3 = "Title for the game is invalid"
            txtGameName.Clear()
            valid = False
        End If
        frmGame.turn = "p1"
    End Sub

    Private Sub create_board()
        'This sub must create the board which consists of 3 text boxes, 3 labels and 3 buttons
        Dim x_ratio As Single = Screen.PrimaryScreen.Bounds.Width                   'get the width of the screen in pixels
        Dim y_ratio As Single = Screen.PrimaryScreen.Bounds.Height                  'get the height of the screen in pixels
        Dim x_coordinate, y_coordinate As Integer
        Me.Size = New System.Drawing.Size(0.2 * x_ratio, 0.3 * y_ratio)             'set the size of the board
        Me.Location = New System.Drawing.Point(0.42 * x_ratio, 0.25 * y_ratio)      'set the location of the board
        x_ratio = x_ratio / 1920                                                    'get the ratio of the x and y bounds compared to a  1080p screen
        y_ratio = y_ratio / 1080
        Me.Text = "New Game"                                                        'set the text of the board
        x_coordinate = 40 * x_ratio
        y_coordinate = 40 * y_ratio
        For i = 1 To 3                                                              'the text boxes and lables are in arrays
            x_coordinate = x_coordinate + 200 * x_ratio                             'x changes for the textboxes and labels by 200
            With txt_array(i)
                .Size = New System.Drawing.Size(100 * x_ratio, 45 * y_ratio)
                .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
            End With
            x_coordinate = x_coordinate - 200 * x_ratio                             'This resets x for the labels
            With lbl_array(i)                                                       'Set the values for the labels
                .Size = New System.Drawing.Size(165 * x_ratio, 45 * y_ratio)
                .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
                .Font = New System.Drawing.Font("standard", Int(13 * x_ratio), FontStyle.Regular)
                Select Case i                                                       'The text needs to be one of three values, so a multiway selection is used
                    Case 1 : .Text = "Game Name"
                    Case 2 : .Text = "Player 1 Initials"
                    Case 3 : .Text = "Player 2 Initials"
                End Select
            End With
            y_coordinate = y_coordinate + 40 * y_ratio
            Me.Controls.Add(lbl_array(i))                                           'Add both the labels and text boxes to the controls
            Me.Controls.Add(txt_array(i))
        Next
        With btnCreate                                                              'deal with the three buttons that are used
            .Location = New System.Drawing.Point(130 * x_ratio, 200 * y_ratio)
            .Size = New System.Drawing.Size(100 * x_ratio, 40 * y_ratio)
            .Text = "Create Game"
        End With
        With btnCancel
            .Location = New System.Drawing.Point(240 * x_ratio, 200 * y_ratio)
            .Size = New System.Drawing.Size(100 * x_ratio, 40 * y_ratio)
            .Text = "Cancel"
        End With
        With btnHelp
            .Location = New System.Drawing.Point(20 * x_ratio, 200 * y_ratio)
            .Size = New System.Drawing.Size(100 * x_ratio, 40 * y_ratio)
            .Text = "Help"
        End With
        'Add the click events to the buttons
        AddHandler btnHelp.Click, AddressOf Help_Click
        AddHandler btnCreate.Click, AddressOf Create_Click
        AddHandler btnCancel.Click, AddressOf Cancel_Click
        'Add buttons to the controls
        Me.Controls.Add(btnHelp)
        Me.Controls.Add(btnCreate)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub save_to_file()
        'This creates a new text file with the name of game_title + ".txt"
        Dim path As String = frmMain.bin & "\saved games\" & game_title & ".txt"    'The path is declared here to shorten the next statment
        Dim fs As FileStream = File.Create(path)                                                'This uses the import IO module to create a file in the location given
        fs.Close()                                                                              'Close the file
    End Sub

    Private Sub fill_control_arrays()
        'Fill the two control arrays that hold the text boxes and labels
        'input text boxes
        txt_array(1) = txtGameName
        txt_array(2) = txtP1name
        txt_array(3) = txtP2name

        'input labels
        lbl_array(1) = lblGameName
        lbl_array(2) = lblP1name
        lbl_array(3) = lblP2name
    End Sub

    Private Function validate_game(candidate As String) As Boolean
        'This sub is designed to validate a given game title
        'A game title is valid is it has only upper & lower case letter or numbers

        'This is a lambda function that takes an ascii code and checks that it is within an acceptable range
        'That is between 48 & 57 or between 65 & 90 or between 97 & 122
        'It is a good way to avoid large boolean computations within an if statement
        Dim valid = Function(x) (x >= 48 And x <= 57) Or (x >= 65 And x <= 90) Or (x >= 97 And x <= 122)
        If candidate = "" Then                      'Stop a game title being nothing
            Return False
        End If
        For Each character In candidate             'Cycle through each character in the candidate
            If Not valid(Asc(character)) Then       'Check the ascii code using the valid function
                Return False                        'Return false if any of the characters are invalid
            End If
        Next
        Return True                                 'Return True in all other cases
    End Function

    Private Function validate_initials(initials As String) As Boolean
        'This function will take a string and return a boolean indicating whether it meets the criteria of initials
        'initials are defined as at least one character in length, no more than 5 characters in length
        'all characters are capital letters of the english alphabet (ascii 65 to 90)
        If IsNothing(initials) Then                         'checks that nothing is not entered
            Return False
        End If
        If Len(initials) <= 0 Or Len(initials) >= 6 Then    'checks the length of the input
            Return False
        End If
        For Each i In initials                              'cycles through each character initials
            If Asc(i) < 65 Or Asc(i) > 90 Then              'checks that the characters ascii number is that of a capital letter
                Return False
            End If
        Next
        Return True                                         'return true in all other cases
    End Function

End Class