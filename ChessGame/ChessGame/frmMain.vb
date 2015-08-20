Public Class frmMain
    Public bin As String = Application.StartupPath 'This variable will be used to store the path to the bin folder
#Region "dynamic variables"
    Dim btnNewGame, btnLoadGame, btnQuit, btnHelp, btnRules, btnDelGame, btnClear As New Button
    Dim btn_array(7) As Button
#End Region
    Private Sub frmMain_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        'this creates the form when the game starts to load
        fill_btn_array()    'Fills the button array
        create_board()      'Dynamically creates the board
    End Sub

    Private Sub Load_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This opens the load form
        frmLoad.Show()
        Me.Hide()
    End Sub

    Private Sub New_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This sub is activated when the new game button is clicked
        Me.Hide()               'This hides the main window
        frmNewGame.Show()       'This moves to the new game window for the player to enter their details
    End Sub

    Private Sub Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This opens the delete form
        Me.Hide()
        frmDeleteGame.Show()
    End Sub

    Private Sub Clear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This button clears the high scores file
        Dim clicked As MsgBoxResult                                         'This is set to the result from a message box button click
        clicked = MsgBox("Are you sure you want to clear the scores?" & vbNewLine & _
               "This cannot be undone", MsgBoxStyle.OkCancel)               'The last argument defines the type of message box displayed, in this case it is an ok/cancel
        If clicked = MsgBoxResult.Ok Then                                   'clicked is set to the result and if the ok button was clicked, clicked is set to MsgBoxResult.Ok
            Dim path As String = bin & "\scores.txt"                        'Path is set here to avoid long, unwieldy sentences
            System.IO.File.Open(path, System.IO.FileMode.Truncate).Close()  'This code opens the file in truncate mode and then closes it
        End If
    End Sub

    Private Sub Help_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Rules_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This opens a pdf that has the rules of chess
        'for this reason, Adobe reader is a required installation for the program to work properly
        System.Diagnostics.Process.Start(bin & "\help material\rules of chess.pdf")
    End Sub

    Private Sub Quit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'End the game
        End
    End Sub

    Private Sub create_board()
        'This creates the five buttons that are used to access certain functionality of the game
        Dim x_coordinate, y_coordinate As Integer
        Dim x_ratio As Single = Screen.PrimaryScreen.Bounds.Width                   'get the width of the screen in pixels
        Dim y_ratio As Single = Screen.PrimaryScreen.Bounds.Height                  'get the height of the screen in pixels
        Me.Size = New System.Drawing.Size(0.165 * x_ratio, 0.5 * y_ratio)           'set the size of the form in relation to the screen
        Me.Location = New System.Drawing.Point(0.42 * x_ratio, 0.25 * y_ratio)      'set the location of the form in relation to the screen
        x_ratio = x_ratio / 1920                                            'set the x and y ratios of the screen in relation to a 1920x1080 screen
        y_ratio = y_ratio / 1080
        x_coordinate = 70 * x_ratio
        y_coordinate = 40 * y_ratio
        Me.Text = "Chess Game"
        For i = 1 To 7                                                              'This cycles through the button array
            With btn_array(i)
                .Size = New System.Drawing.Size(150 * x_ratio, 45 * y_ratio)        'The size is the same for all the buttons
                .Location = New System.Drawing.Point(x_coordinate, y_coordinate)    'The location stays the same in the x co-ordinate and changes in the y co-ordinate
                .Font = New System.Drawing.Font("standard", Int(13 * x_ratio), FontStyle.Regular)  'The font is the same
                Select Case i                                                       'This selects what text to display on each button as well as which click event to assign
                    Case 1
                        .Text = "New Game"
                        AddHandler btn_array(i).Click, AddressOf New_Click
                    Case 2
                        .Text = "Load Game"
                        AddHandler btn_array(i).Click, AddressOf Load_Click
                    Case 3
                        .Text = "Delete Game"
                        AddHandler btn_array(i).Click, AddressOf Delete_Click
                    Case 4
                        .Text = "Clear Scores"
                        AddHandler btnClear.Click, AddressOf Clear_Click
                    Case 5
                        .Text = "Rules"
                        AddHandler btn_array(i).Click, AddressOf Rules_Click
                    Case 6
                        .Text = "Help"
                        AddHandler btn_array(i).Click, AddressOf Help_Click
                    Case 7
                        .Text = "Quit"
                        AddHandler btn_array(i).Click, AddressOf Quit_Click
                End Select
            End With
            y_coordinate = y_coordinate + 60 * y_ratio                              'The y co-ordinate increments by 60 to allow for a 15 pixel gap between button
            Me.Controls.Add(btn_array(i))                                           'Add to controls
        Next i
    End Sub

    Private Sub fill_btn_array()
        'This places all the buttons into an array for easy management
        btn_array(1) = btnNewGame
        btn_array(2) = btnLoadGame
        btn_array(3) = btnDelGame
        btn_array(4) = btnClear
        btn_array(5) = btnRules
        btn_array(6) = btnHelp
        btn_array(7) = btnQuit
    End Sub

End Class