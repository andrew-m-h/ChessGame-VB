Public Class frmDeleteGame
    Dim games As New ArrayList()            'Used for it’s add and remove methods which aren't available with arrays
#Region "dynamic variables"
    'These objects are used in the dynamic creation of the board
    Dim lstGames As New ListBox
    Dim lblHeading As New Label
    Dim btnDelete, btnCancel As New Button
#End Region

    Private Sub frmDeleteGame_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        'This activates when the player clicks the delete button on the main form
        create_board()      'This creates the board
        get_games()         'This looks in the "saved games" folder and gets a list of all files there
    End Sub

    Private Sub Cancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        frmMain.Show()  'This quits the form
        Me.Close()
    End Sub

    Private Sub Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'This activates when the delete button is clicked
        'It gets the selected item in the list box and deletes the appropriate file
        Dim index As Integer
        Dim path As String = frmMain.bin & "\saved games\"          'This is the path that all the game files are located in
        index = lstGames.SelectedIndex                                          'This is the index of the selected item in the listbox
        If index <> -1 Then                                                     'index is -1 if nothing is selected
            lstGames.Items.RemoveAt(index)                                      'This removes the item from the listbox
            My.Computer.FileSystem.DeleteFile(path & games(index) & ".txt")     'This deletes the file
            games.RemoveAt(index)                                               'an Arraylist has a RemoveAt method which is used here for efficiency
        End If

    End Sub

    Private Sub create_board()
        'This creates the board which is a simple list box, label and two buttons
        'This sub has to consider the screen resolution of the screen running the program compared to a 1080p screen
        Dim x_ratio As Single = Screen.PrimaryScreen.Bounds.Width   'Get the screen dimensions in pixels
        Dim y_ratio As Single = Screen.PrimaryScreen.Bounds.Height
        Me.Size = New System.Drawing.Size(x_ratio * 0.15, y_ratio * 0.4)        'This sets the size of the form in relation to that of the screen
        Me.Location = New System.Drawing.Point(0.42 * x_ratio, 0.25 * y_ratio)  'Set the location of the form in relation to that of the screen
        x_ratio = x_ratio / 1920                                                'The x ratio is the width of the screen / width of my screen (1920)
        y_ratio = y_ratio / 1080                                                'The y ratio is thae height of the screen / height of my screen (1080)
        Me.Text = "Delete Game"                                                 'Set the text of the form
        With lstGames                                                           'Describe the list box in relation to the screen
            .Size = New System.Drawing.Size(140 * x_ratio, 250 * y_ratio)
            .Location = New System.Drawing.Point(50 * x_ratio, 40 * y_ratio)
        End With
        'Add to controls
        Me.Controls.Add(lstGames)

        With lblHeading                                                         'Describe the heading label in relation to the screen
            .Size = New System.Drawing.Size(200 * x_ratio, 30 * y_ratio)
            .Location = New System.Drawing.Point(35 * x_ratio, 10 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(13 * x_ratio), FontStyle.Underline)
            .Text = "Available Games"
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

        With btnDelete                                                          'Describe the Delete button
            .Location = New System.Drawing.Point(130 * x_ratio, 320 * y_ratio)
            .Size = New System.Drawing.Size(115 * x_ratio, 35 * y_ratio)
            .Font = New System.Drawing.Font("standard", Int(12 * x_ratio), FontStyle.Regular)
            .Text = "Delete"
        End With
        'Add to controls
        Me.Controls.Add(btnDelete)
        'Add Click Event
        AddHandler btnDelete.Click, AddressOf Delete_Click
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
End Class