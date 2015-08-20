Public Class frmPawnSelection
    Public promotion_selection As String

#Region "dynamic variables"
    Dim rbtnRook, rbtnHorse, rbtnBishop, rbtnQueen As New RadioButton
    Dim lblHeading As New Label
    Dim btnSelect As New Button
    Dim rbtn_array(4) As RadioButton
#End Region

    Private Sub PawnSelection_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        'The form is opened as a customised message box so code execution will cease in the game form once this is opened
        'When activated, this sub dynamically creates the  board
        fill_control_array()
        create_board()
        promotion_selection = "q"
        'This sets the selection to queen if the cancel button is clicked.
        'thus avoiding a "file not found" error when the game tries to load a non-existent image
    End Sub

    Private Sub Select_Click(sender As Object, e As System.EventArgs)
        'This button submits the selected choice of promotion
        For i = 1 To 4                                      'Cycle through the array of radio buttons and find the selected one
            If rbtn_array(i).Checked Then
                promotion_selection = rbtn_array(i).Text    'get the text value of the selected radio button
            End If
        Next i
        Select Case promotion_selection                     'This converts the button text to a one letter acronym (one letter because it is independent of color)
            Case "Rook" : promotion_selection = "r"
            Case "Knight" : promotion_selection = "h"
            Case "Bishop" : promotion_selection = "b"
            Case "Queen" : promotion_selection = "q"
        End Select
        Me.Close()                                          'Close the form
    End Sub

    Private Sub create_board()
        'This creates the board elements
        Dim y_coordinate, x_coordinate As Integer
        Dim x_ratio As Single = Screen.PrimaryScreen.Bounds.Width
        Dim y_ratio As Single = Screen.PrimaryScreen.Bounds.Height
        Me.Size = New System.Drawing.Size(250, 300)             'set the size of the board
        Me.Location = New System.Drawing.Point(0.42 * Width, 0.25 * Height)      'set the location of the board
        x_coordinate = 65                         'The x co-ordinate remains constant for all the radio buttons
        y_coordinate = 20                               'The y co-ordinate starts at 55 pixels but is set at 20 to account for the increase in the for loop
        For i = 1 To 4                                  'Cycle through the array of radio buttons
            y_coordinate = y_coordinate + 35            'Set a 35 pixel gap between the radio buttons
            With rbtn_array(i)
                .Location = New System.Drawing.Point(x_coordinate, y_coordinate)
                .Size = New System.Drawing.Size(80, 30)
                Select Case i                           'select what text to add to the radio buttons
                    Case 1 : .Text = "Rook"
                    Case 2 : .Text = "Knight"
                    Case 3 : .Text = "Bishop"
                    Case 4 : .Text = "Queen"
                End Select
                Me.Controls.Add(rbtn_array(i))          'add to the controls
            End With
        Next i
        'define the select button
        With btnSelect
            .Location = New System.Drawing.Point(70, 190)
            .Size = New System.Drawing.Size(70, 30)
            .Text = "Select"
        End With
        'add to the controls
        Me.Controls.Add(btnSelect)
        'add click event
        AddHandler btnSelect.Click, AddressOf Select_Click

        'define the lable
        With lblHeading
            .Location = New System.Drawing.Point(20, 15)
            .Size = New System.Drawing.Size(210, 40)
            .Text = "Your Pawn has been promoted, what would you like to select?"
        End With
        'Add to the controls
        Me.Controls.Add(lblHeading)
    End Sub

    Private Sub fill_control_array()
        'fill the control arrays that are used to hold the radio buttons
        rbtn_array(1) = rbtnRook
        rbtn_array(2) = rbtnHorse
        rbtn_array(3) = rbtnBishop
        rbtn_array(4) = rbtnQueen
    End Sub
End Class