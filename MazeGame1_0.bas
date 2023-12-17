Attribute VB_Name = "MazeGame"
'----------------------

 Dim mazeSize As Integer
Sub MazeGame()

' Module: MazeGame
' Version: 1.0
' Date: 12-16-2023
' Description: Updated Fixes for MazeGame

'set game details
Const GameVersion As String = "1.0"
Const GameName As String = "MazeGame1.0"
Const GameAuthor As String = "Corey Swink"
Const GameDescription As String = "My first program ever built in code, many more to come"

    Dim userResponse As Integer
    Dim ws As Worksheet
    Dim i As Integer, j As Integer
    Dim clearPathColumn As Integer
    Dim density As Double
    
' Set ws to the active sheet
 Set ws = ActiveSheet
    mazeSize = 25 ' Adjust the size of the maze (dynamic)
    density = 0.305  ' Adjust the density of walls (dynamic)

' set starting message box to initiate game
    userResponse = MsgBox("Are You Ready To Play?", vbYesNo + vbQuestion, "Maze Game")
    If userResponse = vbYes Then
' Clear previous game settings if any
        ws.Cells.Clear
' Clear Buttons
     Dim btn As Object
     For Each btn In ws.Buttons
         btn.Delete
     Next btn

' Draw the complete border around the maze
    For i = 1 To mazeSize
    For j = 2 To mazeSize + 1
    With ws.Cells(i, j)
                If i = 1 Or i = mazeSize Or j = 2 Or j = mazeSize + 1 Then
    .Interior.ColorIndex = 13 ' Border walls Color
    End If
     End With
       Next j
         Next i
            
' Randomly add walls inside the maze using density
    Randomize ' Initialize random seed
      For i = 2 To mazeSize - 1
      For j = 3 To mazeSize
        If ws.Cells(i, j).Interior.Color = RGB(255, 255, 255) Then
         If Rnd() < density Then ws.Cells(i, j).Interior.Color = RGB(0, 0, 0)
           End If
        Next j
        Next i
        
' Assuming the maze is drawn starting from column B (column index 2)
' Adjust startX to be to the right of the maze

Dim lastMazeColumn As Integer
lastMazeColumn = 1 + mazeSize ' Adjust this if your maze starts from a different column

' Calculate the starting X position for the buttons based on column widths
startX = ws.Cells(1, lastMazeColumn).Left + ws.Columns(lastMazeColumn).Width
startY = ws.Cells(mazeSize - 4, 1).Top ' next to the maze
btnHeight = ws.Rows(mazeSize + 1).Height
btnWidth = 100 ' Adjust the width if the button text is longer

' Create new buttons to the right of the maze
btnNames = Array("Up", "Down", "Left", "Right", "Reset Game")

For i = 0 To UBound(btnNames)
    Set btn = ws.Buttons.Add(startX, startY + (i * btnHeight), btnWidth, btnHeight)
    With btn
        .Caption = btnNames(i)
        If btnNames(i) = "Reset Game" Then
        .OnAction = "ResetGame" 'set button to macro "
        Else
        .OnAction = "MovePlayer" & btnNames(i)
        End If
        .Name = "btn" & btnNames(i)
    End With
Next i
    
' Set the starting point
Dim startRow As Integer, startColumn As Integer
startRow = 2 ' Second row to avoid placing it on the border
startColumn = 3 ' Third column to avoid placing it on the border

' Clear the starting point cells
ws.Cells(startRow + 2, startColumn - 2).value = "Start here -->"
ws.Cells(startRow + 2, startColumn - 2).Font.Color = RGB(255, 255, 255)
ws.Cells(startRow + 2, startColumn - 2).Interior.ColorIndex = 13
ws.Cells(startRow + 2, startColumn - 1).Interior.ColorIndex = 41 'STARTING POINT
ws.Cells(startRow + 2, startColumn).Interior.Color = xlNone
ws.Cells(startRow + 2, startColumn + 1).Interior.Color = xlNone
ws.Cells(startRow + 2, startColumn + 2).Interior.Color = xlNone
ws.Cells(startRow + 2, startColumn + 3).Interior.Color = xlNone
ws.Cells(startRow + 1, startColumn + 2).Interior.Color = xlNone
ws.Cells(startRow + 3, startColumn + 2).Interior.Color = xlNone
ws.Cells(startRow + 3, startColumn).Interior.Color = RGB(0, 0, 0)
ws.Cells(startRow + 1, startColumn).Interior.Color = RGB(0, 0, 0)

' Set the exit point
Dim exitRow As Integer, exitColumn As Integer
exitRow = mazeSize ' Last row of the maze
exitColumn = mazeSize ' Last column of the maze

' Clear the exit point cells
ws.Cells(exitRow + 1, exitColumn - 1).value = "<-- Exit"
ws.Cells(exitRow + 1, exitColumn - 2).Interior.ColorIndex = 44 ' EXIT POINT
ws.Cells(exitRow - 1, exitColumn).Interior.Color = xlNone
ws.Cells(exitRow, exitColumn - 2).Interior.Color = xlNone
ws.Cells(exitRow - 1, exitColumn - 2).Interior.Color = xlNone
ws.Cells(exitRow - 2, exitColumn - 2).Interior.Color = xlNone
ws.Cells(exitRow - 2, exitColumn - 3).Interior.Color = xlNone
ws.Cells(exitRow - 2, exitColumn - 1).Interior.Color = xlNone
ws.Cells(exitRow - 1, exitColumn - 1).Interior.Color = RGB(0, 0, 0)
ws.Cells(exitRow - 1, exitColumn - 3).Interior.Color = RGB(0, 0, 0)
ws.Cells(exitRow + 2, exitColumn - 1).value = "Controls-->"
   
'Format Cells
ws.Columns("A:A").AutoFit
        
Dim squareSize As Integer
    squareSize = 20 ' You can adjust this value as needed
    
  ' Set columns to be as wide as the square size
    ws.Columns("B:AE").ColumnWidth = squareSize / ws.StandardWidth * 1
    
  ' Set rows to be as tall as the square size
    ws.Rows("1:" & mazeSize).RowHeight = squareSize
'End Format Cells

'Have to nest all the game together through this else statement for game start
        MsgBox "Maze generated! Find your way to the end!", vbQuestion, "Maze Game"
    Else
        MsgBox "Maybe next time:)", vbQuestion, "Maze Game"
    End If
End Sub

Sub MovePlayerUp()
    Dim ws As Worksheet
     Set ws = ActiveSheet

    Dim currentPlayerCell As Range
    Dim cell As Range
    Dim found As Boolean
    found = False

    ' We'll search for the cell with the specific fill color
    For Each cell In ws.UsedRange
        If cell.Interior.ColorIndex = 41 Then
            Set currentPlayerCell = cell
            found = True
            Exit For
        End If
    Next cell

    If found Then
            ' Check if the cell above is within the maze bounds and not a wall
        If currentPlayerCell.Row > 1 Then
            Dim cellAbove As Range
            Set cellAbove = ws.Cells(currentPlayerCell.Row - 1, currentPlayerCell.Column)
            If cellAbove.Interior.Color <> RGB(0, 0, 0) And cellAbove.Interior.ColorIndex <> 13 Then
                ' Color the current cell to leave a trail
                currentPlayerCell.Interior.ColorIndex = 48 ' Silver for the trail

                ' Move the "player" to the new position by setting the fill color
                cellAbove.Interior.ColorIndex = 41
            Else
                ' Display a message if the player can't move right
                MsgBox "Oops, you can't go that way!"
            End If
        Else
            ' Player is in the last column and cannot move right
            MsgBox "Oops, you can't go that way!"
        End If
    Else
        ' Player starting position not found.
        MsgBox "Player starting position not found."
    End If
End Sub

Sub MovePlayerRight()
    Dim ws As Worksheet
     Set ws = ActiveSheet

    Dim currentPlayerCell As Range
    Dim cell As Range
    Dim found As Boolean
    found = False

    ' We'll search for the cell with the specific fill color
    For Each cell In ws.UsedRange
        If cell.Interior.ColorIndex = 41 Then
            Set currentPlayerCell = cell
            found = True
            Exit For
        End If
    Next cell

    If found Then
        ' Check if the cell to the right is within the maze bounds and not a wall
        If currentPlayerCell.Column < ws.Columns.count Then
            Dim cellToRight As Range
            Set cellToRight = ws.Cells(currentPlayerCell.Row, currentPlayerCell.Column + 1)
            If cellToRight.Interior.Color <> RGB(0, 0, 0) And cellToRight.Interior.ColorIndex <> 13 Then
                ' Color the current cell to leave a trail
                currentPlayerCell.Interior.ColorIndex = 48 ' Silver for the trail

                ' Move the "player" to the new position by setting the fill color
                cellToRight.Interior.ColorIndex = 41
            Else
                ' Display a message if the player can't move right
                MsgBox "Oops, you can't go that way!"
            End If
        Else
            ' Player is in the last column and cannot move right
            MsgBox "Oops, you can't go that way!"
        End If
    Else
        ' Player starting position not found.
        MsgBox "Player starting position not found."
    End If
End Sub

Sub MovePlayerDown()
    Dim ws As Worksheet
     Set ws = ActiveSheet

    Dim currentPlayerCell As Range
    Dim cell As Range
    Dim found As Boolean
    found = False

    ' We'll search for the cell with the specific fill color
    For Each cell In ws.UsedRange
        If cell.Interior.ColorIndex = 41 Then
            Set currentPlayerCell = cell
            found = True
            Exit For
        End If
    Next cell

    If found Then
        ' Check if the cell to the right is within the maze bounds and not a wall
      If currentPlayerCell.Row < ws.Rows.count Then
        Dim cellBelow As Range
        Set cellBelow = ws.Cells(currentPlayerCell.Row + 1, currentPlayerCell.Column)
        If cellBelow.Interior.Color <> RGB(0, 0, 0) And cellBelow.Interior.ColorIndex <> 13 Then
                ' Color the current cell to leave a trail
                currentPlayerCell.Interior.ColorIndex = 48 ' Silver for the trail

                ' Move the "player" to the new position by setting the fill color
                cellBelow.Interior.ColorIndex = 41
     If currentPlayerCell.Row = mazeSize And currentPlayerCell.Column = mazeSize - 2 Then
                         MsgBox "Congratulations, you won the game!", vbInformation, "Game Over"
End If
            Else
                ' Display a message if the player can't move right
                MsgBox "Oops, you can't go that way!"
            End If
        Else
            ' Player is in the last column and cannot move right
            MsgBox "Oops, you can't go that way!"
        End If
    Else
        ' Player starting position not found.
        MsgBox "Player starting position not found."
    End If
End Sub

Sub MovePlayerLeft()
    Dim ws As Worksheet
     Set ws = ActiveSheet

    Dim currentPlayerCell As Range
    Dim cell As Range
    Dim found As Boolean
    found = False

    ' We'll search for the cell with the specific fill color
    For Each cell In ws.UsedRange
        If cell.Interior.ColorIndex = 41 Then
            Set currentPlayerCell = cell
            found = True
            Exit For
        End If
    Next cell

    If found Then
        ' Check if the cell to the right is within the maze bounds and not a wall
   If currentPlayerCell.Column > 1 Then
        Dim cellToLeft As Range
        Set cellToLeft = ws.Cells(currentPlayerCell.Row, currentPlayerCell.Column - 1)
        If cellToLeft.Interior.Color <> RGB(0, 0, 0) And cellToLeft.Interior.ColorIndex <> 13 Then
                ' Color the current cell to leave a trails
                currentPlayerCell.Interior.ColorIndex = 48 '  silver for the trail

                ' Move the "player" to the new position by setting the fill color
                cellToLeft.Interior.ColorIndex = 41
            Else
                ' Display a message if the player can't move right
                MsgBox "Oops, you can't go that way!"
            End If
        Else
            ' Player is in the last column and cannot move right
            MsgBox "Oops, you can't go that way!"
        End If
    Else
        ' Player starting position not found.
        MsgBox "Player starting position not found."
    End If
End Sub

Sub ResetGame()
   ' Ask the user if they are sure about resetting the game
    Dim response As VbMsgBoxResult
    response = MsgBox("Reset game?", vbYesNo + vbQuestion, "Reset Game")

    ' If the user clicks 'Yes', reset the game
    If response = vbYes Then
        ' Call the MazeGame macro to restart the game
        MazeGame
    End If
End Sub
