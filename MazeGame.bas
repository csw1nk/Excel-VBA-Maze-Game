 Dim mazeSize As Integer
 Dim blackSquareCount As Integer
Sub MazeGame()

' Module: MazeGame
' Version: 1.1
' Date: 12-16-2023
' Description: Updated Fixes for MazeGame

'set game details
Const GameVersion As String = "1.0.0"
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
        If ws.Cells(i, j).Interior.color = RGB(255, 255, 255) Then
         If Rnd() < density Then ws.Cells(i, j).Interior.color = RGB(0, 0, 0)
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
ws.Cells(startRow + 2, startColumn - 2).Value = "Start here -->"
ws.Cells(startRow + 2, startColumn - 2).Font.color = RGB(255, 255, 255)
ws.Cells(startRow + 2, startColumn - 2).Interior.ColorIndex = 13
ws.Cells(startRow + 2, startColumn - 1).Interior.ColorIndex = 41 'STARTING POINT
ws.Cells(startRow + 2, startColumn).Interior.color = xlNone
ws.Cells(startRow + 2, startColumn + 1).Interior.color = xlNone
ws.Cells(startRow + 2, startColumn + 2).Interior.color = xlNone
ws.Cells(startRow + 2, startColumn + 3).Interior.color = xlNone
ws.Cells(startRow + 1, startColumn + 2).Interior.color = xlNone
ws.Cells(startRow + 3, startColumn + 2).Interior.color = xlNone
ws.Cells(startRow + 3, startColumn).Interior.color = RGB(0, 0, 0)
ws.Cells(startRow + 1, startColumn).Interior.color = RGB(0, 0, 0)

' Set the exit point
Dim exitRow As Integer, exitColumn As Integer
exitRow = mazeSize ' Last row of the maze
exitColumn = mazeSize ' Last column of the maze

' Clear the exit point cells
ws.Cells(exitRow + 1, exitColumn - 1).Value = "<-- Exit"
ws.Cells(exitRow + 1, exitColumn - 2).Interior.ColorIndex = 44 ' EXIT POINT
ws.Cells(exitRow - 1, exitColumn).Interior.color = xlNone
ws.Cells(exitRow, exitColumn - 2).Interior.color = xlNone
ws.Cells(exitRow - 1, exitColumn - 2).Interior.color = xlNone
ws.Cells(exitRow - 2, exitColumn - 2).Interior.color = xlNone
ws.Cells(exitRow - 2, exitColumn - 3).Interior.color = xlNone
ws.Cells(exitRow - 2, exitColumn - 1).Interior.color = xlNone
ws.Cells(exitRow - 1, exitColumn - 1).Interior.color = RGB(0, 0, 0)
ws.Cells(exitRow - 1, exitColumn - 3).Interior.color = RGB(0, 0, 0)
ws.Cells(exitRow + 2, exitColumn - 1).Value = "Controls-->"
   
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
        ' Check if the cell to the right is within bounds and not a wall
         If currentPlayerCell.Row > 1 Then
            Dim cellAbove As Range
            Set cellAbove = ws.Cells(currentPlayerCell.Row - 1, currentPlayerCell.Column)

            ' Check for black square
            If cellAbove.Interior.color = RGB(0, 0, 0) Then
                ' Increment counter for black square encounters
                blackSquareCount = blackSquareCount + 1

                ' On second encounter, offer to activate superpower
                If blackSquareCount = 2 Then
                    Dim response As VbMsgBoxResult
                    response = MsgBox("Do you want to Hulk Smash?", vbYesNo + vbQuestion, "Activate Superpower")

                     If response = vbYes Then
                    ' Flash effect
                    Call FlashEffect1(ws, currentPlayerCell, RGB(0, 255, 0)) ' Flash green
                        ' Activate superpower: move through the black square
                        cellAbove.Interior.ColorIndex = 41
                        currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                        blackSquareCount = 0 ' Reset the count
                    Else
                        blackSquareCount = 0 ' Reset the count
                        MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
                    End If
                Else
                    MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
                End If
                Exit Sub
            End If

            ' Regular movement logic
            If cellAbove.Interior.ColorIndex <> 13 Then
                ' Move the player to the new position
                cellAbove.Interior.ColorIndex = 41
                currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
            Else
                MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
            End If
        Else
            MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
        End If
    Else
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

    ' Search for the player's cell
    For Each cell In ws.UsedRange
        If cell.Interior.ColorIndex = 41 Then
            Set currentPlayerCell = cell
            found = True
            Exit For
        End If
    Next cell

    If found Then
        ' Check if the cell to the right is within bounds and not a wall
        If currentPlayerCell.Column < ws.Columns.Count Then
            Dim cellToRight As Range
            Set cellToRight = ws.Cells(currentPlayerCell.Row, currentPlayerCell.Column + 1)

            ' Check for black square
            If cellToRight.Interior.color = RGB(0, 0, 0) Then
                ' Increment counter for black square encounters
                blackSquareCount = blackSquareCount + 1

                ' On second encounter, offer to activate superpower
                If blackSquareCount = 2 Then
                    Dim response As VbMsgBoxResult
                    response = MsgBox("Do you want to Hulk Smash?", vbYesNo + vbQuestion, "Activate Superpower")

                     If response = vbYes Then
                    ' Flash effect
                    Call FlashEffect1(ws, currentPlayerCell, RGB(0, 255, 0)) ' Flash green
                        ' Activate superpower: move through the black square
                        cellToRight.Interior.ColorIndex = 41
                        currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                        blackSquareCount = 0 ' Reset the count
                              Else
                        blackSquareCount = 0 ' Reset the count
                        MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
                    End If
                Else
                     MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
                End If
                Exit Sub
            End If

            ' Regular movement logic
            If cellToRight.Interior.ColorIndex <> 13 Then
                ' Move the player to the new position
                cellToRight.Interior.ColorIndex = 41
                currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
            Else
                MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
            End If
        Else
            MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
        End If
    Else
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
        ' Check if the cell to the right is within bounds and not a wall
       If currentPlayerCell.Row < ws.Rows.Count Then
        Dim cellBelow As Range
        Set cellBelow = ws.Cells(currentPlayerCell.Row + 1, currentPlayerCell.Column)

            ' Check for black square
            If cellBelow.Interior.color = RGB(0, 0, 0) Then
                ' Increment counter for black square encounters
                blackSquareCount = blackSquareCount + 1

                ' On second encounter, offer to activate superpower
                If blackSquareCount = 2 Then
                    Dim response As VbMsgBoxResult
                    response = MsgBox("Do you want to Hulk Smash?", vbYesNo + vbQuestion, "Activate Superpower")

                     If response = vbYes Then
                    ' Flash effect
                    Call FlashEffect1(ws, currentPlayerCell, RGB(0, 255, 0)) ' Flash green
                        ' Activate superpower: move through the black square
                        cellBelow.Interior.ColorIndex = 41
                        currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                        blackSquareCount = 0 ' Reset the count
                               Else
                        blackSquareCount = 0 ' Reset the count
                        MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
                    End If
                Else
                    MsgBox "Oops, you can't go that way!"
                End If
                Exit Sub
            End If

            ' Regular movement logic
            If cellBelow.Interior.ColorIndex <> 13 Then
                ' Move the player to the new position
                cellBelow.Interior.ColorIndex = 41
                currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                  If currentPlayerCell.Row = mazeSize And currentPlayerCell.Column = mazeSize - 2 Then
            MsgBox "Congratulations, you won the game!", vbInformation, "Game Over"
        End If
            Else
                MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
            End If
        Else
            MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
        End If
    Else
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
        ' Check if the cell to the right is within bounds and not a wall
        If currentPlayerCell.Column > 1 Then
        Dim cellToLeft As Range
        Set cellToLeft = ws.Cells(currentPlayerCell.Row, currentPlayerCell.Column - 1)

            ' Check for black square
            If cellToLeft.Interior.color = RGB(0, 0, 0) Then
                ' Increment counter for black square encounters
                blackSquareCount = blackSquareCount + 1

                ' On second encounter, offer to activate superpower
                If blackSquareCount = 2 Then
                    Dim response As VbMsgBoxResult
                    response = MsgBox("Do you want to Hulk Smash?", vbYesNo + vbQuestion, "Activate Superpower")

                     If response = vbYes Then
                    ' Flash effect
                    Call FlashEffect1(ws, currentPlayerCell, RGB(0, 255, 0)) ' Flash green
                        ' Activate superpower: move through the black square
                        cellToLeft.Interior.ColorIndex = 41
                        currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                        blackSquareCount = 0 ' Reset the count
                               Else
                        blackSquareCount = 0 ' Reset the count
                        MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
                    End If
                Else
                    MsgBox "Oops, you can't go that way!", vbQuestion, "Maze Game"
                End If
                Exit Sub
            End If

            ' Regular movement logic
            If cellToLeft.Interior.ColorIndex <> 13 Then
                ' Move the player to the new position
                cellToLeft.Interior.ColorIndex = 41
                currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
            Else
                MsgBox "Oops, you can't go that way!"
            End If
        Else
            MsgBox "Oops, you can't go that way!"
        End If
    Else
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

Sub FlashEffect(ws As Worksheet, cell As Range, color As Long)
    Dim originalColor As Long
    originalColor = cell.Interior.color
    cell.Interior.color = color
    Application.Wait (Now + TimeValue("0:00:02")) ' Wait 1 second
    cell.Interior.color = originalColor
End Sub

Sub FlashEffect1(ws As Worksheet, cell As Range, color As Long)
    Dim r As Integer, c As Integer
    Dim originalColors() As Long
    ReDim originalColors(1 To 3, 1 To 3)

    ' Store original colors and apply flash color
    For r = -1 To 1
        For c = -1 To 1
            With ws.Cells(cell.Row + r, cell.Column + c)
                If .Interior.ColorIndex <> 13 Then ' Avoid changing border cells
                    originalColors(r + 2, c + 2) = .Interior.color
                    .Interior.color = color
                Else
                    originalColors(r + 2, c + 2) = .Interior.color ' Store color but don't change
                End If
            End With
        Next c
    Next r

    Application.Wait (Now + TimeValue("0:00:03")) ' Wait 3 seconds

    ' Revert to original colors
    For r = -1 To 1
        For c = -1 To 1
            With ws.Cells(cell.Row + r, cell.Column + c)
                If .Interior.ColorIndex <> 13 Then ' Avoid changing border cells
                    .Interior.color = originalColors(r + 2, c + 2)
                End If
            End With
        Next c
    Next r
End Sub


