Attribute VB_Name = "MazeGame"



'--------------------
Dim conn As ADODB.Connection
Dim mazeSize As Integer
 Dim blackSquareCount As Integer
 Dim moveCount As Integer
   Dim startTime As Date
    Dim endTime As Date
    Dim density As Double
    Dim UniqueID As String
    Dim gameDuration As Long
    
Sub MazeGame()

' Module: MazeGame
' Version: 1.2
' Date: 12-17-2023
' Description: Adding Record Counting and Database Integration

'set game details
Const GameVersion As String = "1.2"
Const GameName As String = "MazeGame1.2"
Const GameAuthor As String = "Corey Swink"
Const GameDescription As String = "My first learning into local database integrations"

    Dim userResponse As Integer
    Dim ws As Worksheet
    Dim i As Integer, j As Integer
    Dim clearPathColumn As Integer

' Set ws to the active sheet
 Set ws = ActiveSheet
 
Call SetMazeParameters(20, 0.33) ' Set mazeSize and density - density need to be between .2 & .4

' set starting message box to initiate game
    userResponse = MsgBox("Are You Ready To Play?", vbYesNo + vbQuestion, "Maze Game")
    If userResponse = vbYes Then
    ConnectToDatabase 'connects to database for record keeping
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
Debug.Print density
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
        MsgBox "Good luck! Find your way to the end!", vbQuestion, "Maze Game"
        startTime = Now
        Debug.Print "Game Start Time: " & startTime
        moveCount = 0
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
                        moveCount = moveCount + 1  ' Increment move count here, right after the player moves
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
                moveCount = moveCount + 1  ' Increment move count here, right after the player moves
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
                        moveCount = moveCount + 1  ' Increment move count here, right after the player moves
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
                cellToRight.Interior.ColorIndex = 41 'color new cell
                currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                moveCount = moveCount + 1  ' Increment move count here, right after the player moves
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
                If blackSquareCount >= 2 Then
                    Dim response As VbMsgBoxResult
                    response = MsgBox("Do you want to Hulk Smash?", vbYesNo + vbQuestion, "Activate Superpower")

                     If response = vbYes Then
                    ' Flash effect
                    Call FlashEffect1(ws, currentPlayerCell, RGB(0, 255, 0)) ' Flash green
                        ' Activate superpower: move through the black square
                        cellBelow.Interior.ColorIndex = 41 'color new cell
                        currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                        blackSquareCount = 0 ' Reset the count
                        moveCount = moveCount + 1  ' Increment move count here, right after the player moves
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
                cellBelow.Interior.ColorIndex = 41 'color new cell
                currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                moveCount = moveCount + 1  ' Increment move count here, right after the player moves
                  If currentPlayerCell.Row = mazeSize And currentPlayerCell.Column = mazeSize - 2 Then
            'Congrats You Won the Game
             endTime = Now
             gameDuration = DateDiff("s", startTime, endTime)
            MsgBox "Congratulations, you won the game!" & vbNewLine & _
        moveCount & " Moves" & vbNewLine & _
        gameDuration & " seconds", vbInformation, "Game Over"

'important end game paramters that set time difference and inserts data to database
    Debug.Print "Game End Time: " & endTime & vbCrLf & "Total Duration: " & gameDuration & " seconds"
    Debug.Print "Total Moves: " & moveCount
    Debug.Print density
    Debug.Print gameDuration
    GenerateUniqueID
    AppendDataToCSV
    AppendDataToJson
    Application.Wait (Now + TimeValue("0:00:01"))
    InsertGameData UniqueID, startTime, moveCount, gameDuration, mazeSize, density
    CloseDatabaseConnection
    Exit Sub
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
                        cellToLeft.Interior.ColorIndex = 41 'color new cell
                        currentPlayerCell.Interior.ColorIndex = 48 ' Leave a trail
                        blackSquareCount = 0 ' Reset the count
                        moveCount = moveCount + 1  ' Increment move count here, right after the player moves
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
    Application.Wait (Now + TimeValue("0:00:008")) ' Wait 1 second
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

Sub ConnectToDatabase()
    On Error Resume Next ' Turn on error handling

    ' Attempt to create and open the database connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\corey\Personal Projects\Excel-VBA-Maze-Game\DB\MazeGameDB.accdb"
    conn.Open

    ' Check if an error occurred during connection
    If Err.Number <> 0 Then
        ' Handle the error here (e.g., display a message)
        Debug.Print "Error Number: " & Err.Number
        Debug.Print "Error Description: " & Err.Description
        ' Ensure the conn object is set to Nothing if connection failed
        Set conn = Nothing
    End If
    On Error GoTo 0 ' Reset error handling to default
End Sub
Sub InsertGameData(ByVal UniqueID As String, ByVal startTime As Date, ByVal moveCount As Long, ByVal gameDuration As Long, ByVal mazeSize As Integer, ByVal density As Double)
    If conn Is Nothing Then
    Exit Sub
    End If
    
    If conn.State = 0 Then ' 0 means the connection is closed
  Exit Sub
    End If
    Dim sql As String
    sql = "INSERT INTO Data (UniqueID, GameDateTime, PlayerMoves, CompletionTime, MazeSize, Density) VALUES ('" & UniqueID & "', #" & Format(startTime, "yyyy-mm-dd hh:mm:ss") & "#, " & moveCount & ", " & gameDuration & ", " & mazeSize & ", " & Replace(density, ",", ".") & ")"
     On Error Resume Next
    conn.Execute sql
      ' Check for errors
        Debug.Print "Data successfully inserted: " & sql
        On Error GoTo 0 ' Reset error handling to default
End Sub

Sub CloseDatabaseConnection()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

Sub SetMazeParameters(size As Integer, d As Double)
    mazeSize = size ' Set the global mazeSize variable
    density = d ' Assuming density is also a global variable
End Sub

Sub GenerateUniqueID()
    ' Generate a unique ID based on the concatenation of variables
    UniqueID = Format(startTime, "yyyymmddhhmmss") & "_" & CStr(moveCount) & CStr(gameDuration) & CStr(mazeSize) & CInt(density * 1000)
End Sub

Sub AppendDataToCSV()
    ' Specify the file path for the CSV file
    Dim filePath As String
    filePath = "C:\Users\corey\Personal Projects\Excel-VBA-Maze-Game\DB\Data.csv"
    
    ' Determine if the file exists
    Dim fileExists As Boolean
    fileExists = (Dir(filePath) <> "")
    
    ' Open the CSV file in Append mode or create it if it doesn't exist
    Dim fileNum As Integer
    fileNum = FreeFile

    ' Create a CSV string with the data
    Dim csvStr As String
    csvStr = """" & UniqueID & """,""" & Format(startTime, "yyyy-mm-dd hh:mm:ss") & """," & moveCount & "," & gameDuration & "," & mazeSize & "," & Replace(density, ",", ".")

    If fileExists Then
        ' File exists, open for appending
        Open filePath For Append As fileNum
    Else
        ' File doesn't exist, create it and add a header row
        Open filePath For Output As fileNum
        Print #fileNum, "UniqueID,GameDateTime,PlayerMoves,CompletionTime,MazeSize,Density"
    End If

    ' Write the CSV string to the file
    Print #fileNum, csvStr
    
    ' Close the CSV file
    Close fileNum
    
    Debug.Print "Data appended to CSV file: " & filePath
End Sub

Sub AppendDataToJson()
    ' Specify the file path for the JSON file
    Dim filePath As String
    filePath = "C:\Users\corey\Personal Projects\Excel-VBA-Maze-Game\DB\Data.json"
    
    ' Open the JSON file in Append mode or create it if it doesn't exist
    Dim fileNum As Integer
    fileNum = FreeFile
    On Error Resume Next
    Open filePath For Append As fileNum
    If Err.Number <> 0 Then
        ' File doesn't exist, create it and add an empty array
        Open filePath For Output As fileNum
   Print #fileNum, "[" ' Start with an opening bracket for the JSON array
    Close fileNum ' Close the file to reset the pointer
    Open filePath For Append As fileNum ' Now open it for appending
    End If
    On Error GoTo 0
    
    ' Create a JSON string with the data
    Dim jsonStr As String
    jsonStr = "{""UniqueID"":""" & UniqueID & """,""GameDateTime"":""" & _
              Format(startTime, "yyyy-mm-ddThh:mm:ss") & """,""PlayerMoves"":" & _
              moveCount & ",""CompletionTime"":" & gameDuration & ",""MazeSize"":" & _
              mazeSize & ",""Density"":" & Replace(density, ",", ".") & "}"
    
    ' Read the existing JSON from the file, assume it's an array
    Dim fileContent As String
    Close fileNum ' Close the file to reset the pointer
    Open filePath For Input As fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close fileNum
    
    ' Remove the last bracket and append the new data
    If Len(fileContent) > 2 Then ' Check if the array is not empty
        fileContent = Left(fileContent, Len(fileContent) - 1) & "," ' Remove the last bracket and add a comma
    End If
    
    ' Reopen the file for output and rewrite the modified content
    Open filePath For Output As fileNum
    Print #fileNum, fileContent & jsonStr & "]"
    
    ' Close the JSON file
    Close fileNum
    
    Debug.Print "Data appended to JSON file: " & filePath
End Sub

