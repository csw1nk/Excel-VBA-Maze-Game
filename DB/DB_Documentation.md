# Maze Game Data Documentation

## CSV File Generation and Structure
- **Purpose**: Store game data in a structured text format.
- **File Creation**: A new CSV file (`Data.csv`) is created post the first game.
- **Appending Data**: Subsequent game data are appended to this file.
- **Structure**: Header includes `UniqueID, GameDateTime, PlayerMoves, CompletionTime, MazeSize, Density`.

## JSON File Generation and Structure
- **Purpose**: Store game data in a readable format.
- **File Creation**: A new JSON file (`Data.json`) is generated after the first game.
- **Unique Formatting**: First record enclosed within `[ ]`, subsequent records require manual closure.
- **Appending Data**: New game data appended, maintaining array structure.

## Database Integration
- **Setup**: Requires Access database setup with OLE DB Connection and ADODB objects.
- **Error Handling**: Manages no database connection scenarios.
- **Data Fields**: As per original documentation - `GameID`, `GameDateTime`, `PlayerMoves`, `CompletionTime`, etc.

## Gameplay and Data Recording
- Involves navigating a maze, tracking game parameters and player movements.
- Data recorded in CSV, JSON, and Access database (if available) post game completion/reset.

# Maze Game Database Fields Documentation

## GameID

- **Purpose**: To uniquely identify each game session.
- **Data Type**: AutoNumber (automatically increments for each new record).
- **Possible Values**: Automatically generated number.
- **Relation to Game**: Identifies individual game sessions.

## GameDateTime

- **Purpose**: To record when the game was played.
- **Data Type**: Date/Time.
- **Possible Values**: Any valid date and time.
- **Relation to Game**: Allows analysis of game usage over time.

## PlayerMoves

- **Purpose**: To track the number of moves the player makes.
- **Data Type**: Number (Integer).
- **Possible Values**: Any whole number.
- **Relation to Game**: Indicates the player's activity and potential game difficulty.

## CompletionTime

- **Purpose**: To measure how long it takes to complete the game.
- **Data Type**: Number (Long Integer or Double for precision).
- **Possible Values**: Time in seconds.
- **Relation to Game**: Reflects the game's duration and can indicate difficulty or player skill.
