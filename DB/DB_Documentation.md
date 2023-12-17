
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
