# Goal Supremacy betting strategy testing #


## Summary ##

This is a program I wrote to test a standart Goal Supremacy betting strategy ([exact rules found in this pdf](http://www.football-data.co.uk/ratings.pdf)), by using more than 72.000 football game data entries from [football-data.co.uk](http://www.football-data.co.uk/downloadm.php).

Parts of the scrips can be used to manipulate the data for various different betting strategies.

## How to use ##

1. Run *xlsx_preparation.py* (change filenames to yours). The script removes all unwanted columns and converts .xls to .xlsx, so it can be accessed by the next script.
2. Run *suoremacy_calc.py* (change filenames to yours). It loads all FIXED .xlsx files and does the following:
	1. Calculates Team History - the last 6 game history of each team at each point in time.
	2. Calculates the **Goal Supremacy** for each game.
	3. Merges everthing (except games with no *6-game-history*) on a single, large, .xlsx file.
3. Run *profits_calc.py* to simulate a betting strategy that bets only on **value bets** (check the prementioned pdf for details) and prints out win ratio, profit points and yield.

Sample output:

![goal supremacy testing output](https://i.imgur.com/o5RRPB1.jpg)

## Changelog ##

### Version 1.0 ###

* All the above. No specification needed.
