Script commands

note: view this in TERMINAL or other Fixed font

note: all parameters must be specified, if not needed, just input a 0

EXAMPLE:

not valid: print|debug|my text|
valid:     print|debug|my text|0|0|0|0

parameter seperator: |

# comment
 #comment
		#comment

#PRINT:  print|debug:custom:title|string|x|y|time|order
#WAITTIME: waittime|time
#SPAWN: spawn|id|x|y
#HELPER helper|weapon:health:score:cash|cashval|scoreval|healthval|weaponid|x|y
#OBJECT object|bmp_file|x|y|spdx|spdy|name|stretch[0:1]|stchW|stchH

NOTE: other commands exist in the code but are not implemented yet (laziness on my part)
but you can make a fully functional and fun level with these commands, and you can easilly add your own!




load a script file:
note: all script files must be located in the \SCRIPTS\ directory

load|script_file





spawn an enemy:
note: enemy_id must be a valid enemy id (set these in the program code [LoadEnemies sub)

default enemies ids are 1 and 2

spawn|enemy_id|x_location|y_location





Print text to the screen:
note: the first parameter can be one of three:
	debug:
          prints in order, auto Y, X and Y is optional
	custom:
	  prints exact given X,Y
	Title
          prints large text, if X is 0, it centers X, if Y is 0, it centers Y
time_ms is the time to display the text in milliseconds after it is printed
order is the Y order if using debug type

print|debug:custom:title|string|x_location|y_location|time_ms|order




Pause script execution:
note: time_ms is the amount of time to pause in milliseconds

waittime|time_ms




spawn a helper object (goodie)
note: there are 4 types of helpers (currently):
	weapon
		spawns a weapon, you MUST specify a weapon_id
		default weapon ids are 11,12,13,14,21,22,23,71
		define weapon ids in the LoadGuns sub
	health
		health_value must be non_zero
	score
		score_value must be non-zero
	cash	
		cash_value must be non_zero
if X is 0, the engine will spawn at random X point
if Y is 0, the engine will spawn at random Y point

helper|weapon:health:score:cash|cash_value|score_value|health_value|weapon_id|x|y





spawn an external moving or static bitmap:

note: The mars object in the main.spt file is loaded this way...

bmp_file is the name of any bitmap file in the \GFX\ folder

speed_x is the movement on the x coordinate
speed_y is the movement on the y coordinate

stretch[0:1] can be 0 or 1, 1 designates that this object should be stretched
to match the stretch_w and stretch_H dimensions when drawn.

NOTE: NAME must not be null (""), choose any name!!!

object|bmp_file|x|y|speed_x|seepd_y|name|stretch[0:1]|stretch_W|stretch_H





***************************************************************
Open the Main.spt file in notepad for examples of all commands!
***************************************************************
