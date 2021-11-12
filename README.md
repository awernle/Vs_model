# Vs_model

7/7/2021
Purpose:
A python script to automate the creation of Vs30 model figures in Excel for multiple rst files from Seisimager.

Setup:
In order to run this script, User should copy this script into a folder with Excel doc xxxxx_SCHOOL_profile, and final model .rst (you can have multiple rst files). User should also change variables that are below “#...CHANGE THESE VARIABLES” (line 103) in script. This is a straightforward script to run and does not require much setup.

Output:
An excel file with data from rst and plotted final model output to the folder with the rst files AND the designated folder path—you should edit this in the script if you prefer one way or the other.

Potential Errors:
Script may not output excel file to its designated path if the path folder is named incorrectly (lines 176..).
Script may not work if rst file is not saved with modeled picks.

Potential Improvements:
If we want to directly edit the final velocity model in excel with the script, we would need to create the figure from scratch within the script using openpyxl module, however this is time consuming and so we only input rst data into existing excel plot.
