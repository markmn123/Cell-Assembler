# Cell-Assembler
**V1.1 Update***

NOTE: There seems to be an issue when saving to an excel file that can be attributed to the installation of excel. If you chose the option to save to excel and it just closes out, you may need to repair the excel installation
1. Script actually adds the formulas for the following:
	- SUM for all parallel cells in each series
	- AVG formula for all series cells
	- Max series cell difference formula as a percent
2. Added round function to calculated values to stop long ugly numbers appearing
3. Added LTO battery chemistry

With the formulas written to the sheet you can now sort the columns or even move cells around to cut down on the max series cell difference.
INTERESTING NOTE: After you choose "Yes" when asked if you would like to make another pack at the end, any cells that were used will not be made available for future packs even if you elect not to delete them from capacities.txt.

Python script and .EXE (available in the release section) to logically arrange Lithium cells in a semi-optimized format (must stress semi-optimized)

I have been building lithium battery packs for some time now and have never enjoyed the daunting task of arranging the cells in an orderly fashion so that all series cells are matched and so...the idea for this tool was born.

What it does:

Takes all the cell capacities you wrote down and arranges them in battery packs as efficiently as possible (its not perfect but its not bad either!!)
There are literally no limitations to cell configurations so long as you have that many cells in the capacities.txt file (Ex. 5s20p, 100s4p)

You can select to output to the terminal or save to an Excel file for further arrangement.

It gives some info along with it such as total pack capacity, max capacity difference in % between largest and smallest series cell, cut off, nominal, and fully charged voltage. (Lion and LiFePo4 but can add more chemistries if needed)

It gives you the option to remove the values from your capacities.txt file after you are done so future packs made dont use the same cells. Also the selection script will NOT use the same cell twice ever during the process. Not even for multiple packs. But if you have multiples of the same value it will (obviously)

The basic workflow for this tool to be useful goes as follows:
  1. You need to have an abundance of cells that have already been capacity tested.
  2. Once those cells have been tested you write down the capacity using the program of your choice (notepad, notepad++) (havent tried word but maybe it will work?)
      -the capacities need to be one per line. no spaces after or before the value. also no need to put "mAh" for anything
  3. Once that is done you name that file "capacities.txt" and you save in in the same location as the script.
  4. Run the script. I tried to make it as no brainer as possible.
  5. Arrange your pack and start welding!

This is the first real coding project I have ever worked on so any feedback is appreciated and I am open to feature requests.
I am also providing the full standalone executable in the releases for the cell assembler script that can be run on Windows without installing Python or any additional dependencies.
Please note that when the file is first executed there may be a bit of a delay.
Also when the Excel file option is selected it will output to a .xlsx file. You either need Excel installed or you can upload to Google sheets in order to view the output
