README.md
Matthew Idle
University of Minnesota Data Analytics and Visualization Bootcamp 
Challenge 2

A for-each loop executes this script in each worksheet sequentially when an empty
cell is detected in the ticker column (1), indicating that sheet is finished.



Within the for each loop are a series of conditionals. The first conditional if statement begins when the eval_started variable is set to false, this value indicates a new ticker has been detected. Next the script reads in the ticker, open, and close values and begins totaling the daily volume.

if the previous condition wasn't met then an ifelse statement is evaluated comparing the current ticker value with the ticker value in the cell below it. if they are not equal, the final volume is added to the yearly total and the final closing value is read. At this point, the opening value is subtracted from the closing value to give us the yearly change, the yearly change change is calculated by dividing the closing value by the opening value and subtracted one. Next a series of conditional if statements are executed to determine the greatest increase, decrease, and greatest total volume values. The stored greatest/least value is compared with the current value, if the current value is greater, or lesser in one case, then it replaces it, otherwise it just rewrites the value to itself. Next the values for display are writtten to columns with the row being indexed by variable, "j". Next the yearly change, percent yearly change, and the greatest increase/decrease are formatted. Lastly data row variable index is incremented, and evaluation_started variables are re-initialized.


The else statement just adds the daily volume to the daily volume total.


