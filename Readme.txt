# VBA-challenge

VBA scripting to analyze generated stock market data

# Declare variables
	Declare all variables to be used in this code so we have items to move value through them.
 
# Main Loop to go through all the sheets in workbook


	Create the first loop that will going to go through all the sheets and output data to summary table.

	After loop the summary table should start from row 2.
	Here also the open year variable will be placed to pick the first value from year open.

# Sub loop 1 to find the values and moving it into the variables

	This loops is traversing through sheets and based on the ticker will pick the open year and close year.

# Start with an if statement

	The close year variable will pick the last value from year close making sure all values are coming in based for that first ticker.

	All calculations will be done here By each ticker 

	Quarterly Change

	Percentage
			
	Total Stock volume 

# The Values each ticker will be output after calculations by their variables in each row.
	output variables


# Format columns colors based on values in column K that is Quarterly change.
	 Use of if statement if the value is negative color red and if negative then green.

# End of Sub loop 1 


# Calculate three categories the Greatest Percent increase, Greatest Percent Decrease & Greatest Volume 

	
    	Sub Loop 2 for going through the summary table to find the values of three variables 
'
 		If statements for each category

	End of Sub loop 2

	Output the Values by the variables


#End of Main loop






