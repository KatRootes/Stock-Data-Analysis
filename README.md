# Stock-Data-Analysis
Analyze stock market data over several years using VBA scripting.

# 2014 Results

Greatest % Increase:			  	DM	      5581.16%

Greatest % Decrease:	 			 	CBO	     	-95.73%

Greatest Total Volume:				BAC	    	21595474700

# 2015 Results

Greatest % Increase:        	ARR	      491.30%

Greatest % Decrease:	        KMI.W	    -98.59%

Greatest Total Volume:	      BAC	      21277761900

# 2016 Results

Greatest % Increase:	        SD	      11675.00%

Greatest % Decrease:	        DYN.W	    -91.49%

Greatest Total Volume:	      BAC	      27428529600

# Observations

* BAC has the greatest volume all three years

* Greatest % increases can be very large if stock prices start very low.

# Assumptions

* Only 1 year is represented per Worksheet and the Year is the Name of the Worksheet

* Data is sorted by Ticker then by date ascending for the year

#  Notes

* Used scripting to set % and currency values to two decimal places

* Centered data in cells

* Placed grid lines around all cells within tables

* Created headers with a navy background and bold white font

* Read the Worksheet name and placed in the upper left-hand corner of the metrics tables following the formatting of other headers but with a size 14 font

* Used Sub routines to avoid a single long method and to reduce code by calling multiple times.  Also learned how to pass a parameter to a method

* Did not break apart the main loop that performs calculations into multiple methods as I have not figured out how to create objects in VBA yet.
