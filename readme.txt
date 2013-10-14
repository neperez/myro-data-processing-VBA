Experiment: 
	Dynamic measurment of Scribbler robot #5's proximity signal, recording 3 sets of multiple runs at 3 different speeds.  
Procedure:
	The robot was placed 1 inch from an obstacle and commanded to move in a backwards fashion to stop at the 20 inch mark.  The target travel distance was 19 inches. 3 sets of 30 cycles were performed for speeds -.2,-.6, 	and -1.  The data was written to a text file and then transfered to this excel workbook.    
		
	Starting point: 1 inch from a large geometric surface (wall)
	Ending point: 20 inch mark from a large geometric surface.
	
Excel Workbook:
	The sheet names start with the speed traveled, the number after the underscore is the number of which of the three runs it was.  3 speeds, 3 sets of 30 cycles each.  The speeds  are negative values for backwards 	movement.  Each sheet contains the raw signal values collected for each run.  The current time was recorded for each signal, so we know at what time interval the signals were recorded.  The actual distance travelled at 	each interval was calulated as well.
	
	The first 9 work sheets (excluding KEY) contains tables and line graphs of 
		-unprocessed signal and distance table
		-a linear version of the distance scale with interpolated (linear) signal values
		-Windowed averages of the interpolated signal.
	for each of the 30 runs/cycles at each speed.
	
	The last 3 sheets are for determining the average of averages for each speed.  90 values (3 sets of 30 cycles) of each speed were taken and averaged.  Graphs of the averages and standard deviation are included.

VBA Code:   
	AverageAndInterpolate:  Takes the unprocessed values and interpolates the signal so it will fit a linear scale.  The processed signal values are then averaged using a multi-value window averaging method.
	GetLongestDistance:  Helped determine the longest distance travelled at each speed to help set up the linear distance scale on the averages sheets (last 3 sheets in workbook).
	PrepareAndAverage: I had to change the windowed average value to another number, this sub reduced my time greatly.
	TransferAverage:  Transfers the averages of all cycles to their respective averaging sheet.
