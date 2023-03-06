************************************************************************;
* Activity 2.04: Applying Styles to an Excel Report                    *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


ods excel file="&outpath/xls02a04.xlsx" 
          style=illuminate;
		  	      
* Print the data *;
proc print data=pg.cars(obs=10) noobs label;
run;

* Create a visualization *;
title height=14pt justify=left "Average City Miles Per Gallon (MPG) by Car Make";
proc sgplot data=pg.cars;
	vbar Origin /
		response=MPG_City
		stat=mean
		categoryorder=respdesc;
run;
title;

ods excel close;
























*****************************;
* SOLUTION                  *;
*****************************;
/*
ods excel file="&outpath/xls02a04.xlsx"
          style=Analysis;
		  	      
* Print the data *;
proc print data=pg.cars(obs=10) noobs label;
run;


* Create a visualization *;
title height=14pt justify=left "Average City Miles Per Gallon (MPG) by Car Make";
proc sgplot data=pg.cars;
	vbar Make /
		response=MPG_City
		stat=mean
		categoryorder=respdesc;
run;
title;

ods excel close;
*/