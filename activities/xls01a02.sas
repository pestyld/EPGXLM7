************************************************************************;
* Activity 1.02: Importing an Excel Worksheet into SAS                 *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;

proc import ;

run;

proc contents data=work.Europe;
run;























*****************************;
* SOLUTION                  *;
*****************************;
/*
proc import datafile="&path/data/cars_origin.xlsx"
			dbms=xlsx
			out=work.Europe
			replace;
	sheet='Europe';
run;

proc contents data=work.Europe;
run;
*/