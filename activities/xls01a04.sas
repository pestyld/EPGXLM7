************************************************************************;
* Activity 1.04: Connecting to Excel Using the LIBNAME Statement       *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


* Add the LIBNAME statement below *;


proc freq data=pgxl.Asia;
	table Drivetrain;
run;

libname pgxl clear;





















*****************************;
* SOLUTION                  *;
*****************************;
/*
libname pgxl xlsx "&path/data/cars_origin.xlsx";

proc freq data=pgxl.Asia;
	table Drivetrain;
run;

libname pgxl clear;
*/