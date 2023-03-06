************************************************************************;
* Activity 1.03: Exporting a SAS Table to Microsoft Excel              *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


proc contents data=pg.cars;
run;

proc print data=pg.cars(obs=10) label;
run;

* Complete the EXPORT procedure *;
proc export;

run;






















*****************************;
* SOLUTION                  *;
*****************************;
/*
proc contents data=pg.cars;
run;

proc print data=pg.cars(obs=10) label;
run;

proc export data=pg.cars
			dbms=xlsx
			outfile="&outpath/xls01a03.xlsx"
			replace
			label;
	sheet='Raw Data';
run;
*/