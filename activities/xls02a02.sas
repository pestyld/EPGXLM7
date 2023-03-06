************************************************************************;
* Activity 2.02: Modifying Output to Excel Worksheets                  *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


ods excel file="&outpath/xls02a02.xlsx";
		  	      
* Sort the data *;
proc sort data=pg.cars 
		  out=car_sorted;
	by Origin;
run;

* Print Data by Groups *;
proc print data=car_sorted noobs label;
	by Origin;
run;

ods excel close;
























*****************************;
* SOLUTION                  *;
*****************************;
/*
ods excel file="&outpath/xls02a02.xlsx"
		  options(sheet_interval="bygroups"
		  	      suppress_bylines='on'
		  	      sheet_label='Origin');
		  	      
* Sort the data *;
proc sort data=pg.cars 
		  out=car_sorted;
	by Origin;
run;

* Print Data by Groups *;
proc print data=car_sorted noobs label;
	by Origin;
run;

ods excel close;
*/