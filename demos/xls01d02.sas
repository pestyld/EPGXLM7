************************************************************************;
* Demo 1.02: Importing Excel Data into SAS Using PROC IMPORT           *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
****************************************************************************;
* Download and open the Microsoft Excel workbook (prdsales_countries.xlsx) *;
****************************************************************************;



**********;
* Step 2 *;
**********;
******************************************************************;
* Import the Excel Workbook as a SAS Table                       *;
******************************************************************;

* Best practice is to specify the VALIDVARNAME=V7 option to convert column names to valid SAS column names *;                                                   *;
options validvarname=v7; 

* Import the Excel workbook *;
proc import datafile="&path/data/prdsales_countries.xlsx" 
			dbms=xlsx 
		    out=work.prdsales_usa
			replace;
run;



**********;
* Step 3 *;
**********;
*********************************************************************************;
* Use the SHEET= statement option to import a specific worksheet as a SAS table *;
*********************************************************************************;

options validvarname=v7;
proc import datafile="&path/data/prdsales_countries.xlsx" 
			dbms=xlsx 
			out=work.prdsales_usa
			replace;
	sheet='USA';
run;

* Preview the new SAS table *;
proc contents data=work.prdsales_usa;
run;

proc print data=work.prdsales_usa(obs=10);
run;



**********;
* Step 4 *;
**********;
******************************************************************;
* Use data set options with the IMPORT procedure                 *;
******************************************************************;

options validvarname=v7; 
proc import datafile="&path/data/prdsales_countries.xlsx"
			dbms=xlsx 
			out=work.prdsales_ds_options(where=(State in ('California', 'New York')))
			replace;
	sheet='USA';
run;

* View the total number of rows in the new table *;
proc contents data=work.prdsales_ds_options;
run;

* View the frequency distribution of the State column *;
proc freq data=work.prdsales_ds_options;
	tables State;
run;



**********;
* Step 5 *;
**********;
******************************************************************;
* Import a specific range of data from an Excel worksheet        *;
******************************************************************;

options validvarname=v7; 
proc import datafile="&path/data/prdsales_countries.xlsx"
			dbms=xlsx 
			out=work.prdsales_usa_range
			replace;
	sheet='USA';
	range='A1:D10'; *<----Specify a range or named range *;
run;

* Preview the new table *;
proc print data=work.prdsales_usa_range;
run;