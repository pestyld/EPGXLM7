************************************************************************;
* Demo 1.04: Importing and Exporting Data Using the LIBNAME Engine     *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
******************************************************************;
* Make a LIBNAME connection to the Excel workbook                *;
******************************************************************;
options validvarname=v7;
libname pgxl xlsx "&path/data/prdsales_countries.xlsx" access=readonly;



**********;
* Step 2 *;
**********;
******************************************************************;
* View the SAS library                                           *;
******************************************************************;
 
* View Available Tables in the pgxl Library programmatically *;
options validvarname=v7;  
proc contents data=pgxl._all_;
run;

******************************************************************;
* View Available tables in the pgxl library manually             *;
******************************************************************;
* Go to the navigation bar and select libraries, find pgxl.      *;
* Notice all 3 worksheets are shown as SAS data sets             *;
******************************************************************;



**********;
* Step 3 *;
**********;
******************************************************************;
* Work with the worksheets through the SAS library               *;
******************************************************************;
options validvarname=v7;
proc print data=pgxl.mexico(obs=10);
run;

proc print data=pgxl.canada(obs=10);
run;

proc print data=pgxl.usa(obs=10);
run;



**********;
* Step 4 *;
**********;
************************************************************************;
* Use the DATA Step to concatenate the Excel worksheets as a SAS table *;
************************************************************************;
options validvarname=v7;

data work.prd_all_libname;
	length Country $12. State $40.; /*<---Specify the length of the character columns */
	set pgxl.usa
		pgxl.canada
		pgxl.mexico;
		
	* Create a calculated column *;
	TotalDiff = Actual - Predict;
	
	* Format columns *;
	format TotalDiff Actual Predict dollar16.2
	       Manager_ID z10. 
		   _CHARACTER_;  /*<---Remove the formats for all character columns */
		  
	label Manager_ID = 'Manager ID';
run;

proc contents data=work.prd_all_libname;
run;

proc print data=work.prd_all_libname(obs=10) label;
run;

proc freq data=work.prd_all_libname;
	tables Country State;
run;



**********;
* Step 5 *;
**********;
******************************************************************;
* Write out the SAS table as a new Excel workbook                *;
******************************************************************;

* Create a new Excel workbook named prdsales_libname_engine.xlsx *;
libname outxl xlsx "&outpath/xls01d04.xlsx";

data outxl.prdsales_all;
	set work.prd_all_libname;
run;

* Clear connection *;
libname outxl clear;

******************************************************************;
* NOTE: Step 4 and 5 can be done in a single step                *;
******************************************************************;



**********;
* Step 6 *;
**********;
*******************************************************************;
* Download and open the Excel workbook (xls01d04.xlsx)            *;
*******************************************************************;



**********;
* Step 7 *;
**********;
******************************************************************;
* Add three worksheets to the existing Excel workbook            *;
******************************************************************;

* Create a library reference to the Excel workbook *;
libname outxl xlsx "&outpath/xls01d04.xlsx";

* Worksheet USA *;
data outxl.USA;
	set outxl.prdsales_all;
	where Country = 'USA';
	USDollarsActual=Actual;
	USDollarsPredict=Predict;
run;

* Worksheet Canada *;
data outxl.Canada;
	set outxl.prdsales_all;
	where Country = 'Canada';
	Conversion=1.26;
	CanadianDollarsActual = Actual * Conversion;
	CanadianDollarsPredict = Predict * Conversion;
	drop Conversion;
run;

* Worksheet Mexico *;
data outxl.Mexico;
	set outxl.prdsales_all;
	where Country = 'Mexico';
	Conversion=19.58;
	PesosActual = Actual * Conversion;
	PesosPredict = Predict * Conversion;
	drop Conversion;
run;

* Clear connection *;
libname outxl clear;


**********;
* Step 8 *;
**********;
*******************************************************************;
* Download and open the Excel workbook (xls01d04.xlsx)            *;
*******************************************************************;



**********;
* Step 9 *;
**********;
*******************************************************************;
* Delete the .bak file                                            *;
*******************************************************************;

*****;
* a *;
*****;
* Reference the Excel workbook .bak file *;
filename f_bak "&outpath/xls01d04.xlsx.bak";

* The FEXISTS function indicates if the file exists *;
data work.files;
	f_bak_exists = fexist('f_bak');
run;


*****;
* b *;
*****;
* If the files exists, delete using the FDELETE function *;
data _null_;
	if fexist('f_bak') = 1 then do;
		f_bak_exists = fdelete('f_bak');
		put 'NOTE: The file was deleted';
	end;
run;