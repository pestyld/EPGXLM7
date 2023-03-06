************************************************************************;
* Demo 2.01: Using the ODS EXCEL Statement to Create Excel Reports     *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
******************************************************************;
* Create a simple Microsoft Excel report                         *;
******************************************************************;

ods excel file="&outpath/xls02d01a_ODS_Excel_Basics.xlsx";

proc print data=pg.prdsales_final(obs=10);
run;

ods excel close;



**********;
* Step 2 *;
**********;
**************************************************************************;
* Download and open the Excel workbook (xls02d01a_ODS_Excel_Basics.xlsx) *;
**************************************************************************;



**********;
* Step 3 *;
**********;
******************************************************************;
* Close the SAS results                                          *;
******************************************************************;

* Close the default SAS output destination *;
ods _all_ close;

ods excel file="&outpath/xls02d01b_CloseDefault.xlsx";

proc print data=pg.prdsales_final(obs=10);
run;

ods excel close;

*********************************************************;
* This is required for the SAS Windowing Environment    *;
*********************************************************;
* Run the ODS HTML statement to turn on the default     *; 
* SAS destination.                                      *;
*********************************************************;
/* ods html; */



**********;
* Step 4 *;
**********;
******************************************************************;
* Modify the Excel workbook properties                           *;
******************************************************************;
ods excel file="&outpath/xls02d01c_ExcelProperties.xlsx"
		  title='This is the title of the Excel workbook'
		  author='Peter S'
		  category='Sales Overview for the Team'
		  status='Updated Version'
		  comments='These are the comments I want to add to the workbook.'
		  keywords='Sales, Revenue, Predict, Actual';

proc print data=pg.prdsales_final(obs=10);
run;

ods excel close;



**********;
* Step 5 *;
**********;
************************************************************************;
* Download and open the Excel workbook (xls02d01c_ExcelProperties.xlsx)*;
************************************************************************;



**********;
* Step 6 *;
**********;
******************************************************************;
* Create a dynamic Excel file name                               *;
******************************************************************;

*****;
* a *;
*****;
%let currentDate = %sysfunc(today(),yymmdd10.);
%put &=currentDate;


*****;
* b *;
*****;
ods excel file="&outpath/xls02d01d_&currentDate..xlsx";

proc print data=pg.prdsales_final(obs=10);
run;

ods excel close;



**********;
* Step 7 *;
**********;
******************************************************************;
*  Use multiple procedures with the ODS EXCEL destination        *;
******************************************************************;
ods excel file="&outpath/xls02d01d_&currentDate..xlsx";

proc print data=pg.prdsales_final(obs=10);
run;


proc freq data=pg.prdsales_final;
	tables Country State;
run;


proc means data=pg.prdsales_final;
	var Actual Predict;
	class Year;
run;


title height=16pt justify=left color=gray "Total Sales by Year";
proc sgplot data=pg.country_sales;
	vbar Year /
		response=Actual
		fillattrs=(color=dodgerblue);
	format Actual dollar16.;
	yaxis display=(nolabel);
run;
title;


ods excel close;



**********;
* Step 8 *;
**********;
************************************************************************;
* Download and open the Excel workbook (xls02d01d_<currentDate>.xlsx)  *;
************************************************************************;