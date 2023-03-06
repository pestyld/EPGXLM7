************************************************************************;
* Demo 2.02: Modifying the Excel Worksheet Output, Names, and Colors   *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
******************************************************************;
* Create an Excel report with multiple procedures                *;
******************************************************************;
ods excel file="&outpath/xls02d02_Worksheets.xlsx";

proc print data=pg.prdsales_final(obs=10);
run;


proc freq data=pg.prdsales_final;
	tables Country State;
run;


proc means data=pg.prdsales_final;
	var Actual Predict;
	class Year;
quit;


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
* Step 2 *;
**********;
*******************************************************************;
* Download and open the Excel workbook (xls02d02_Worksheets.xlsx) *;
*******************************************************************;



**********;
* Step 3 *;
**********;
******************************************************************;
* Manually specify when to create a worksheet                    *;
******************************************************************;

ods excel file="&outpath/xls02d02_Worksheets.xlsx"
		  options(sheet_interval = 'none');          /*<---when a new worksheet is created */

* Worksheet 1 *;
proc print data=pg.prdsales_final(obs=10);
run;


* Worksheet 2 *;
ods excel options(sheet_interval = 'now');           /*<---when a new worksheet is created */

proc freq data=pg.prdsales_final;
	tables Country State;
run;


proc means data=pg.prdsales_final;
	var Actual Predict;
	class Year;
quit;


* Worksheet 3 *;
ods excel options(sheet_interval = 'now');           /*<---when a new worksheet is created */

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
* Step 4 *;
**********;
*******************************************************************;
* Download and open the Excel workbook (xls02d02_Worksheets.xlsx) *;
*******************************************************************;



**********;
* Step 5 *;
**********;
******************************************************************;
* Modify Worksheet Names and Tab Color                           *;
******************************************************************;

ods noproctitle;                                       /*<---remove procedure titles */

ods excel file="&outpath/xls02d02_Worksheets.xlsx"
		  options(sheet_interval = 'none'
		  	      sheet_name = 'Data Preview'          /*<---worksheet name */
		  	      tab_color = 'red');                  /*<---tab color      */

* Worksheet 1 *;
proc print data=pg.prdsales_final(obs=10);
run;


* Worksheet 2 *;
ods excel options(sheet_interval = 'now'
				  sheet_name = 'Analysis'             /*<---worksheet name */
				  tab_color = 'blue');                /*<---tab color      */

proc freq data=pg.prdsales_final;
	tables Country State;
run;

proc means data=pg.prdsales_final;
	var Actual Predict;
	class Year;
quit;


* Worksheet 3 *;
ods excel options(sheet_interval = 'now'
				  sheet_name = 'Visualization'        /*<---worksheet name */
				  tab_color = 'Green');               /*<---tab color      */
 
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

ods proctitle;                                        /*<---restore procedure titles */



***********;
* Step 6  *;
***********;
*******************************************************************;
* Download and open the Excel workbook (xls02d02_Worksheets.xlsx) *;
*******************************************************************;