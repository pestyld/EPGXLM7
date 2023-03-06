*************************************************************************;
* Demo 2.04: Set the Printing Options                                   *;
* NOTE: Execute the libname.sas program if necessary.                   *;
*************************************************************************;


**********;
* Step 1 *;
**********;
******************************************************************;
* Use the REPORT procedure to create a SAS report                *;
******************************************************************;

* Sort the data *;
proc sort data=pg.prdsales_final 
          out=final;
	by Country;
run;

* Create a report *;
proc report data=final
			spanrows;
	* Create BY groups *;
	by Country;
	* Select columns for the report *;
	columns ('Company Sales Overview' Country Year Product Actual);
	* Specify what each column does in the report *;
	define Country / group;
	define Year / group;
	define Product / group;
	define Actual / analysis sum ;
	* Break after each group summarization *;
	break after Country / summarize suppress;
	break after Year / summarize suppress;
run;



**********;
* Step 2 *;
**********;
******************************************************************;
* Create an Excel report with default printing options           *;
******************************************************************;
ods excel file="&outpath/xls02d04_Print_Layout.xlsx"
          options(sheet_name='#byval1'         /*<--- insert the name of a BY-group variable      */
                  suppress_bylines='on');      /*<--- suppress BY lines in the worksheet          */

* Create a report *;
proc report data=final
			spanrows;
	* Create BY groups *;
	by Country;
	* Select columns for the report *;
	columns ('Company Sales Overview' Country Year Product Actual);
	* Specify what each column does in the report *;
	define Country / group;
	define Year / group;
	define Product / group;
	define Actual / analysis sum ;
	* Break after each group summarization *;
	break after Country / summarize suppress;
	break after Year / summarize suppress;
run;

ods excel close;



**********;
* Step 3 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d04_Print_Layout.xlsx)*;
********************************************************************;



**********;
* Step 4 *;
**********;
******************************************************************;
* Modify the print layout                                        *;
******************************************************************;

* Store the current date as a macro variable and use as a header *;
%let currDate = %sysfunc(today(), weekdate.);
%put &=currDate;

ods excel file="&outpath/xls02d04_Print_Layout.xlsx"
          options(sheet_name='#byval1'                                     
                  suppress_bylines='on'                  
                  center_vertical='on'                                     /*<--- center the table vertically    */
                  center_horizontal='on'                                   /*<--- center the table horizontally  */
                  orientation='landscape'                                  /*<--- modify the orientation         */
                  blackandwhite='on'                                       /*<--- change the printing color      */
                  print_footer='Contact Peter at 555-5555 with questions'  /*<--- add footer when printing       */
                  print_header="Report created on &currDate");             /*<--- add header when printing       */


* Create a report *;
proc report data=final
			spanrows;
	* Create BY groups *;
	by Country;
	* Select columns for the report *;
	columns ('Company Sales Overview' Country Year Product Actual);
	* Specify what each column does in the report *;
	define Country / group;
	define Year / group;
	define Product / group;
	define Actual / analysis sum ;
	* Break after each group summarization *;
	break after Country / summarize suppress;
	break after Year / summarize suppress;
run;

ods excel close;



**********;
* Step 5 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d04_Print_Layout.xlsx)*;
********************************************************************;