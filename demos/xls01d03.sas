************************************************************************;
* Demo 1.03: Exporting SAS Data into Excel Using PROC EXPORT           *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


***********;
* Step 1  *;
***********;
******************************************************************;
* Preview the SAS table                                          *;
******************************************************************;
proc contents data=pg.prdsales_final;
run;

proc print data=pg.prdsales_final(obs=10);
run;



**********;
* Step 2 *;
**********;
******************************************************************;
* Export the SAS table to Excel                                  *;
******************************************************************;
proc export data=pg.prdsales_final
			dbms=xlsx 
			outfile="&outpath/xls01d03_export.xlsx"
			replace;
	sheet='Raw Data';
run;



**********;
* Step 3 *;
**********;
******************************************************************;
* Download and open the Excel workbook (xls01d03_export.xlsx)    *;
******************************************************************;



**********;
* Step 4 *;
**********;
******************************************************************;
* Apply column labels to the Excel workbook                      *;
******************************************************************;

proc export data=pg.prdsales_final 
			dbms=xlsx 
			outfile="&outpath/xls01d03_export.xlsx"
			replace
			label;         *<-------Apply column labels *;
	sheet='Column Labels'; *<-------Create a new worksheet in the existing workbook *;
run;



**********;
* Step 5 *;
**********;
******************************************************************;
* Download and open the Excel workbook (xls01d03_export.xlsx)    *;
******************************************************************;



**********;
* Step 6 *;
**********;
******************************************************************;
* Preserve SAS formats                                           *;
******************************************************************;

data work.prdsales_format;
	set pg.prdsales_final;
	Manager_ID_Format = put(Manager_ID,z10.);  *<-----Create a new column and apply the format *;
run;

proc export data=work.prdsales_format
			dbms=xlsx
			outfile="&outpath/xls01d03_export.xlsx"
			replace;
	sheet='Add Format';
run;



**********;
* Step 7 *;
**********;
******************************************************************;
* Download and open the Excel workbook (xls01d03_export.xlsx)    *;
******************************************************************;