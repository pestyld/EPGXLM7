************************************************************************;
* Demo 2.06: Specifying Excel Formats in a SAS Program                 *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
******************************************************************;
* Create a sample table                                          *;
******************************************************************;
data sampledates;
    SASDate = mdy(11,3,2020);
    DateFmt = SASDate;
    format DateFmt date9.;
run;

ods select variables;
proc contents data=sampledates;
run;

proc print data=sampledates noobs;
run;



**********;
* Step 2 *;
**********;
******************************************************************;
* Specify Excel date formats in a SAS procedure                  *;
******************************************************************;
ods excel file="&outpath/xls02d06_ExcelFormats.xlsx"
          options(sheet_interval='none' 
                  embedded_titles='on'
                  absolute_column_width='10,10,20,10,10');

* Specifying Excel formats correctly*;
title "Specifying an Excel format on a column with a SAS date format";
proc print data=sampledates noobs;
	var DateFmt;                 
	var DateFmt / style(data) = {tagattr='format:ddmmyyyy'};
	var DateFmt / style(data) = {tagattr='format:dddd-mm-yyyy'};
	var DateFmt / style(data) = {tagattr='format:yyyy'};
	var DateFmt / style(data) = {tagattr='format:mmmm'};
	var DateFmt / style(data) = {tagattr='format:dd'};
run;

* Specifying Excel formats incorrectly *;
title "Specifying an Excel format on a column without a SAS date format";
proc print data=sampledates noobs;
	var SASDate;
	var SASDate / style(data) = {tagattr='format:ddmmyyyy'};
	var SASDate / style(data) = {tagattr='format:dddd-mm-yyyy'};
	var SASDate / style(data) = {tagattr='format:yyyy'};
	var SASDate / style(data) = {tagattr='format:mmmm'};
	var SASDate / style(data) = {tagattr='format:dd'};
run;

ods excel close;


***************************************************;
* EXCEL DATE FORMAT TIPS                          *;
***************************************************;
* TO DISPLAY                              : USE   *;
****************************************************
* Months as 1–12                          : m     *;
* Months as 01–12                         : mm    *;
* Months as Jan–Dec                       : mmm   *;
* Months as January–December              : mmmm  *;
* Months as the first letter of the month : mmmmm *;
* Days as 1–31                            : d     *;
* Days as 01–31                           : dd    *;
* Days as Sun–Sat                         : ddd   *;
* Days as Sunday–Saturday                 : dddd  *;
* Years as 00–99                          : yy    *;
* Years as 1900–9999                      : yyyy  *;
***************************************************;



**********;
* Step 3 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d06_ExcelFormats.xlsx)*;
********************************************************************;



**********;
* Step 4 *;
**********;
******************************************************************;
* Specify other Excel formats                                    *;
******************************************************************;

*****;
* a *;
*****;
* Create a sample table *;
data samplenumbers;
    Default1 = .1234;
    'xlFmt: #.00'n = Default1;
    'xlFmt: 0.0#%'n = Default1;
    'xlFmt: 0.00\%'n = Default1;
    Default2 = 12345.678;
    'xlFmt: 0.###'n = Default2;   
    'xlFmt: #.0'n = Default2;
	'xlFmt: $#,###.00'n = Default2;    
run;
proc print data=samplenumbers;
run;


*****;
* b *;
*****;
* Create an Excel report with Excel formats *;
ods excel file="&outpath/xls02d06_ExcelFormats.xlsx";
          
proc print data=samplenumbers split=" " noobs;
    var Default1;
    var 'xlFmt: #.00'n / style(data) = {tagattr='format:#.00'};    
    var 'xlFmt: 0.0#%'n / style(data) = {tagattr='format:0.0#%'};
    var 'xlFmt: 0.00\%'n / style(data) = {tagattr='format:0.00\%'};
    var Default2;
    var 'xlFmt: 0.###'n / style(data) = {tagattr='format:0.###'};
    var 'xlFmt: #.0'n / style(data) = {tagattr='format:#.0'};
	var 'xlFmt: $#,###.00'n / style(data) = {tagattr='format:$#,###.00'};
run;

ods excel close;

**************************************************************************************************;
* EXCEL NUMERIC FORMAT TIPS                                                                      *;
**************************************************************************************************;
* - The 0 placeholder displays insignificant zeros if a number has fewer digits than the format. *;
* - The # placeholder does not display zeros when the number has fewer digits than the format.   *;
* - If a number has more digits to the right of the decimal point than there are placeholders,   *;
*   the number rounds.                                                                           *;
* - If a number has more digits to the left of the decimal point than there are placeholders,    *;
*   the digits are shown.                                                                        *;
* - Special characters such as the dollar sign and comma can be included in the format.          *;
* - Including a percent sign with no leading backslash in a format will multiply the number by   *;
*   100 and display the percent sign.                                                            *;
* - Including a percent sign with a leading backslash in a format will display the percent sign  *;
*   and not multiply the number by 100 .                                                         *;
**************************************************************************************************;



**********;
* Step 5 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d06_ExcelFormats.xlsx)*;
********************************************************************;



**********;
* Step 6 *;
**********;
******************************************************************;
* Add Excel formulas                                             *;
******************************************************************;

*****;
* a *;
*****;
* Create a SAS table *;
data work.country_sales;
	set pg.country_sales;
	CalcExcelColumn = .;          * <------ Create a dummy column to add the Excel formula to *;
	drop ForecastAccuracy
	     DiffFromPredict;
run; 


*****;
* b *;
*****;
* Add an Excel formula *;
ods excel file="&outpath/xls02d06_ExcelFormats.xlsx";

proc print data=work.country_sales;
	var Country Year Actual Predict;
	var CalcExcelColumn / style(data)={tagattr='formula:=RC[-2] - RC[-1]'};
run;

ods excel close;


*************************************************************************;
* EXCEL FORMULA FORMAT TIPS                                             *;
*************************************************************************;
* - The RC value corresponds to the cell relative to the current cell.  *;
* - For example, RC[-1] means 1 cell to the left of the current cell.   *;
* - Any valid Excel formula can be used.                                *;
*************************************************************************;



**********;
* Step 7 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d06_ExcelFormats.xlsx)*;
********************************************************************;