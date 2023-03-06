************************************************************************;
* Demo 2.03: Modifying the Appearance of Tables and Worksheets         *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
******************************************************************;
* Create a default Excel report with titles and footnotes        *;
******************************************************************;

ods excel file="&outpath/xls02d03_Appearance.xlsx";

* Worksheet 1 - Summary Table *;  
title height=14pt "Actual vs Predicted Sales by Country and Year";
footnote "Created on June 18th";
proc print data=pg.country_sales noobs;
	sum Actual Predict DiffFromPredict;
run;
title;


* Worksheet 2 - Visualization *;
title height=14pt "Total Sales by Year";
proc sgplot data=pg.country_sales;
	vbar Year /
		response=Actual
		fillattrs=(color=dodgerblue);
	yaxis display=(nolabel);
run;
title;
footnote;

ods excel close;



**********;
* Step 2 *;
**********;
*******************************************************************;
* Download and open the Excel workbook (xls02d03_Appearance.xlsx) *;
*******************************************************************;



**********;
* Step 3 *;
**********;
******************************************************************;
* Include titles and footnotes in the Excel workbook             *;
******************************************************************;
ods excel file="&outpath/xls02d03_Appearance.xlsx"
          nogtitle nogfootnote                  /*<--- prints title and footnote outside the graph border    */
          options(embedded_titles="on"          /*<----embed titles in the worksheet.                        */
                  embedded_footnotes="on");     /*<----embed footnotes in the worksheet.                     */

* Worksheet 1 - Summary Table *;  
title height=14pt "Actual vs Predicted Sales by Country and Year";
footnote "Created on June 18th";
proc print data=pg.country_sales noobs;
	sum Actual Predict DiffFromPredict;
run;
title;


* Worksheet 2 - Visualization *;                                
title height=14pt "Total Sales by Year";
proc sgplot data=pg.country_sales;
	vbar Year /
		response=Actual
		fillattrs=(color=dodgerblue);
	yaxis display=(nolabel);
run;
title;
footnote;

ods excel close;



**********;
* Step 4 *;
**********;
*******************************************************************;
* Download and open the Excel workbook (xls02d03_Appearance.xlsx) *;
*******************************************************************;



**********;
* Step 5 *;
**********;
******************************************************************;
* Add a filter, modify column widths, and freeze columns and rows*;
******************************************************************;
ods excel file="&outpath/xls02d03_Appearance.xlsx"
          nogtitle nogfootnote                  
          options(embedded_titles="on"          
                  embedded_footnotes="on");     

* Worksheet 1 - Summary Table *;  
ods excel options(autofilter='all'                     /*<--- Add a filter to the table            */                       
                  absolute_column_width='15,10,20,20'  /*<--- Modify the column widths             */
                  frozen_headers='3'                   /*<--- Freezes column headers vertically    */
                  frozen_rowheaders='2');              /*<--- Freezes rows headers horizontally    */


title height=14pt "Actual vs Predicted Sales by Country and Year";
footnote "Created on June 18th";
proc print data=pg.country_sales noobs;
	sum Actual Predict DiffFromPredict;
run;
title;


* Worksheet 2 - Visualization *;    
ods excel options(frozen_headers='off'         /*<--- Remove freeze on column headers       */ 
                  frozen_rowheaders='off');    /*<--- Remove freeze on row headers          */  
                                            
title height=14pt "Total Sales by Year";
proc sgplot data=pg.country_sales;
	vbar Year /
		response=Actual
		fillattrs=(color=dodgerblue);
	yaxis display=(nolabel);
run;
title;
footnote;

ods excel close;



**********;
* Step 6 *;
**********;
*******************************************************************;
* Download and open the Excel workbook (xls02d03_Appearance.xlsx) *;
*******************************************************************;