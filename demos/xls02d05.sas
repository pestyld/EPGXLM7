************************************************************************;
* Demo 2.05: Applying Styles to the Excel Report                       *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
******************************************************************;
* Create a simple Excel report using the default style           *;
******************************************************************;
ods excel file="&outpath/xls02d05a_Styles.xlsx"                
          options(embedded_titles="on"          
                  embedded_footnotes="on");  

* Worksheet 1 - Tabular  Report *;  
title height=14pt "Actual vs Predicted Sales by Country and Year";
footnote "Created on June 18th";
proc print data=pg.country_sales noobs;
	sum Actual Predict;
run;
title;


* Worksheet 2 - Visualization *;
title height=14pt "Total Sales by Year";
proc sgplot data=pg.country_sales;
	vbar Year /
		response=Actual;
run;
title;
footnote;

ods excel close;



**********;
* Step 2 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d05a_Styles.xlsx)     *;
********************************************************************;



**********;
* Step 3 *;
**********;
******************************************************************;
* View available style templates                                 *;
******************************************************************;
proc template;
	list styles;
run;



**********;
* Step 4 *;
**********;
******************************************************************;
* Create an Excel report using a specified style                 *;
******************************************************************;
ods excel file="&outpath/xls02d05a_Styles.xlsx"                
          options(embedded_titles="on"          
                  embedded_footnotes="on")  
          style=Daisy;                          /*<---- apply style */

* Worksheet 1 - Tabular  Report *;   
title height=14pt "Actual vs Predicted Sales by Country and Year";
footnote "Created on June 18th";
proc print data=pg.country_sales noobs;
	sum Actual Predict;
run;
title;


* Worksheet 2 - Visualization *;
title height=14pt "Total Sales by Year";
proc sgplot data=pg.country_sales;
	vbar Year /
		response=Actual;
run;
title;
footnote;

ods excel close;



**********;
* Step 5 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d05a_Styles.xlsx)     *;
********************************************************************;



**********;
* Step 6 *;
**********;
******************************************************************;
* Modify the style of a procedure                                *;
******************************************************************;

* Set macro variables as the colors to use *;
%let darkBlue=cx04304b;
%let darkGray=cx768396;
%let lightGray=lightgray;
%let offWhite=whitesmoke;


ods excel file="&outpath/xls02d05b_ProcStyle.xlsx"    
          options(embedded_titles='on'
                  sheet_interval='none');


/* Default Style */
title "Actual vs Predict Sales by Year and Country";
proc print data=pg.country_sales label;
	id Country;
	var Year Actual Predict DiffFromPredict;
	sum Actual Predict DiffFromPredict;
run;
title;


/* Modified Style */
title height=14pt justify=left color=&darkBlue "Actual vs Predict Sales by Year and Country";
proc print data=pg.country_sales 
           label
           style(header obs obsheader)=[color=&offWhite 
                                        backgroundcolor=&darkBlue 
                                        fontsize=11pt]
           style(grandtotal)=[backgroundcolor=&lightgray];                                                    
	id Country / style(data)=[textalign=right];
	var Year Actual Predict DiffFromPredict / style(data)=[backgroundcolor=&offWhite 
	                                                       color=&darkBlue 
	                                                       fontsize=9pt] ;
	sum Actual Predict DiffFromPredict / style(grandtotal)=[fontsize=12pt 
	                                                        fontweight=bold
	                                                        color=&darkBlue];
run;
title;

ods excel close;



**********;
* Step 7 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d05b_ProcStyle.xlsx)  *;
********************************************************************;



**********;
* Step 8 *;
**********;
******************************************************************;
* Add conditional formatting to highlight certain values         *;
******************************************************************;

*****;
* a *;
*****;
%let modernRed=CXFFA0A0;
%let modernGreen=CX99FFBB;

* Create the Formats *;
proc format;
	* Set color for negative and positive values *;
	value ColorFmt
		low - 0 = "&modernRed"
		0 <- high = "&modernGreen";
		
	* Bold negative values *;
	value WeightFmt
		low - 0 = "bold";
quit;


*****;
* b *;
*****;
* Apply the Format in the Style Option *;
ods excel file="&outpath/xls02d05c_Formatting.xlsx";
                  
proc print data=pg.country_sales label;
	var Country Year Actual Predict;
	var DiffFromPredict / style(data)=[backgroundcolor=ColorFmt.           
	                                   fontweight=WeightFmt.];            
	sum DiffFromPredict / style(grandtotal)=[backgroundcolor=ColorFmt.];  
run;

ods excel close;



**********;
* Step 9 *;
**********;
********************************************************************;
* Download and open the Excel workbook (xls02d05c_Formatting.xlsx) *;
********************************************************************;