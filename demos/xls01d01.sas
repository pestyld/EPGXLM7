************************************************************************;
* Demo 1.01: Integrating SAS and Microsoft Excel                       *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********;
* Step 1 *;
**********;
*********************************************************;
* Open, review, and execute the libname.sas program     *;
*********************************************************;



**********;
* Step 2 *;
**********;
*********************************************************;
* Download and open the Microsoft Excel workbook        *;
*********************************************************;
* Download and open the prdsales_countries.xlsx         *;
* Microsoft Excel workbook from the data folder. View   *;
* the Excel workbook and confirm that it contains three *; 
* worksheets named USA, Canada, and Mexico.             *; 
*********************************************************;



**********;
* Step 3 *;
**********;
*********************************************************;
* Import the data from Excel and prepare a SAS table    *;
*********************************************************;
options validvarname=v7;
libname pgxl xlsx "&path/data/prdsales_countries.xlsx" access=readonly;

* Create a single SAS data set with additional data preparation *;
data work.prd_sales_final;
	length Country $15. State $30.;
	set pgxl.usa
		pgxl.canada
		pgxl.mexico;
		
	* Create calculated columns *;
	TotalDiffPredict = Actual - Predict;
	PctDiffPredict = TotalDiffPredict / Predict;
	if TotalDiffPredict < 0 then MonthlyForecast = 'Under Predicted';
		else if TotalDiffPredict >= 0 then MonthlyForecast ='Over Predicted';
			
	* Drop columns *;
	drop County MonYR;
	
	* Format columns *;
	format Manager_ID z10. 
		   Actual Predict TotalDiffPredict dollar16.2
		   PctDiffPredict percent7.2
		   _CHARACTER_;
	
	* Add column labels *;
	label Country = "Business Country"
		  Actual = "Actual Income"
		  MonthlyForecast = "Monthly Forecast"
		  Predict = "Predicted Income"
		  TotalDiffPredict = 'Actual Amount Difference from Predicted'
		  PctDiffPredict = 'Actual Pct Diff from Predicted';
run;

libname pgxl clear;

proc sort data=prd_sales_final;
	by Country;
run;

proc print data=work.prd_sales_final(obs=20);
run;



**********;
* Step 4 *;
**********;
*********************************************************;
* Create an Excel report                                *;
*********************************************************;

******************************;
* Set file name and formats  *;
******************************;
* Create a dynamic file name by using the current date *;
%let currDate = %sysfunc(today(), yymmdd10.);
%let fileName = xls01d01_&currDate..xlsx;


* Create the Formats *;
%let Blue=CX33a3ff;
%let LightRed=cxFFA0A0;
%let LightGreen=cx99FFBB;
proc format;
	* Set color for negative and positive values *;
	value ColorFmt
		low - 0 = "&LightRed"
		0 <- high = "&LightGreen";
		
	* Bold negative values *;
	value WeightFmt
		low - 0 = "bold";
quit;

*********************************************Start Excel Report*********************************************;

***********************;
* Create Excel report *;
***********************;
ods _all_ close;
ods excel file = "&outpath/&fileName"
		  author = "Peter S"
		  title = "Company Sales Overview"
		  keywords = "Sales, Country Sales, Yearly Sales, Business Sales, Company Sales"
		  nogtitle
		  options(sheet_interval='none' embedded_titles='on' autofilter='all')
		  style=Daisy;


***********;
* Sheet 1 *;
***********;
ods excel options(sheet_name='Company Analysis' 
				  frozen_headers='2');   
			
* Worksheet title *;			
title height=18pt "Company Overview";
title2 " ";

* Bar plot *;
title3 justify=left "Company Sales by Year";
proc sgplot data=work.prd_sales_final;
	vbar Year / 
		response = Actual
		fillattrs=(color=&Blue)
		nooutline;
	xaxis display=(nolabel);
	yaxis display=(nolabel);
run;
title;

* Table *;
title justify=left "Company Actual vs Predicted Sales by Year";
proc print data=pg.yearly_sales label noobs;
	var Year Actual Predict ;
	var DiffFromPredict ForecastAccuracy / style(data)=[backgroundcolor=ColorFmt.           
	                                                    fontweight=WeightFmt.];
	sum Actual Predict;
    sum DiffFromPredict / style(grandtotal)=[backgroundcolor=ColorFmt.];  
run;
title;

* Line plot *;
title justify=left "Product Sales by Year";
proc sgplot data=work.prd_sales_final;
	vline Year / 
		response = Actual
		group=Product
		markers markerattrs=(symbol=circleFilled)
		lineattrs=(thickness=1pt pattern=solid)
		curvelabel;
	xaxis display=(nolabel);
	yaxis display=(nolabel);
	styleattrs datacontrastcolors=(blue green orange purple);
run;
title;


***********;
* Sheet 2 *;
***********;
ods excel options(sheet_name='Country Analysis' 
				  sheet_interval='now' 
				  autofilter='all' 
				  frozen_headers='2'
				  suppress_bylines='on' 
				  absolute_column_width='20,20,20,20');
	
* Worksheet title *;	
title height=18pt "Country Analysis";
title2 " ";	

* Table *;
title3 justify=left "Country Sales by Year";
proc print data=pg.country_sales noobs;
	var Country Year Actual Predict;
	var DiffFromPredict ForecastAccuracy / style(data)=[backgroundcolor=ColorFmt.           
	                                                    fontweight=WeightFmt.];
	sum Actual Predict;
	sum DiffFromPredict / style(grandtotal)=[backgroundcolor=ColorFmt.];  
run;
title;

* Line plot *;
title justify=left "Sales by Country";
proc sgplot data=work.prd_sales_final;
	vline Year / 
		response=Actual
		group=Country
		markers markerattrs=(symbol=circlefilled)
		lineattrs=(thickness=1pt pattern=solid)
		curvelabel;
	styleattrs datacontrastcolors=(red green blue);
	xaxis display=(nolabel);
	yaxis display=(nolabel);
run;
title;

title justify=left "#byval1 Yearly and Product Analysis"; 
proc report data=work.prd_sales_final
			spanrows;
	* Create BY groups *;
	by Country;
	
	* Select columns for the report *;
	columns ('Company Sales Overview' Country Year Product Actual);
	
	* Specify what each column does in the report *;
	define Country / group style( column)=[fontweight=bold fontsize=10pt];
	define Year / group;
	define Product / group;
	define Actual / analysis sum ;
	
	* Break after each group summarization *;
	break after Country / summarize suppress style=[fontweight=bold fontsize=12pt];
	compute after Country;
		Country = 'Grand Total';
	endcomp;
	
	break after Year / summarize suppress style=[fontweight=bold fontsize=11pt];
	compute after Year;
		Product = 'Total';
	endcomp;
run;
title;

ods excel close;



**********;
* Step 5 *;
**********;
***********************************************************;
* Download and open the Microsoft Excel workbook          *;
***********************************************************;
* Go to the main course folder and open the output        *;
* folder. Download and open the xls01d01_<DDMMMYYYY>.xlsx *;
* Excel workbook. The Excel workbook name will end with   *;
* today's date.                                           *;
***********************************************************;