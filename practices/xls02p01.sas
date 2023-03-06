************************************************************************;
* Practice 2.01: Creating an Excel Report                              *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


**********************************************************************************************************;
*                             STEP 1: Prepare the Excel data and create a SAS table                      *;
**********************************************************************************************************;

*****************************************************************************;
* a. Make a library reference to the cars.xlsx Excel workbook named carsxl. *;
*****************************************************************************;




* Prepare the data *;
data work.cars_final;
	set carsxl.cars;
	MPG_Avg = mean(MPG_City, MPG_Highway);
	format MSRP Invoice dollar16.;
	label MSRP = "Manufacturer's Suggested Retail Price"
		  MPG_City = 'MPG City'
		  MPG_Highway = 'MPG Highway'
		  MPG_Avg = 'MPG Average'
		  EngineSize = 'Engine Size (L)'
		  Weight = 'Weight (LBS)'
		  Wheelbase = 'Wheelbase (IN)'
		  Length = 'Length (IN)';
run;

* Sort the data *;
proc sort data=work.cars_final;
	by Origin Make Type Invoice;
run;


****************************************************************************;
* b. Clear the carsxl library reference.                                   *;
****************************************************************************;





**********************************************************************************************************;
*                                   STEP 2: Create the Excel Report                                      *;
**********************************************************************************************************;

****************************************************************************;
* a. .                                   *;
****************************************************************************;
ods noproctitle;
%let currentDate = %sysfunc(today(),yymmdd10.);



*****************;
* Worksheet 1   *;
*****************;

****************************************************************************;
* b. xxx.                                   *;
****************************************************************************;

         


title height=18pt 'Detailed Car Data';

title2 "Data as of &currentDate";
proc print data=work.cars_final label noobs;
	id Make Model;
run;
title;



*****************;
* Worksheet 2   *;
*****************;

****************************************************************************;
* c. xxx.                                   *;
****************************************************************************;


   

ods graphics / width=8in;
title height=14pt justify=left 'Miles Per Gallon by Car Make';
proc sgplot data=work.cars_final;
	vbar Make /
		response=MPG_Avg
		stat=mean
		categoryorder=respdesc
		nooutline
		fillattrs=(color=dodgerBlue);
	yaxis label='MPG Average';
quit;
ods graphics / reset;
title;

proc means data=work.cars_final min mean max maxdec=0;
	class Make;
	var MPG_Avg;
quit;



*****************;
* Worksheet 3   *;
*****************;

****************************************************************************;
* d. xxx.                                   *;
****************************************************************************;


				  

* Create the format *;
proc format;
	value under30kMSRP
		0 - 30000 = "lightblue";
quit;


****************************************************************************;
* e. xxx.                                   *;
****************************************************************************;
title height=10pt 'Highlighted cars are under $30,000';
proc print data=work.cars_final noobs;
	by Origin;
	var Make Model Type;
	var MSRP;  
run;


ods excel close;