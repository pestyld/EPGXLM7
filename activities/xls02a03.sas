************************************************************************;
* Activity 2.02: Modifying Table and Worksheet Appearance              *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


ods excel file="&outpath/xls02a03.xlsx";
		  	      
title height=16pt "Cars Price Data";
proc print data=pg.cars noobs label;
	ID Make Model;
run;
title;

ods excel close;




























*****************************;
* SOLUTION                  *;
*****************************;
/*
ods excel file="&outpath/xls02a03.xlsx"
          options(embedded_titles='on'
                  autofilter='on'
                  frozen_headers='3'
                  frozen_rowheaders='2');
		  	      
title height=16pt "Cars Price Data";
proc print data=pg.cars noobs label;
	ID Make Model;
run;
title;

ods excel close;
*/