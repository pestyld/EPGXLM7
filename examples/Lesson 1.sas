*********************;
* Lesson 1 Examples *;
*********************;

*****************;
* SECTION 2     *;
*****************;

* PROC IMPORT *;
proc import datafile="&path/data/prdsales_countries.xlsx"
            dbms=xlsx
            out=work.USA 
            replace;
    sheet="USA"; 
run;


* PROC EXPORT *;
proc export data=pg.prdsales_final
            dbms=xlsx 
            outfile="&outpath/prd_sales_data.xlsx"
            replace;
    sheet="Raw Data"; 
run;



*****************;
* SECTION 3     *;
*****************;

* LIBNAME Statement *;
libname pgxl xlsx "&path\data\prdsales_countries.xlsx";

data work.USA;
    set pgxl.USA;
run;

* This will add a worksheet to our course prdsales_countries.xlsx workbook that we use in this course *;
/* data pgxl.California; */
/*     set pgxl.USA; */
/*     where State = 'California'; */
/* run; */

proc print data=pgxl.USA(obs=10);
run;


proc freq data=pgxl.USA;
    tables State;
run;