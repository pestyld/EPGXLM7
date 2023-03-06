***********************************************************************************************;
* COURSE: Working with SAS and Microsoft Excel (PGXLM7)                                       *;
* DATE CREATED:2/17/2023                                                                      *;
* DESCRIPTION:                                                                                *;
*     - VIRTUAL LAB: Simply run this program to create the data and libname.sas program.      *;
*     - OTHER ENVIRONMENTS: The otherSAS_setup.sas downloads the course zip file              *;
*       from the internet and unpacks the zip file with all course folders and SAS programs.  *; 
*       A path is specified in the otherSAS_setup.sas and then runs this program              *;
*       at the end to create the course data and libname.sas program.                         *;  
*         a. (Required) USER HAS WRITE ACCESS to the SAS server                               *; 
*         B. (not supported) USER DOES NOT HAVE WRITE ACCCESS to the SAS server.              *;
*            You will need to contact your system admin to get write access to a location on  *; 
*            the SAS server.                                                                  *;
* SETUP: Creates the course data in the following environments                                *;   
*    - SAS virtual lab setup (eLearning and instructor led labs)                              *;
*    - Other SAS Environments:                                                                *;
*        a. SAS OnDemand for Academics                                                        *;
*        b. SAS Viya for Learners                                                             *;
*        c. SAS installed in a desktop                                                        *;
*        d. SAS installed on a server (write access required)                                 *;
* FILE(S) CREATED:                                                                            *;
*     1. program: libname.sas program (in the main course folder)                             *;
*     2. course data: 5 SAS data sets and 4 Excel workbooks (in the data folder)              *;                                                                         
* REQUIREMENTS:                                                                               *;
*  - Virtual lab: Course data must be in S:/workshop.                                         *;
*  - Other environments (write) : User must have write access to the SAS server and ONLY run  *;
*                                 otherSAS_setup.sas program.                                 *;
*  - Other environments (no write access) : If user does not have write access to the SAS     *;
*                                           server you will need to contact your system admin.*;                                    *;
***********************************************************************************************;
***********************************************************************************************;
*                             DO NOT MODIFY THE CODE BELOW                                    *;
***********************************************************************************************;




/*  */
/* %let path=s:/workshopx; */


**************************************************************************************************;
* INITIAL PATH MACRO VARIABLE                                                                    *;
**************************************************************************************************;
* Test to see if the path macro variable is already set.                                         *;
* If it is not set, assume a virtual lab is being used and set path to s:/workshop.              *;                                                      *;
**************************************************************************************************;
%macro checkPathMacroVariable;

/* Check to see if the path macro variable is set. If it's not set use s:/workshop default for virtual labs */
%if %symexist(path) = 0 %then %do;
	%let path = S:/workshop;
	%put NOTE: ***********************************************************************************;
	%put NOTE: Path macro variable not found. Assuming course is being taken in a SAS virtual lab.;
	%put NOTE: The path macro variable will be set to the s:/workshop folder. This is the default folder for the SAS virtual lab.;
	%put NOTE: ***********************************************************************************;
%end;
%else %do; /* If it's already set, keep that location */
	%put NOTE: Path macro variable found. Will use the folder path &path as the course folder location.;
%end;

/* Check to see if that the path specified valid? */
%let pathExists=%sysfunc(fileexist(%superq(path)));

/* If path is not found, then return an error */
%if &pathExists = 0 %then %do;
   	%put %sysfunc(sysmsg());
   	%put ERROR: ***********************************************************************************;
   	%put ERROR: Path specified for data files in (%superq(path)) is not valid. ;
   	%put ERROR: ***********************************************************************************;
   	%put ERROR- NOTE: If you already have your course setup an are trying to recreate the course data,;
   	%put ERROR-       please run the libname.sas program first to set the path of the course.;
   	%put ERROR-       Then run this program.;
    %put ERROR- OTHERWISE: The path that is specified is not valid. This can occur for a variety of reasons.;
    %put ERROR-            Please retry the Using Other SAS Environments instructions provided with the course.; 	
   	%put ERROR- ************************************************************************************;
	%abort cancel;
%end;
%else %do; /* Return note confirming path exists */
	%put NOTE: ***********************************************************************************;
	%put NOTE: Confirming the following folder path exists: &path.;
	%put NOTE: The folder path does exists. Will attempt to create course data.;
	%put NOTE: ***********************************************************************************;
%end;

%mend;
%checkPathMacroVariable
;


/* *************************************************; */
/* * CHECK FOR IBT/LW LAB OR ELEARNING LAB PATH    *; */
/* *************************************************; */
/* * IBT/LW - path is s:/workshop                  *; */
/* * eLearning - path is s:/workshop/coursecode    *; */
/* * Macro here determines which path to use       *; */
/* *************************************************; */
/*  */
/* %macro virtual_lab_path(coursecode,          /* Course code used as the folder name in the elearning lab */
/*                         programNameCheck);   /* The create data program that is being located            */
/*  */
/* 	* Specify the eLearning course folder name *; */
/* 	%let elearningCourseFolder = &coursecode; */
/* 	 */
/* 	* Reference the cre8data.sas program in the DATA folder using the IBT/LW path *; */
/* 	filename cre8data "&path/data/cre8data/&programNameCheck"; */
/*  */
/* 	* Check to see if cre8data.sas is in the IBT/LW folder path. If not, change to elearning lab path *; */
/* 	data _null_; */
/* 		cre8data_exists = fexist('cre8data'); */
/* 		if cre8data_exists = 0 then call symputx('path',"&path./&elearningCourseFolder"); */
/* 	run; */
/* 	 */
/* 	filename cre8data; */
/* %mend; */
/*  */
/* %virtual_lab_path(EPGXLM7, cre8data_EPGXLM7.sas); */



*************************************************************************************;
* CHECK FOR otherSAS_setup.sas PROGRAM TO CHANGE PATH IF NECESSARY                  *;
*************************************************************************************;
* Macro program tests to see if the otherSAS_setup.sas program was used.            *;
* If it was used, it changes the location of the path macro variable to the value   *;
* _createdataEPGXLM_used_, which was created in the otherSAS_setup.sas              *;
* program. It holds the user's specified writable location.                         *;
*************************************************************************************;
%macro check_if_createdataEPGXLM();
       
		* Test to see if the otherSAS_setup.sas was used.  *;
		* If so, change the location of the path macro variable       *;
        %if %symexist(_createdataEPGXLM_used_) = 1 %then %do;
        	%let path=&_createdataEPGXLM_used_;
        	%put %str(NOTE: The createdataEPGXLM.sas program was used to unpack all the files.);
        	%put NOTE: Changed the location of the path macro variable to &_createdataEPGXLM_used_;
        %end;
		%else %do; * Otherwise leave the path macro variable *;
			%put %str(NOTE: The createdataEPGXLM.sas program was not used.);
			%put %str(NOTE: Using the path s:/workshop from the the virtual lab.);
		%end;
		
		%put &=path;
		
%mend check_if_createdataEPGXLM;

%check_if_createdataEPGXLM();



***************************************************;
* CREATE AND RUN THE LIBNAME.SAS PROGRAM          *;
***************************************************;
* Create the libname.sas program in the root      *;
* course folder. Then run the libname.sas program *;
* to set the libref and macro variables.          *;
***************************************************;

data _null_;
	file "&path./libname.sas";
	put '%let path='"&path.;";
   	put '%let outpath='"&path./output;";
  	put 'libname pg "&path./data";';
run;
%include "&path/libname.sas";



*************************************************;
* CREATE MACRO VARIABLE FOR DATA FOLDER         *;
*************************************************;
 %let datapath = &path./data; 
 libname pg "&datapath"; 
 	
/*  * Code if write access to SAS *;  */
/*  %if %symexist(EnterpriseGuideNoWriteAccess) = 0 %then %do;  */
/*  	%let datapath = &path./data;  */
/*  	libname pg "&datapath";  */
/*  %end;  */
/*  %else %do; * Code if EG on desktop with no write access *;  */
/*  	%let datapath = &tempDataPath;  */
/*  	libname pg "&datapath";  */
/*  %end;  */



******************************************;
* 1. USER MACROS PROGRAMS TO CREATE DATA *;
******************************************;
options validvarname=ANY;

****************************;
* a. Delete all .bak files *;
****************************;
%macro del_bak_files(filePath = );
	data _null_;
		length folder $8 file_name $300;
		folderPath=filename(folder, "&filePath");
		folderPath=dopen(folder);
	
		do i=1 to dnum(folderPath);
			file_name=dread(folderPath, i);
			if find(file_name,'.bak','i') then do;
				fname="tempbak";
	    		rc=filename(fname, cats("&filePath",'/',file_name));
	    		rc=fdelete(fname);
	    		put "NOTE: DELETED " file_name;
			end;
			output;
		end;
		folderPath=dclose(folderPath);
		folderPath=filename(folder);
		*keep file_name found;
	run;
%mend;


************************************************;
* b. Make simple XLSX file from a SAS data set *;
************************************************;
* make an Excel file with all data *;
%macro make_xlsx_file(ds = );
	%let ds_name = %scan(&ds,2,'.()');
	
	proc export data=&ds
				dbms=xlsx 
				outfile = "&datapath./&ds_name..xlsx"
				replace;
	quit;
%mend;


*************************************************************************;
* c. Create a XSLX workbook with multiple worksheets by specific groups *;
*************************************************************************;
%macro xl_output_by_groups(input_dsn = , group_name = , column =);
	data xl.&group_name;
	     set &input_dsn;
	     if upcase(&column) = upcase("&group_name") then output;
	run;
%mend;



****************************************;
* 2. CREATE SAS DATA SETS FROM SASHELP *;
****************************************;

*****;
* a *;
************************************;
* SASHELP.CARS                     *;
************************************;
* This will be used for activities *;
************************************;

*************;
* CARS      *;
*************;
* Make a SAS data set *;
proc sql noprint;
	select mean(MPG_City) as MeanMPG, 
	       mean(MSRP) as MeanMSRP
		into :meanMPG, :meanMSRP
		from sashelp.cars;
quit;
%put &=meanMPG &=meanMSRP;
	
data pg.cars;
	set sashelp.cars;
	if MPG_City > &meanMPG and MSRP < &meanMSRP then CarGroup='High MPG, Low Cost';
    else CarGroup='Low MPG, High Cost';
run;


*************;
* cars.xlsx *;
*************;
* Make SAS data set an XLSX file *;
%make_xlsx_file(ds = pg.cars);



********************;
* cars_origin.xlsx *;
********************;
%let make_excel_file = cars_origin.xlsx;
libname xl xlsx "&datapath/&make_excel_file";

%xl_output_by_groups(input_dsn = pg.cars , group_name = Asia, column = Origin)
%xl_output_by_groups(input_dsn = pg.cars , group_name = USA, column = Origin)
%xl_output_by_groups(input_dsn = pg.cars , group_name = Europe, column = Origin)
;
libname xl clear;



*****;
* b *;
**************************************;
* SASHELP.PRDSALES(2 and 3)          *;
**************************************;
* This will be used for demos/slides *;
**************************************;

* Get column names *;
proc sql noprint;
	select catx('=',Name,propcase(Name))
		into :renameCols separated by ' '
	from dictionary.columns
	where libname='SASHELP' and memname='PRDSAL2';
quit;
%put &=renameCols;



**********************;
* prdsales_raw       *;
**********************;
data pg.prdsales_raw;
	set sashelp.prdsal2(rename=(&renameCols))
		sashelp.prdsal3(rename=(DATE=MONYR)) 
		indsname=curr_data_set_name;
	call streaminit(10);
	
* Remove periods from U.S.A *;
	if Country = 'U.S.A.' then Country = 'USA';
	
* Update the dates from the sashelp data *;
	if curr_data_set_name = "SASHELP.PRDSAL2" then Year = Year + 22;
		else if curr_data_set_name = "SASHELP.PRDSAL3" then Year = Year + 24;
		
	_temp_month_numeric_value = month(monyr);
	_temp_last_day_of_month = day(intnx('month',mdy(_temp_month_numeric_value,1,YEAR),1) - 1);
	Date = mdy(_temp_month_numeric_value, _temp_last_day_of_month, YEAR);
	Monyr = DATE;
	
* Add a Distinct Manager ID to the data *;
	if Country = 'Mexico' then Manager_ID=88080;
		else if Country = 'Canada' then Manager_ID=99603;
		else if State = 'California' then Manager_ID=1129;
		else if State in ('Colorado', 'Texas') then Manager_ID=2399;
		else if State = 'Illinois' then Manager_ID = 3672;
		else if State = 'New York' then Manager_ID = 4102;
		else if State in ('North Carolina', 'Florida', 'Washington') then Manager_ID=5903;

* Fix Actual values of 0;
	if Actual = 0 then Actual = rand('uniform',350,1500);
	
	
* Clean and update acutal/predict values based on location *;
	if Country = 'USA' then do;
		Actual = Actual * 9;
		Predict = Predict * 8.5;
		* Decrease the sales of beds in 2021 and 2022*;
		if (Product = 'BED' and Year in (2021,2022)) then Actual = Actual * .70;	
	end;
	else if Country in ('Mexico','Canada') then do;
		Actual = Actual * 5;
		Predict = Predict * 4.5;	
	end;
	if Year in (2021, 2022) then Predict = Predict * .90;

* Fix predict values less than 500;
	if Predict < 500 then Predict=rand('uniform',.60,1.05) * Actual;

* Format and drop *;
	format  Date date9. Manager_ID z10.;
	drop _temp: Month;
run;


***************************;
* prdsales_final          *;
***************************;
data  pg.prdsales_final;
	set pg.prdsales_raw;
		
	* Create calculated columns *;
	TotalDiffPredict = Actual - Predict;
	
	* The Forecast Accuracy Formula (Percentage Error): ((Actual-Forecast)/(Actual)) * 100 *;
	PctDiffPredict = TotalDiffPredict / Actual;
	if TotalDiffPredict < 0 then MonthlyForecast = 'Under Predicted';
		else if TotalDiffPredict >= 0 then MonthlyForecast ='Over Predicted';
			
	* Drop columns *;
	drop MonYR;
	
	format Manager_ID z10. 
		   Actual Predict TotalDiffPredict dollar16.2
		   PctDiffPredict percent7.2
		   _CHARACTER_;
	
	label Country = "Business Country"
		  Actual = "Actual Sales"
		  MonthlyForecast = "Monthly Forecast"
		  Predict = "Predicted Sales"
		  TotalDiffPredict = 'Difference from Predicted'
		  PctDiffPredict = 'Pct Diff from Predicted'
		  Manager_ID = 'Manager ID';
run;



*********************************;
* country_sales & yearly_sales  *;
*********************************;
proc means data=pg.prdsales_final noprint;
	class Country Year;
	var Actual Predict;
	output out=work.country_sales(where=(_TYPE_ in (1,3)))
		   sum(Actual)=Actual
		   sum(Predict)=Predict;
run;

* yearly_sales *;
data pg.yearly_sales;
	set work.country_sales;
	where _TYPE_ = 1;
	Actual = round(Actual,.01);
	Predict = round(Predict, .01);
	DiffFromPredict = round(Actual - Predict,.01);
	ForecastAccuracy = (Actual - Predict) / Actual;
	if DiffFromPredict < 0 then MonthlyForecast = 'Under Predicted';
		else if DiffFromPredict >= 0 then MonthlyForecast ='Over Predicted';
	format DiffFromPredict dollar16.2
	       ForecastAccuracy percent8.2;
	label DiffFromPredict = 'Difference from Predicted'
		  ForecastAccuracy = 'Forecast Accuracy';
	drop _TYPE_ _FREQ_ Country;
run;

* country_sales *;
data pg.country_sales;
	set work.country_sales;
	where _TYPE_ = 3;
	Actual = round(Actual,.01);
	Predict = round(Predict, .01);
	DiffFromPredict = round(Actual - Predict,.01);
	ForecastAccuracy = (Actual - Predict) / Actual;
	format DiffFromPredict dollar16.2
	       ForecastAccuracy percent8.2;
	label DiffFromPredict = 'Difference from Predicted'
		  ForecastAccuracy = 'Forecast Accuracy';
	drop _TYPE_ _FREQ_;
run;



***************************;
* prdsales_raw.xlsx       *;
***************************;
* Make SAS data set an XLSX file (prdsales_raw) *;
%make_xlsx_file(ds = pg.prdsales_raw(rename=(Manager_ID='Manager ID'n)));



***************************;
* prdsales_countries.xlsx *;
***************************;
* Make an Excel file with each country on it's own worksheet (prdsales_raw) *;

%let make_excel_file = prdsales_countries.xlsx;
libname xl xlsx "&datapath/&make_excel_file";

%xl_output_by_groups(input_dsn = pg.prdsales_raw(rename=(Manager_ID='Manager ID'n)), group_name = USA, column = Country);
%xl_output_by_groups(input_dsn = pg.prdsales_raw(rename=(Manager_ID='Manager ID'n)), group_name = Canada, column = Country);
%xl_output_by_groups(input_dsn = pg.prdsales_raw(rename=(Manager_ID='Manager ID'n)), group_name = Mexico, column = Country);

libname xl clear;



**********************************;
* 3. Delete all .bak Excel files *;
**********************************;
%del_bak_files(filePath = &path./data);



**********************************;
* View available SAS data sets   *;
**********************************;
proc contents data=pg._all_ nods;
run;


**********************************;
* Delete macro variables         *;
**********************************;
%if %symexist(_createdataEPGXLM_used_) = 1 %then %do;
	%symdel _createdataEPGXLM_used_;
    %put %str(NOTE: Deleted _createdataEPGXLM_used_);
%end;
%if %symexist(tempDataPath) = 1 %then %do;
	%symdel tempDataPath;
    %put %str(NOTE: Deleted tempDataPath);
%end;
