***********************************************************************************************;
* COURSE: Working with SAS and Microsoft Excel (PGXLM7)                                       *;
* DATE CREATED:2/17/2023                                                                      *;
* SETUP: Creates the course data in the other SAS environments                                *;   
*    - SAS OnDemand for Academics                                                             *;
*    - SAS Viya for Learners and SAS Viya                                                     *;
*    - SAS installed in a desktop                                                             *;
*    - SAS installed on a server (write access required)                                      *;
* DESCRIPTION:                                                                                *;
*     - This createCourseFiles_EPGXLM7.sas program sets up your SAS environment to take this  *;
*       course. The program downloads the course zip file from the internet and unpacks the   *;
*       zip file with all course folders and SAS programs in the specified writable path      *;
*       below. After all files are unpacked, the cre8data_EPGXLM7.sas program is executed and *;
*       all data and the libname.sas program is created.                                      *;
*         a. (Required) USER HAS WRITE ACCESS to specified path below on the SAS server       *;
*         B. (Not supported) USER DOES NOT HAVE WRITE ACCCESS to the SAS server.              *;
*            You will need to contact your system admin to get write access to a location on  *; 
*            the SAS server.                                                                  *;
* FILE(S) CREATED:                                                                            *;
*     1. folders: activities, data, demos, examples, output, practices                        *;
*     2. program: libname.sas program (in the main course folder)                             *;
*     3. course data: 5 SAS data sets and 4 Excel workbooks (in the data folder)              *;                                                                         
* REQUIREMENTS:                                                                               *;
*    - Specify a writable path on the SAS environment to unzip the course folders and files   *;
*      to. If the folder path is not writable, an error will occur.                           *;
*    - This program will not run properly on z/OS. Only Windows, Linux and UNIX are supported.*;
*      Values for PATH are CASE SENSITIVE.                                                    *;
***********************************************************************************************;

/* Replace FILEPATH with the full path to your EPGXLM7 folder */
%let path = FILEPATH;

*************************************************;
* EXAMPLES                                      *;
*************************************************;
* %let path=/home/usersas/EPGXLM7;              *;
* %let path=S:/workshop/EPGXLM7;                *;
* %let path=/shared/home/usersas/PGXLS          *;
*************************************************;











/*********************************************************************************************************
 WARNING: DO NOT ALTER CODE BELOW THIS LINE IN ANY WAY
*********************************************************************************************************/

/* Make sure path consistently uses forward slashes */
%let path=%qsysfunc(translate(%superq(path),/,\));
%let original_path=%superq(path);

/* Make sure there is no slash at the end of the path. If there is one remove it */
%put &=path;
data _null_;
	showPath = "&path";
	if substr(reverse(showPath),1,1) in ('/','\') then do;
		newPath = substr(showPath,1,length(showPath)-1);
		call symputx('path',newPath);
	end;
run;

* Holds the path specified from the this program and stores it in a new macro variable. *;
* If this program was used, it will use this path in the cre8data program.              *;
%let _createdataEPGXLM_used_ = %superq(path);


/* options nomprint nosymbolgen nonotes nosource dlcreatedir; */
/* options mprint symbolgen notes source; */


***************************************************************************;
* Create unpack macro program                                             *;
***************************************************************************;
%macro unpack(unzip /* Full path pointing to where to create the EPG2V2 data */
             ,zipfilename /* ZIP File name (used to download with PROC HTTP) */
             ,urlzipdownload);


* Create global and local macro variables *;
%local rc fid fileref fnum memname big_zip big_zip_found data_zip data_zip_found url;
%global cre8data_success path;
%let cre8data_success=0;

* URL to download the course zip file *;
%let url=%str(&urlzipdownload);

/* Is the path specified valid? */
%let fileref=unzip;  /*Unzip is the path value */
%let rc=%sysfunc(filename(fileref,%superq(unzip)));
%let path_found=%sysfunc(fileref(unzip));

/*If path is not found, then return an error */
%if &path_found ne 0 %then %do;
   %put %sysfunc(sysmsg());
   %put ERROR: ************************************************************;
   %put ERROR- Path specified for data files (%superq(unzip)) is not valid.;
   %put ERROR- Remember: PATH values in UNIX and LINUX are case sensitive. ;
   %put ERROR- ************************************************************;
   %let rc=%sysfunc(filename(fileref));
   %return;
%end;


/* Get just the filename of the zipfile, not the .ZIP extension */
%if %qscan(%qupcase(%superq(zipfilename)),2,.) = %str(ZIP) %then %do;
   %let zipfilename = %qscan(%superq(zipfilename),1,.) ;
%end;

/* Test for the presence of the main ZIP file in the file path */
%let fileref=bigzip;
%let rc=%sysfunc(filename(fileref,%superq(unzip)/%superq(zipfilename).zip,zip));
%let big_zip_found=%sysfunc(fileref(bigzip));

/* If the main course zip file is not in the specified path, download the zip file from the internet */
%if &big_zip_found ne 0 %then %do;
   %put NOTE: *******************************************************************;
   %put NOTE: %superq(zipfilename).zip not found in %superq(unzip).;
   %put NOTE: Attempting to download the ZIP file from the internet.;
   %put NOTE: *******************************************************************;
   
/* Check to see if the download will be successful */
   proc http url="%superq(url)";
   run;
   
/*    If the download is unsuccessful, return a download error */
   %if &SYS_PROCHTTP_STATUS_CODE = 404 %then %do;
      %put ERROR: *******************************************************************;
      %put ERROR- Attempt to download %superq(zipfilename).zip from the following url;
      %put ERROR- %superq(urlzipdownload) was unsuccessful.                          ;
      %put ERROR- Your SAS environment might not allow you to download files from the internet.;
      %put ERROR- If this is the case, please follow the instructions to manually download;
      %put ERROR- the zip file and upload it to your course folder. Then rerun this program.;
      %put ERROR- *******************************************************************;
   %end;
   %else %do; /* Otherwise download the zip file to the specified path */
	   filename BigZip "%superq(unzip)/%superq(zipfilename).zip";
	   
	   /* Download the zip file and save it to the folder */
	   proc http 
	      url="%superq(url)"
	      out=BigZip method="get";
	   run;
	   
	   /*Successful download note to the log */
	   %if &SYS_PROCHTTP_STATUS_CODE = 200 %then %do; 
      	  %put NOTE: *********************************************************************;
          %put NOTE- The zip file %superq(zipfilename).zip was successfully downloaded from ;
          %put NOTE- the internet at the following url %superq(urlzipdownload). ;
          %put NOTE- *******************************************************************;
       %end;
   %end;
%end;
%else %do; /* Note that the zip file was already found in the specified folder */
   	%put NOTE: *******************************************************************;
	%put NOTE: course zip file %superq(zipfilename).zip was found in the specified;
   	%put NOTE: path %superq(unzip). Will unpack this zip file for the course.     ;
   	%put NOTE: *******************************************************************;
%end;


/* Unpack the zip file */
options dlcreatedir;
libname xx "%superq(path)";
libname xx clear;

/* Read the "members" (files) from the ZIP file */
/* Create the data folder structure and get a list of files in macro variables */
filename BigZip zip "%superq(unzip)/%superq(zipfilename).zip";
data _null_;
   length memname pathname $500;
   fid=dopen("bigzip");
   if fid=0 then stop;
   memcount=dnum(fid);
   do i=1 to memcount;
      memname=dread(fid,i);
/*       Create and empty folder for each folder in the ZIP file */
/*       check for trailing / in folder name */
      isFolder = (first(reverse(trim(memname)))='/');
        if isfolder then put memname= isfolder=;
      if isfolder then do;
         pathname=cats("&path/",substr(memname,1,length(memname)-1));
         put "NOTE: Creating path " pathname;
         rc1=libname('xx',pathname);
         rc2=libname('xx');
      end;
      else do;
         filecount+1;
         call symputx(cats('out',filecount),memname,'L');
      end;
   end;
   rc=dclose(fid);
   call symputx('filecount',filecount,'L');
run;

%do i=1 %to &filecount;
   filename out "%superq(unzip)/%superq(out&i)";
    data _null_;
      infile bigzip(%superq(out&i))
      lrecl=256 recfm=F length=length eof=eof unbuf;
      file out  lrecl=256 recfm=N;
      input;
      put _infile_ $varying256. length;
      return;
    eof:
      stop;
   run;
%end;

* Clear filename references *;
filename bigzip;
filename out;
filename unzip;


* Create a macro variable pointing to the 2_cre8data_other.sas program *;
%let cre8data_program=%superq(unzip)/data/cre8data/cre8data_EPGXLM7.sas;

* Check for a cre8data.sas program and execute. Return an error if not found *;
%let cre8data_ready=%sysfunc(fileexist(%superq(cre8data_program)));

* Return error of 2_cre8data_other.sas program is not found *;
%if not &cre8data_ready %then %do;
      %put;
      %put ERROR: *************************************************************************;
      %put ERROR- After unzipping %superq(zipfilename).zip, cre8data_EPGXLM7.sas program  was not found ;
      %put ERROR- in folder %superq(unzip).;
      %put ERROR- *************************************************************************;
      %put;
%end;

* Execute cre8data.sas program from the data folder *;
%include "&cre8data_program";

%mend unpack;



/*************************************************************
 Execute the macro program to unzip the course data
*************************************************************/
%unpack(%superq(path), 
		EPGXLM7.zip, 
		https://github.com/pestyld/xls/raw/main/EPGXLM7.zip)
