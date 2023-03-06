************************************************************************;
* Activity 2.01: Creating a Simple Excel Report                        *;
* NOTE: Execute the libname.sas program if necessary.                  *;
************************************************************************;


* Add the ODS EXCEL statement below *;


* Create a visualization *;
title height=14pt justify=left color=charcoal 'Cars with Lower than Average MSRP and Higher than Average City MPG';
footnote justify=left color=charcoal 'Created by Peter S';
proc sgplot data=pg.cars 
            noautolegend 
            noborder;
    scatter x=MPG_City y=MSRP
            / group=CarGroup
              markerattrs=(symbol=CircleFilled size=7pt)
              markeroutlineattrs=(color=white thickness=.5px)
              filledoutlinedmarkers;
    yaxis offsetmin=0;
    styleattrs datacolors=(LightGray cx33a3ff);
run;
title;
footnote;

* Print all rows *;
proc print data=pg.cars label noobs;
	var Make Model MPG_City MSRP CarGroup;
	where CarGroup = 'High MPG, Low Cost';
run;

* Add the ODS EXCEL close statement below *;



	
























*****************************;
* SOLUTION                  *;
*****************************;
/*
ods excel file="&outpath/xls02a01.xlsx";

* Create a visualization *;
title height=14pt justify=left color=charcoal 'Cars with Lower than Average MSRP and Higher than Average City MPG';
footnote justify=left color=charcoal 'Created by Peter S';
proc sgplot data=pg.cars 
            noautolegend 
            noborder;
    scatter x=MPG_City y=MSRP
            / group=CarGroup
              markerattrs=(symbol=CircleFilled size=7pt)
              markeroutlineattrs=(color=white thickness=.5px)
              filledoutlinedmarkers;
    yaxis offsetmin=0;
    styleattrs datacolors=(LightGray cx33a3ff);
run;
title;
footnote;

* Print all rows *;
proc print data=pg.cars label noobs;
	var Make Model MPG_City MSRP CarGroup;
	where CarGroup = 'High MPG, Low Cost';
run;


ods excel close;
*/