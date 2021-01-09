## DISCLAIMER
As far as abap2xlsx is concerned, SAPlink is obsolete and should only be used on systems lower than 702.

## Prerequisites
* [SAPlink](http://www.saplink.org) installed in your system.
* [SAPlink Plugins](http://www.saplink.org) installed in your system: DDic, Interface (I would suggest to install the complete nugg package in build/SAPlink-plugins_Daily.nugg)SAP

## Procedure
Download the nugg file from [build folder](https://github.com/sapmentors/abap2xlsx/tree/master/build) and save it locally on your system. Logon on your SAP system and execute report ZSAPLINK Select "Import Nugget" and locate your nugg file, check overwrite originals only if you have a previous installation of abap2xlsx and you want to update.
 
Execute the report.
If you have checked overwrite originals a popup could appears in order to confirm the overwrite; press Yes to all.
You should get this result (all green light)
 
SAPLinks puts the objects into $tmp space and all the objects are inactive. So we need to activate now.
From SE80, select inactive objects
 
Activate objects in the following order:
I tried a new nugg import on a new system here the detailed steps I performed:

1. Activate all domains
1. Activate all data elements
1. Activate all Database Tables / Structures except: ZEXCEL_S_FIELDCATALOG, ZEXCEL_S_STYLEMAPPING, ZEXCEL_S_WORKSHEET_COLUMNDIME, ZEXCEL_S_WORKSHEET_ROWDIMENSIO
1. Activate all Table Types except: ZEXCEL_T_FIELDCATALOG, ZEXCEL_T_STYLEMAPPING, ZEXCEL_T_WORKSHEET_COLUMNDIME, ZEXCEL_T_WORKSHEET_ROWDIMENSIO
1. Activate all interface/classes (activate anyway)
1. Activate remaining Database Tables /  Structures (if any error occurs open the structure and double click on  the class object, SAP needs to refresh its buffer): ZEXCEL_S_FIELDCATALOG, ZEXCEL_S_WORKSHEET_COLUMNDIME, ZEXCEL_S_WORKSHEET_ROWDIMENSIO
1. Activate remaining Table Types (if any error occurs open the structure and double click on the class object, SAP  needs to refresh its buffer): ZEXCEL_T_FIELDCATALOG, ZEXCEL_T_WORKSHEET_COLUMNDIME, ZEXCEL_T_WORKSHEET_ROWDIMENSIO
1. Activate all demo reports
1. ~~Due to a issue with interfaces in SAPlink you need to import the abap2xlsx nugget twice and activate again all the objects.~~ Fixed with latest SAPlink release

Report ZAKE_SVN_A2X is used to retrieve and commit object into subversion server, Delete it if you are not a contributor.

[Getting ABAP2XLSX to work on a 620 System](Getting-ABAP2XLSX-to-work-on-a-620-System)
