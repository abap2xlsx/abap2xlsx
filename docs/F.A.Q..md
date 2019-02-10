* Verify the version that you have installed

Starting from version 7.0.1 we introduced the version tracking.
To verify the version that you have installed in your system just generate a demo report and verify the properties of the generated XLSX file. In field description you should find the current version.
You can also check the attribute VERSION in the class ZCL_EXCEL.

* After import error "Interface method are not implemented" is thrown

This is by a well know problem with SAPLink, Import nugget twice.
 
* Demo reports do not compile, class CL_BCS_CONVERT is not available.

Implement SAP OSS Notes:

[Note 1151257 - Converting document content](https://service.sap.com/sap/support/notes/1151257)

[Note 1151258 - Error when sending Excel attachments](https://service.sap.com/sap/support/notes/1151258)
 
* You want to start a report with SUBMIT. The report you called has a parameter of type string. If a variable of type string is transferred to the parameter with SUBMIT, the system issues termination message DB036

Implement SAP OSS Notes:

[Note 1385713 - SUBMIT: Allowing parameter of type STRING](https://service.sap.com/sap/support/notes/1385713)
 
* Macro-Enabled workbook
 
Run report ZDEMO_EXCEL29 and use as VBA source file [TestMacro.xlsm](https://github.com/ivanfemia/abap2xlsx/blob/master/resources/TestMacro.xlsm) attached.
Basically abap2xlsx works using an existing VBA binary (we do not want to create a VBA editor).

* Download XLSX files in Background

Run report ZDEMO_EXCEL25.

* Use Calibri font in Excel and auto width calculation

You can upload Excel's default font Calibri from the regular .TTF font files using transaction SM73. Make sure to upload all four files (standard, bold, italic, bold and italic) and to use the description "Calibri", i.e. exactly the name Excel displays for the font. 
 
