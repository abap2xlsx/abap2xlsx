In the [contribution process](https://github.com/abap2xlsx/abap2xlsx/blob/main/CONTRIBUTING.md), before merging the pull request, there are 3 types of automated tests which should be executed, to detect regressions:
- ABAP Unit tests
- Automated integration tests
- Syntax checks

The procedure was written for ABAP 7.52, it can vary in the other ABAP versions.

# ABAP Unit tests
[ABAP Unit](https://help.sap.com/doc/abapdocu_latest_index_htm/latest/en-US/index.htm?file=abenabap_unit.htm) is the SAP technology for Unit Tests in the ABAP language. These tests are provided with abap2xlsx under ABAP local test classes stored in the productive `ZCL_EXCEL*` classes.

You may run these tests in two ways:
- In Eclipse ADT, select the main abap2xlsx package (you have chosen its name when installing abap2xlsx) in the Project Explorer view, run the tests by pressing <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F10</kbd> or via the context menu Run As > ABAP Unit Test, you'll see the results in the ABAP Unit view:\
  <img width="361" height="245" alt="abap2xlsx ABAP Unit Test in Eclipse ADT" src="https://github.com/user-attachments/assets/51912cde-4171-475d-b629-babc8816b092" />
- In the backend IDE, run the program `RS_AUCV_RUNNER`, select the main abap2xlsx package, run the tests by the context menu :\
  <img width="463" height="509" alt="abap2xlsx ABAP Unit RS_AUCV_RUNNER selection screen" src="https://github.com/user-attachments/assets/dc9d7c8e-4f7e-4ae3-be27-08d118bf125c" />\
  <img width="364" height="28" alt="abap2xlsx ABAP Unit RS_AUCV_RUNNER result" src="https://github.com/user-attachments/assets/543b5a1d-9669-471e-aab9-f85fb52a513a" />

These tests also run automatically when merging the pull requests, via [GitHub Actions and the abaplint transpiler](https://community.sap.com/t5/application-development-and-automation-blog-posts/running-abap-unit-tests-on-github-actions/ba-p/13496536). For information, you can see the past executions here: https://github.com/abap2xlsx/abap2xlsx/actions/workflows/unit.yml.

# Automated integration tests
These tests are based on the example programs proposed in the [demos](https://github.com/abap2xlsx/demos) repository. They make sure that the generated XLSX files are not changed by mistake.
Here is the principle:
- The generated XLSX files have been stored as reference in the `demos` repository as SAP Web Repository files (files with `w3mi` in their names).
- They are installed by abapGit at the same time as the demo programs. These Excel files can be seen via the transaction `SMW0` > Binary data > select the object names `ZDEMO_EXCEL*`. To see them, you may define the MIME type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` and associate it with your local Microsoft Excel application (e.g., the program `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`), via the menu Settings in `SMW0` > Binary data. This [blog post](https://community.sap.com/t5/technology-blog-posts-by-members/mime-types-truncated-in-sap-web-repository-smw0/ba-p/14229638) shows a few screenshots and may help if there's any question or issue.
- Run the program `ZDEMO_EXCEL_CHECKER` which is provided in the `demos` repository. It runs many demo programs and compare the generated XLSX files with the reference files stored in the SAP Web Repository.
  <img width="586" height="142" alt="ZDEMO_EXCEL_CHECKER selection screen" src="https://github.com/user-attachments/assets/ce85501b-c89a-4b21-b52f-314d8971377a" />
- The program shows which XLSX files are the same (green) or different (red). All should be green. A different file means either an intended change (new abap2xlsx feature or demo program changed) or a regression.
  <img width="835" height="193" alt="image" src="https://github.com/user-attachments/assets/5b156724-e213-41be-9fcf-2dd8b6af9fdf" />
- It's possible to see which files inside a given XLSX file are concerned (an XLSX file being a ZIP file, mainly composed of XML files). Also, VSCode can be installed to see the exact differences inside a given XML file.
  <img width="394" height="118" alt="ZDEMO_EXCEL_CHECKER different XLSX file" src="https://github.com/user-attachments/assets/94b8dc69-e7f0-4c7f-8279-2d0440c26c9d" />
  Click the button to view which XML files (and others) inside the XLSX file are different:
  <img width="19" height="18" alt="ZDEMO_EXCEL_CHECKER button 'View differences'" src="https://github.com/user-attachments/assets/71dbd40e-673a-4a94-b6b2-7fa279e3413e" />
  Click each XML file to see the differences:
  <img width="845" height="471" alt="ZDEMO_EXCEL_CHECKER view files inside XLSX" src="https://github.com/user-attachments/assets/86a10cce-086f-4c62-9763-c309aadb5014" />
  VSCode is started and shows the differences:
  <img width="1086" height="591" alt="VSCode to compare one XML file" src="https://github.com/user-attachments/assets/26320f63-5c9a-4329-bf32-b87ffe608338" />
- In the case the change is intended and as expected, the program has a "Save" button to replace the XLSX file in the SAP Web Repository with the one just generated.
  <img width="19" height="18" alt="ZDEMO_EXCEL_CHECKER button 'Save to SAP Web Repository'" src="https://github.com/user-attachments/assets/b467b7f4-54bc-4fc5-b9cb-3bd9aef5b943" />

# Syntax checks
Run abapGit:
<img width="1592" height="679" alt="image" src="https://github.com/user-attachments/assets/6229e3d9-4fdd-4c9d-89d8-8aeed10542f4" />
Select the abap2xlsx repository > Press "Check" > Select `SYNTAX_CHECK`:
<img width="1284" height="220" alt="image" src="https://github.com/user-attachments/assets/f4f87f18-9330-4058-9921-0c81348d8402" />
There should be no error and no warning.
