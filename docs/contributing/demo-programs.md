abap2xlsx is delivered with demo programs located in a dedicated folder. Currently, they are located in the same repository but there are [plans](https://github.com/abap2xlsx/abap2xlsx/issues/861) to dedicate a whole repository so that clients can install them outside the transport request of the main abap2xlsx code, to not have the demo programs in production.

Table of contents:
- [Program name and title](#program-name-and-title)
- [Include `ZDEMO_EXCEL_OUTPUTOPT_INCL`](#include-zdemo-excel-outputopt-incl)
- [Local class](#local-class)
- [Integration into general purpose programs](#integration-into-general-purpose-programs)
- [Most simple program](#most-simple-program)

<small><i><a href='http://ecotrust-canada.github.io/markdown-toc/'>Table of contents generated with markdown-toc</a></i></small>

<hr>

# Introduction

The demo programs can generally be used all these different ways:
- they can be called directly
- they can be called from the program `ZABAP2XLSX_DEMO_SHOW`
  - This program shows the list of the programs whose names start with the characters "ZDEMO_EXCEL" and whose ABAP code includes `ZDEMO_EXCEL_OUTPUTOPT_INCL`, along with their titles, and double-clicking any of them displays the produced Excel file (the workbook and its sheets) on the same screen.
- they can be called from the program `ZDEMO_EXCEL`
  - This program starts a predefined list of demo programs with default parameters to create the Excel files in a given folder.
- they can be called from the program `ZDEMO_EXCEL_CHECKER`
  - This program is used by abap2xlsx developers to detect regressions

# Program name and title

The names and titles of demo programs respect these rules:
- the name starts with the characters "ZDEMO_EXCEL" which are immediately followed with the next available number (e.g. `ZDEMO_EXCEL48`)
  - Some older programs currently don't use these conventions for some reason
- the program title is important because this text is currently displayed in `ZABAP2XLSX_DEMO_SHOW`, and they are normalized to be "abap2xlsx Demo: " followed by any clear and concise text written in "Title Case".

# Include `ZDEMO_EXCEL_OUTPUTOPT_INCL`

This include is used by most of the demo programs to:
- propose a selection screen with the below 4 output options:
  - Save to frontend
  - Save to backend
  - Direct display
  - Send via email
- the include contains a local class `LCL_OUTPUT` to display the produced Excel file:

Example:
```abap
CONSTANTS: gc_save_file_name TYPE string VALUE '01_HelloWorld.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.
  DATA lo_excel TYPE REF TO zcl_excel.
  ...
  lcl_output=>output( lo_excel ).
```

# Local class `LCL_EXCEL_GENERATOR`

The demo program should define a local class which inherits from `ZCL_EXCEL_GENERATOR`, named `LCL_EXCEL_GENERATOR`, so that it can be called by `ZDEMO_EXCEL_CHECKER`.

`ZCL_EXCEL_GENERATOR` defines 3 methods, only the 2 first ones are mandatory.
- Method `GET_INFORMATION`
  - It is to return the demo title (displayed by `ZDEMO_EXCEL_CHECKER`, and the plan is to have `ZABA2XLSX_DEMO_SHOW` use this name too),
  - an object ID which refers to the Excel file stored in the Web Repository (to be used by `ZDEMO_EXCEL_CHECKER` as a reference for detecting any regression), 
  - and the name of the Excel file stored in the Web Repository
- Method `GENERATE_EXCEL`
  - It is to generate and return the `ZCL_EXCEL` instance which contains the demonstration
- Method `GET_NEXT_GENERATOR`
  - It's an optional method that you may use if the demo program needs to produce several Excel files
  - It's currently redefined only by `ZDEMO_EXCEL15`.

# Integration into general purpose programs

- Run the program `ZABAP2XLSX_DEMO_SHOW` to check whether your demo program runs nicely
- In the program `ZDEMO_EXCEL`, add one line to call your demo program (here, this line concerns the program `ZDEMO_EXCEL1`):
  ```abap
    SUBMIT zdemo_excel1 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Hello world
  ```
- In the program `ZDEMO_EXCEL_CHECKER`, add one line in the method `LOAD_ALV_TABLE` to reference your local class (here, the one of program `ZDEMO_EXCEL1`):
  ```abap
  APPEND '\PROGRAM=ZDEMO_EXCEL1\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
  ```

# Most simple demo program

This code is taken from `ZDEMO_EXCEL1` and slightly simplified:
```abap
REPORT zdemo_excel1.

CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_excel_demo_generator.

  PUBLIC SECTION.
    METHODS zif_excel_demo_generator~get_information REDEFINITION.
    METHODS zif_excel_demo_generator~generate_excel REDEFINITION.

ENDCLASS.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_excel_generator TYPE REF TO lcl_excel_generator.

CONSTANTS: gc_save_file_name TYPE string VALUE '01_HelloWorld.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel_generator.
  lo_excel = lo_excel_generator->zif_excel_demo_generator~generate_excel( ).

*** Create output
  lcl_output=>output( lo_excel ).



CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD zif_excel_demo_generator~get_information.

    result-program = sy-repid.
    result-text = 'abap2xlsx Demo: Hello World'.
    result-filename = gc_save_file_name.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~generate_excel.

    DATA: lo_excel     TYPE REF TO zcl_excel,
          lo_worksheet TYPE REF TO zcl_excel_worksheet.

    CREATE OBJECT lo_excel.

    lo_worksheet = lo_excel->get_active_worksheet( ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
    ...

    result = lo_excel.

  ENDMETHOD.

ENDCLASS.
```
