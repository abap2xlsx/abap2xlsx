*&---------------------------------------------------------------------*
*& Report  Fill Template
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel_fill_template.

TYPES:  BEGIN OF t_table1,
    person TYPE string,
    salary TYPE i.
TYPES:  END OF t_table1.

  TYPES: tt_table1 TYPE  t_table1 OCCURS 0.

TYPES:  BEGIN OF t_line1,
    carrid TYPE string,
    connid TYPE string,
    fldate TYPE string,
    price  TYPE i.
TYPES:  END OF t_line1.

TYPES:  tt_line1 TYPE  t_line1 OCCURS 0 .

TYPES:  BEGIN OF t_table2,
    carrid TYPE string,
    price  TYPE i,
    line1  TYPE tt_line1.
TYPES:  END OF t_table2.

TYPES:  tt_table2 TYPE   t_table2 OCCURS 0.

TYPES:  BEGIN OF t_sheet1,
    date   TYPE string,
    time   TYPE string,
    user   TYPE string,
    total  TYPE i,
    price  TYPE i,
    table1 TYPE tt_table1,
    table2 TYPE tt_table2.
TYPES:  END OF t_sheet1.

TYPES:  BEGIN OF t_table3,
    person TYPE string,
    salary TYPE string.
TYPES:  END OF t_table3.

TYPES:  tt_table3 TYPE t_table3 OCCURS 0.

TYPES:  BEGIN OF t_sheet2,
    date   TYPE string,
    time   TYPE string,
    user   TYPE string,
    total  TYPE i,
    table3 TYPE tt_table1.
TYPES:  END OF t_sheet2.


FIELD-SYMBOLS
               : <fs_table1> TYPE t_table1
               , <fs_line> TYPE t_line1
               , <fs_table2> TYPE t_table2
               .

DATA
: lo_data TYPE REF TO zcl_excel_template_data
, gs_sheet1 TYPE   t_sheet1
, gs_sheet2 TYPE   t_sheet2
.

* define variables
DATA: lo_excel TYPE REF TO zcl_excel,
      reader   TYPE REF TO zif_excel_reader.


CONSTANTS: gc_save_file_name TYPE string VALUE 'fill_template_example.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_fpath TYPE string OBLIGATORY LOWER CASE DEFAULT 'C:\Users\sadfasdf\Desktop\abap2xlsx\ZABAP2XLSX_EXAMPLE.xlsx'.


PARAMETERS: p_file RADIOBUTTON GROUP rad1 DEFAULT 'X'
          , p_smw0 RADIOBUTTON GROUP rad1
          .

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_fpath.
  PERFORM get_file_path CHANGING p_fpath.


START-OF-SELECTION.

  CREATE OBJECT lo_data.

data
      : lv_date TYPE char10
      , lv_time TYPE char8
      .

  WRITE sy-datum TO lv_date.
  gs_sheet1-date = lv_date.
  WRITE sy-uzeit TO lv_time.
  gs_sheet1-time = lv_time.

  gs_sheet1-user = sy-uname.
  gs_sheet1-total = '5600'.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <fs_table1>.
  <fs_table1>-person = 'Lurch Schpellchek'.
  <fs_table1>-salary = '1200'.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <fs_table1>.
  <fs_table1>-person = 'Russell Sprout'.
  <fs_table1>-salary = '1300'.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <fs_table1>.
  <fs_table1>-person = 'Fergus Douchebag'.
  <fs_table1>-salary = '3000'.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <fs_table1>.
  <fs_table1>-person = 'Bartholomew Shoe'.
  <fs_table1>-salary = '100'.


  gs_sheet1-price = '14003'.

  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <fs_table2>.
  <fs_table2>-carrid ='AC'.
  <fs_table2>-price ='1222'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'AC'.
  <fs_line>-connid = '0820'.
  <fs_line>-fldate = '20.12.2002'.
  <fs_line>-price = '1222'.


  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <fs_table2>.
  <fs_table2>-carrid ='AF'.
  <fs_table2>-price ='2222'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'AF'.
  <fs_line>-connid = '0820'.
  <fs_line>-fldate = '23.12.2002'.
  <fs_line>-price = '2222'.


  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <fs_table2>.
  <fs_table2>-carrid ='LH'.
  <fs_table2>-price ='9488'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'LH'.
  <fs_line>-connid = '0400'.
  <fs_line>-fldate = '28.02.1995'.
  <fs_line>-price = '899'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'LH'.
  <fs_line>-connid = '0400'.
  <fs_line>-fldate = '17.11.1995'.
  <fs_line>-price = '1499'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'LH'.
  <fs_line>-connid = '0400'.
  <fs_line>-fldate = '06.06.1995'.
  <fs_line>-price = '1090'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'LH'.
  <fs_line>-connid = '0400'.
  <fs_line>-fldate = '28.04.1995'.
  <fs_line>-price = '6000'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'LH'.
  <fs_line>-connid = '0400'.
  <fs_line>-fldate = '21.12.2002'.
  <fs_line>-price = '222'.

  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <fs_table2>.
  <fs_table2>-carrid ='SQ'.
  <fs_table2>-price ='849'.

  APPEND INITIAL LINE TO <fs_table2>-line1 ASSIGNING <fs_line>.
  <fs_line>-carrid = 'SQ'.
  <fs_line>-connid = '0026'.
  <fs_line>-fldate = '28.02.1995'.
  <fs_line>-price = '849'.


  MOVE-CORRESPONDING gs_sheet1 TO gs_sheet2.
  gs_sheet2-table3 = gs_sheet1-table1.

* generate data

* add data
  lo_data->add( iv_sheet = 'Sheet1' iv_data = gs_sheet1 ).
  lo_data->add( iv_sheet = 'Sheet2' iv_data = gs_sheet2 ).

* create reader

  CREATE OBJECT reader TYPE zcl_excel_reader_2007.
* load template

  IF p_file IS NOT INITIAL.
    lo_excel = reader->load_file( p_fpath ).
  ELSE.
    lo_excel = reader->load_smw0( 'ZEXCEL_DEMO_TEMPLATE' ).
  ENDIF.

* merge data with template
  lo_excel->fill_template( lo_data ).

*** Create output
  lcl_output=>output( lo_excel ).


*&---------------------------------------------------------------------*
*&      Form  Get_file_path
*&---------------------------------------------------------------------*
FORM get_file_path CHANGING cv_path TYPE string.
  CLEAR cv_path.

  DATA:
    lv_rc          TYPE  i,
    lv_user_action TYPE  i,
    lt_file_table  TYPE  filetable,
    ls_file_table  LIKE LINE OF lt_file_table.

  cl_gui_frontend_services=>file_open_dialog(
  EXPORTING
    window_title        = 'select template  xlsx'
    multiselection      = ''
    default_extension   = '*.xlsx'
    file_filter         = 'Text file (*.xlsx)|*.xlsx|All (*.*)|*.*'
  CHANGING
    file_table          = lt_file_table
    rc                  = lv_rc
    user_action         = lv_user_action
  EXCEPTIONS
    OTHERS              = 1
    ).
  IF sy-subrc = 0.
    IF lv_user_action = cl_gui_frontend_services=>action_ok.
      IF lt_file_table IS NOT INITIAL.
        READ TABLE lt_file_table INTO ls_file_table INDEX 1.
        IF sy-subrc = 0.
          cv_path = ls_file_table-filename.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.
ENDFORM.                    " Get_file_path
