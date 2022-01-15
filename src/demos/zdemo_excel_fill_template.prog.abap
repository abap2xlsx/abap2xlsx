*&---------------------------------------------------------------------*
*& Report  Fill Template
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel_fill_template.

*=================
* Start of generated code.
* All these types were generated
* by ZEXCEL_TEMPLATE_GET_TYPES based
* on the Excel file ZDEMO_EXCEL_TEMPLATE
* from SMW0.
*=================
TYPES t_number TYPE p LENGTH 16 DECIMALS 4.
TYPES:
  BEGIN OF t_table1,
    person TYPE string,
    salary TYPE t_number,
  END OF t_table1,

  tt_table1 TYPE STANDARD TABLE OF t_table1 WITH DEFAULT KEY,

  BEGIN OF t_line1,
    carrid TYPE string,
    connid TYPE string,
    fldate TYPE d,
    price  TYPE t_number,
  END OF t_line1,

  tt_line1 TYPE STANDARD TABLE OF t_line1 WITH DEFAULT KEY,

  BEGIN OF t_table2,
    carrid TYPE string,
    price  TYPE t_number,
    line1  TYPE tt_line1,
  END OF t_table2,

  tt_table2 TYPE STANDARD TABLE OF t_table2 WITH DEFAULT KEY,

  BEGIN OF t_sheet1,
    date   TYPE d,
    time   TYPE t,
    user   TYPE string,
    total  TYPE t_number,
    price  TYPE t_number,
    table1 TYPE tt_table1,
    table2 TYPE tt_table2,
  END OF t_sheet1,


  BEGIN OF t_table3,
    person TYPE string,
    salary TYPE t_number,
  END OF t_table3,

  tt_table3 TYPE STANDARD TABLE OF t_table3 WITH DEFAULT KEY,

  BEGIN OF t_sheet2,
    date   TYPE d,
    time   TYPE t,
    user   TYPE string,
    total  TYPE t_number,
    table3 TYPE tt_table3,
  END OF t_sheet2.


DATA: lo_data TYPE REF TO zcl_excel_template_data.
*=================
* End of generated code
*=================

* define variables
DATA: gs_sheet1 TYPE t_sheet1,
      gs_sheet2 TYPE t_sheet2.

TABLES: sscrfields.

CONSTANTS: gc_save_file_name TYPE string VALUE 'fill_template.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_smw0 RADIOBUTTON GROUP rad1 DEFAULT 'X'.
PARAMETERS: p_objid TYPE w3objid OBLIGATORY DEFAULT 'ZDEMO_EXCEL_TEMPLATE'.

PARAMETERS: p_file RADIOBUTTON GROUP rad1.
PARAMETERS: p_fpath TYPE string OBLIGATORY LOWER CASE DEFAULT 'c:\temp\whatever.xlsx'.

SELECTION-SCREEN SKIP 1.

SELECTION-SCREEN PUSHBUTTON /1(45) but_txt USER-COMMAND get_types.


INITIALIZATION.
  but_txt = '@BY@ Analyze file to propose TYPES'.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_fpath.
  PERFORM get_file_path CHANGING p_fpath.

AT SELECTION-SCREEN.
  CASE sscrfields-ucomm.
    WHEN 'GET_TYPES'.
      SUBMIT zexcel_template_get_types
        WITH p_smw0  = p_smw0
        WITH p_objid = p_objid
        WITH p_file  = p_file
        WITH p_fpath = p_fpath
        AND RETURN.
  ENDCASE.

START-OF-SELECTION.
  PERFORM load_data.
  PERFORM generate_file.

FORM load_data.

  FIELD-SYMBOLS: <ls_table1> TYPE t_table1,
                 <ls_line>   TYPE t_line1,
                 <lt_table2> TYPE t_table2.

  gs_sheet1-date = sy-datum.
  gs_sheet1-time = sy-uzeit.
  gs_sheet1-user = sy-uname.
  gs_sheet1-total = 5600.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <ls_table1>.
  <ls_table1>-person = 'Lurch Schpellchek'.
  <ls_table1>-salary = 1200.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <ls_table1>.
  <ls_table1>-person = 'Russell Sprout'.
  <ls_table1>-salary = 1300.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <ls_table1>.
  <ls_table1>-person = 'Fergus Douchebag'.
  <ls_table1>-salary = 3000.

  APPEND INITIAL LINE TO gs_sheet1-table1 ASSIGNING <ls_table1>.
  <ls_table1>-person = 'Bartholomew Shoe'.
  <ls_table1>-salary = 100.


  gs_sheet1-price = 14003.

  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <lt_table2>.
  <lt_table2>-carrid = 'AC'.
  <lt_table2>-price = 1222.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'AC'.
  <ls_line>-connid = '0820'.
  <ls_line>-fldate = '20021220'.
  <ls_line>-price = 1222.


  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <lt_table2>.
  <lt_table2>-carrid = 'AF'.
  <lt_table2>-price = 2222.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'AF'.
  <ls_line>-connid = '0820'.
  <ls_line>-fldate = '20021223'.
  <ls_line>-price = 2222.


  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <lt_table2>.
  <lt_table2>-carrid = 'LH'.
  <lt_table2>-price = 9488.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'LH'.
  <ls_line>-connid = '0400'.
  <ls_line>-fldate = '19950228'.
  <ls_line>-price = 899.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'LH'.
  <ls_line>-connid = '0400'.
  <ls_line>-fldate = '19951117'.
  <ls_line>-price = 1499.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'LH'.
  <ls_line>-connid = '0400'.
  <ls_line>-fldate = '19950606'.
  <ls_line>-price = 1090.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'LH'.
  <ls_line>-connid = '0400'.
  <ls_line>-fldate = '19950428'.
  <ls_line>-price = 6000.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'LH'.
  <ls_line>-connid = '0400'.
  <ls_line>-fldate = '20021221'.
  <ls_line>-price = 222.

  APPEND INITIAL LINE TO gs_sheet1-table2 ASSIGNING <lt_table2>.
  <lt_table2>-carrid = 'SQ'.
  <lt_table2>-price = 849.

  APPEND INITIAL LINE TO <lt_table2>-line1 ASSIGNING <ls_line>.
  <ls_line>-carrid = 'SQ'.
  <ls_line>-connid = '0026'.
  <ls_line>-fldate = '19950228'.
  <ls_line>-price = 849.


  MOVE-CORRESPONDING gs_sheet1 TO gs_sheet2.
  gs_sheet2-table3 = gs_sheet1-table1.

ENDFORM.

FORM generate_file.

  DATA: lo_excel  TYPE REF TO zcl_excel,
        lo_reader TYPE REF TO zif_excel_reader,
        lo_root   TYPE REF TO cx_root.

  TRY.

* prepare data
      CREATE OBJECT lo_data.
      lo_data->add( iv_sheet = 'Sheet1' iv_data = gs_sheet1 ).
      lo_data->add( iv_sheet = 'Sheet2' iv_data = gs_sheet2 ).

* create reader
      CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.

* load template
      IF p_file IS NOT INITIAL.
        lo_excel = lo_reader->load_file( p_fpath ).
      ELSE.
        PERFORM load_smw0 USING lo_reader p_objid CHANGING lo_excel.
      ENDIF.

* merge data with template
      lo_excel->fill_template( lo_data ).


*** Create output
      lcl_output=>output( cl_excel = lo_excel iv_info_message = abap_false ).

    CATCH cx_root INTO lo_root.
      MESSAGE lo_root TYPE 'I' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  Get_file_path
*&---------------------------------------------------------------------*
FORM get_file_path CHANGING cv_path TYPE string.

  DATA:
    lv_rc          TYPE  i,
    lv_user_action TYPE  i,
    lt_file_table  TYPE  filetable,
    ls_file_table  LIKE LINE OF lt_file_table.

  CLEAR cv_path.

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

FORM load_smw0
    USING
        io_reader  TYPE REF TO zif_excel_reader
        iv_w3objid TYPE w3objid
    CHANGING
        ro_excel   TYPE REF TO zcl_excel
    RAISING
        zcx_excel.

  DATA: lv_excel_data   TYPE xstring,
        lt_mime         TYPE TABLE OF w3mime,
        ls_key          TYPE wwwdatatab,
        lv_errormessage TYPE string,
        lv_filesize     TYPE i,
        lv_filesizec    TYPE c LENGTH 10.

*--------------------------------------------------------------------*
* Read file into binary string
*--------------------------------------------------------------------*

  ls_key-relid = 'MI'.
  ls_key-objid = iv_w3objid .

  CALL FUNCTION 'WWWDATA_IMPORT'
    EXPORTING
      key    = ls_key
    TABLES
      mime   = lt_mime
    EXCEPTIONS
      OTHERS = 1.
  IF sy-subrc <> 0.
    lv_errormessage = 'A problem occured when reading the MIME object'(004).
    zcx_excel=>raise_text( lv_errormessage ).
  ENDIF.

  CALL FUNCTION 'WWWPARAMS_READ'
    EXPORTING
      relid = ls_key-relid
      objid = ls_key-objid
      name  = 'filesize'
    IMPORTING
      value = lv_filesizec.

  lv_filesize = lv_filesizec.
  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lv_filesize
    IMPORTING
      buffer       = lv_excel_data
    TABLES
      binary_tab   = lt_mime
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.

*--------------------------------------------------------------------*
* Parse Excel data into ZCL_EXCEL object from binary string
*--------------------------------------------------------------------*
  ro_excel = io_reader->load( i_excel2007 = lv_excel_data ).

ENDFORM.
