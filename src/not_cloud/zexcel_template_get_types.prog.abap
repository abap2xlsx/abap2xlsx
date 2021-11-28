*&---------------------------------------------------------------------*
*& Report  ZEXCEL_TEMPLATE_GET_TYPES
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zexcel_template_get_types.

TYPES : tt_text TYPE TABLE OF text80.

DATA: go_excel           TYPE REF TO zcl_excel,
      go_reader          TYPE REF TO zif_excel_reader,
      go_template_filler TYPE REF TO zcl_excel_fill_template,
      go_error           TYPE REF TO zcx_excel.

SELECTION-SCREEN BEGIN OF BLOCK b02 WITH FRAME.

PARAMETERS: p_smw0 RADIOBUTTON GROUP rad2 DEFAULT 'X'.
PARAMETERS: p_objid TYPE w3objid OBLIGATORY DEFAULT 'ZDEMO_EXCEL_TEMPLATE'.

PARAMETERS: p_file RADIOBUTTON GROUP rad2.
PARAMETERS: p_fpath TYPE string OBLIGATORY LOWER CASE DEFAULT 'c:\temp\whatever.xlsx'.

SELECTION-SCREEN END OF BLOCK b02.

SELECTION-SCREEN BEGIN OF BLOCK b01 WITH FRAME.

PARAMETERS: p_normal RADIOBUTTON GROUP rad1 DEFAULT 'X',
            p_other  RADIOBUTTON GROUP rad1.

SELECTION-SCREEN END OF BLOCK b01.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_fpath.
  PERFORM get_file_path CHANGING p_fpath.


START-OF-SELECTION.

  TRY.

      CREATE OBJECT go_reader TYPE zcl_excel_reader_2007.

* load template
      IF p_file IS NOT INITIAL.
        go_excel = go_reader->load_file( p_fpath ).
      ELSE.
        PERFORM load_smw0 USING go_reader p_objid CHANGING go_excel.
      ENDIF.

      go_template_filler = zcl_excel_fill_template=>create( go_excel ).

      PERFORM get_types.

    CATCH zcx_excel INTO go_error.
      MESSAGE go_error TYPE 'I' DISPLAY LIKE 'E'.
  ENDTRY.


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
        OTHERS              = 1 ).
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

FORM get_types .

  DATA: lv_sum   TYPE i,
        lt_res   TYPE tt_text,
        lt_buf   TYPE tt_text,
        lv_lines TYPE i.

  FIELD-SYMBOLS: <ls_sheet> LIKE LINE OF go_template_filler->mt_sheet,
                 <lv_res>   TYPE text80.


  LOOP AT go_template_filler->mt_sheet ASSIGNING <ls_sheet>.

    CLEAR lv_sum.

    READ TABLE go_template_filler->mt_range TRANSPORTING NO FIELDS WITH KEY sheet = <ls_sheet>.

    ADD sy-subrc TO lv_sum.

    READ TABLE go_template_filler->mt_var TRANSPORTING NO FIELDS WITH KEY sheet = <ls_sheet>.

    ADD sy-subrc TO lv_sum.

    CHECK lv_sum <= 4.

    PERFORM get_type_r USING <ls_sheet>   0    CHANGING lt_buf.

    APPEND LINES OF lt_buf TO lt_res.

  ENDLOOP.


  IF p_normal IS INITIAL.
    READ TABLE lt_res ASSIGNING <lv_res> INDEX 1.
    TRANSLATE <lv_res> USING ',:'.
    INSERT INITIAL LINE INTO lt_res ASSIGNING <lv_res> INDEX 1.
    <lv_res> = 'TYPES'.
    APPEND INITIAL LINE TO lt_res ASSIGNING <lv_res>.
    <lv_res> = '.'.

    lv_lines = lines( lt_res ) - 2.
    DELETE lt_res INDEX lv_lines.
    DELETE lt_res INDEX lv_lines.

  ELSE.
    INSERT INITIAL LINE INTO lt_res ASSIGNING <lv_res> INDEX 1.
    <lv_res> = 'TYPES:'.

    lv_lines = lines( lt_res )  - 2.

    READ TABLE lt_res ASSIGNING <lv_res> INDEX lv_lines.
    TRANSLATE <lv_res> USING ',.'.
    ADD 1 TO lv_lines.

  ENDIF.

  INSERT 'TYPES t_number TYPE p length 16 decimals 4.' INTO lt_res INDEX 1.

  IF p_normal IS INITIAL.
    APPEND INITIAL LINE TO lt_res ASSIGNING <lv_res>.
    <lv_res> = 'DATA'.
    APPEND INITIAL LINE TO lt_res ASSIGNING <lv_res>.
    <lv_res> = ': lo_data type ref to ZCL_EXCEL_TEMPLATE_DATA'.
    APPEND INITIAL LINE TO lt_res ASSIGNING <lv_res>.
    <lv_res> = '.'.

  ELSE.
    APPEND INITIAL LINE TO lt_res ASSIGNING <lv_res>.
    <lv_res> = 'DATA: lo_data type ref to ZCL_EXCEL_TEMPLATE_DATA.'.
  ENDIF.

  cl_demo_output=>new( 'TEXT' )->display( lt_res ).
ENDFORM.


FORM get_type_r USING iv_sheet  TYPE zexcel_sheet_title
                      iv_parent TYPE i
                CHANGING ct_result TYPE tt_text.

  DATA: lt_buf             TYPE tt_text,
        lt_tmp             TYPE tt_text,
        lv_sum             TYPE i,
        lv_name            TYPE string,
        lv_type            TYPE string,
        lv_string          TYPE string,
        lt_sorted_counters TYPE TABLE OF i,
        lv_biggest_counter TYPE i.

  FIELD-SYMBOLS: <lv_buf>        TYPE text80,
                 <ls_range>      LIKE LINE OF go_template_filler->mt_range,
                 <ls_var>        TYPE zcl_excel_fill_template=>ts_variable,
                 <ls_name_style> TYPE zcl_excel_fill_template=>ts_name_style.


  CLEAR ct_result.

  LOOP AT go_template_filler->mt_range ASSIGNING <ls_range> WHERE sheet  = iv_sheet
                                                           AND parent = iv_parent.

    PERFORM get_type_r
                USING
                   iv_sheet
                   <ls_range>-id
                CHANGING
                   lt_tmp.

    APPEND LINES OF lt_tmp TO lt_buf.


  ENDLOOP.

  ADD sy-subrc TO lv_sum.

  IF iv_parent = 0.
    lv_name = iv_sheet.
  ELSE.
    READ TABLE go_template_filler->mt_range ASSIGNING <ls_range>
        WITH KEY sheet = iv_sheet
                 id    = iv_parent.
    lv_name = <ls_range>-name.
  ENDIF.

  APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.
  IF p_normal IS INITIAL.
    CONCATENATE ', begin of t_'   lv_name INTO <lv_buf>.
  ELSE.
    CONCATENATE ' begin of t_' lv_name ',' INTO <lv_buf>.
  ENDIF.


  LOOP AT go_template_filler->mt_var ASSIGNING <ls_var> WHERE sheet  = iv_sheet
                                                       AND parent = iv_parent.

    READ TABLE go_template_filler->mt_name_styles
        WITH KEY sheet  = iv_sheet
                 name   = <ls_var>-name
                 parent = iv_parent
        ASSIGNING <ls_name_style>.
    IF sy-subrc <> 0.
      lv_type = 'string'.
    ELSE.
      CLEAR lt_sorted_counters.
      APPEND <ls_name_style>-numeric_counter TO lt_sorted_counters.
      APPEND <ls_name_style>-date_counter TO lt_sorted_counters.
      APPEND <ls_name_style>-time_counter TO lt_sorted_counters.
      APPEND <ls_name_style>-text_counter TO lt_sorted_counters.
      SORT lt_sorted_counters BY table_line DESCENDING.
      READ TABLE lt_sorted_counters INDEX 1 INTO lv_biggest_counter.
      ASSERT sy-subrc = 0.
      CASE lv_biggest_counter.
        WHEN <ls_name_style>-numeric_counter.
          lv_type = 't_number'.
        WHEN <ls_name_style>-date_counter.
          lv_type = 'd'.
        WHEN <ls_name_style>-time_counter.
          lv_type = 't'.
        WHEN <ls_name_style>-text_counter.
          lv_type = 'string'.
      ENDCASE.
    ENDIF.

    APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.
    IF p_normal IS INITIAL.
      CONCATENATE ',     '  <ls_var>-name ' type ' lv_type INTO <lv_buf> RESPECTING BLANKS.
    ELSE.
      CONCATENATE '     '  <ls_var>-name ' type ' lv_type ',' INTO <lv_buf> RESPECTING BLANKS.
    ENDIF.


  ENDLOOP.

  ADD sy-subrc TO lv_sum.

  LOOP AT go_template_filler->mt_range ASSIGNING <ls_range> WHERE sheet  = iv_sheet
                                                           AND parent = iv_parent.

    APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.
    lv_string = <ls_range>-name.
    IF p_normal IS INITIAL.
      CONCATENATE ',     ' <ls_range>-name ' type tt_' lv_string INTO <lv_buf> RESPECTING BLANKS .
    ELSE.
      CONCATENATE '     ' <ls_range>-name ' type tt_' lv_string ',' INTO <lv_buf> RESPECTING BLANKS .
    ENDIF.


  ENDLOOP.

  IF lv_sum > 4.
    APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.
    IF p_normal IS INITIAL.
      <lv_buf> = ',     xz type i'.
    ELSE.
      <lv_buf> = '     xz type i,'.
    ENDIF.

  ENDIF.


  APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.
  IF p_normal IS INITIAL.
    CONCATENATE ', end of t_' lv_name INTO <lv_buf>.
  ELSE.
    CONCATENATE ' end of t_' lv_name ',' INTO <lv_buf>.
  ENDIF.

  APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.

  IF iv_parent NE 0.
    APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.
    IF p_normal IS INITIAL.
      CONCATENATE ', tt_' lv_name ' type standard table of t_' lv_name  ' with default key' INTO <lv_buf> RESPECTING BLANKS .
    ELSE.
      CONCATENATE ' tt_' lv_name ' type standard table of t_' lv_name   ' with default key,' INTO <lv_buf> RESPECTING BLANKS .
    ENDIF.

  ENDIF.

  APPEND INITIAL LINE TO lt_buf ASSIGNING <lv_buf>.


  ct_result = lt_buf.
ENDFORM.

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
