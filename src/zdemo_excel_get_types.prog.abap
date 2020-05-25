*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL_GET_TYPES
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel_get_types.

TYPES : tt_text TYPE TABLE OF text80 .

DATA  : lo_excel     TYPE REF TO zcl_excel
      , reader          TYPE REF TO zif_excel_reader
      , lo_template_filler TYPE REF TO zcl_excel_fill_template
      .


PARAMETERS: p_fpath TYPE string OBLIGATORY LOWER CASE DEFAULT 'C:\Users\sadfasdf\Desktop\abap2xlsx\ZABAP2XLSX_EXAMPLE.xlsx'.

PARAMETERS: p_normal RADIOBUTTON GROUP rad1 DEFAULT 'X'
          , p_other RADIOBUTTON GROUP rad1
          .

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_fpath.
  PERFORM get_file_path CHANGING p_fpath.


START-OF-SELECTION.

  CREATE OBJECT reader TYPE zcl_excel_reader_2007.
  lo_excel = reader->load_file( p_fpath ).

  CREATE OBJECT lo_template_filler .

  lo_template_filler->get_range( lo_excel ).
  lo_template_filler->discard_overlapped( ).
  lo_template_filler->sign_range( ).
  lo_template_filler->find_var( lo_excel ).

  PERFORM get_types.


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

FORM get_types .

  DATA
        : lv_sum TYPE i
        , lt_res TYPE tt_text
        , lt_buf TYPE tt_text
        .

  FIELD-SYMBOLS
                 : <fs_sheet> LIKE LINE OF lo_template_filler->mt_sheet
                 .
  LOOP AT lo_template_filler->mt_sheet ASSIGNING <fs_sheet>.

    CLEAR lv_sum.

    READ TABLE lo_template_filler->mt_range TRANSPORTING NO FIELDS WITH KEY sheet = <fs_sheet>.

    ADD sy-subrc TO lv_sum.

    READ TABLE lo_template_filler->mt_var TRANSPORTING NO FIELDS WITH KEY sheet = <fs_sheet>.

    ADD sy-subrc TO lv_sum.

    CHECK lv_sum <= 4.

    PERFORM get_type_r USING <fs_sheet>   0    CHANGING lt_buf.

    APPEND LINES OF lt_buf TO lt_res.


  ENDLOOP.

  DATA
        : lv_lines TYPE i
        .

  FIELD-SYMBOLS
                 : <fs_res> TYPE text80
                 .

  IF p_normal IS INITIAL.
    READ TABLE lt_res ASSIGNING <fs_res> INDEX 1.
    TRANSLATE <fs_res> USING ',:'.
    INSERT INITIAL LINE INTO lt_res ASSIGNING <fs_res> INDEX 1.
    <fs_res> = 'TYPES'.
    APPEND INITIAL LINE TO lt_res ASSIGNING <fs_res>.
    <fs_res> = '.'.

    lv_lines = lines( lt_res ) - 2.
    DELETE lt_res INDEX lv_lines.
    DELETE lt_res INDEX lv_lines.

  ELSE.
    INSERT INITIAL LINE INTO lt_res ASSIGNING <fs_res> INDEX 1.
    <fs_res> = 'TYPES:'.

    lv_lines = lines( lt_res )  - 2.

    READ TABLE lt_res ASSIGNING <fs_res> INDEX lv_lines.
    TRANSLATE <fs_res> USING ',.'.
    ADD 1 TO lv_lines.

  ENDIF.

  IF p_normal IS INITIAL.
    APPEND INITIAL LINE TO lt_res ASSIGNING <fs_res>.
    <fs_res> = 'DATA'.
    APPEND INITIAL LINE TO lt_res ASSIGNING <fs_res>.
    <fs_res> = ': lo_data type ref to ZCL_EXCEL_TEMPLATE_DATA'.
    APPEND INITIAL LINE TO lt_res ASSIGNING <fs_res>.
    <fs_res> = '.'.

  ELSE.
    APPEND INITIAL LINE TO lt_res ASSIGNING <fs_res>.
    <fs_res> = 'DATA: lo_data type ref to ZCL_EXCEL_TEMPLATE_DATA.'.
  ENDIF.

  cl_demo_output=>new( 'TEXT' )->display( lt_res ).
ENDFORM.


FORM get_type_r USING p_sheet TYPE zexcel_template_sheet_title
                      p_parent  TYPE i
                CHANGING ct_result TYPE tt_text.

  CLEAR ct_result.

  DATA
        : lt_buf TYPE tt_text
        , lt_tmp TYPE tt_text
        , lv_sum TYPE i
        .

  FIELD-SYMBOLS
                 : <fs_buf> TYPE text80
                 .

  FIELD-SYMBOLS
                 : <fs_range> LIKE LINE OF lo_template_filler->mt_range
                 .


  LOOP AT lo_template_filler->mt_range ASSIGNING <fs_range> WHERE sheet = p_sheet
                                                        AND parent = p_parent.

    PERFORM get_type_r
                USING
                   p_sheet
                   <fs_range>-id
                CHANGING
                   lt_tmp.

    APPEND LINES OF lt_tmp TO lt_buf.


  ENDLOOP.

  ADD sy-subrc TO lv_sum.

  DATA
        : lv_name TYPE string
        .

  IF p_parent = 0.
    lv_name = p_sheet.
  ELSE.
    READ TABLE lo_template_filler->mt_range ASSIGNING <fs_range> WITH KEY sheet = p_sheet
                                                      id = p_parent.
    lv_name = <fs_range>-name.
  ENDIF.

  APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.
  IF p_normal IS INITIAL.
    CONCATENATE ', begin of t_'   lv_name INTO <fs_buf>.
  ELSE.
    CONCATENATE ' begin of t_' lv_name ',' INTO <fs_buf>.
  ENDIF.

  FIELD-SYMBOLS
                 : <fs_var> type ZEXCEL_TEMPLATE_S_VAR
                 .

  LOOP AT lo_template_filler->mt_var ASSIGNING <fs_var> WHERE sheet = p_sheet
                                                    AND parent = p_parent.

    APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.
    IF p_normal IS INITIAL.
      CONCATENATE ',     '  <fs_var>-name ' type string' INTO <fs_buf> RESPECTING BLANKS.
    ELSE.
      CONCATENATE '     '  <fs_var>-name ' type string,' INTO <fs_buf> RESPECTING BLANKS.
    ENDIF.


  ENDLOOP.

  ADD sy-subrc TO lv_sum.

  LOOP AT lo_template_filler->mt_range ASSIGNING <fs_range> WHERE sheet = p_sheet
                                                        AND parent = p_parent.

    APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.
    IF p_normal IS INITIAL.
      CONCATENATE ',     ' <fs_range>-name ' type tt_' <fs_range>-name INTO <fs_buf> RESPECTING BLANKS .
    ELSE.
      CONCATENATE '     ' <fs_range>-name ' type tt_' <fs_range>-name ',' INTO <fs_buf> RESPECTING BLANKS .
    ENDIF.


  ENDLOOP.

  IF lv_sum > 4.
    APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.
    IF p_normal IS INITIAL.
      <fs_buf> = ',     xz type i'.
    ELSE.
      <fs_buf> = '     xz type i,'.
    ENDIF.

  ENDIF.


  APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.
  IF p_normal IS INITIAL.
    CONCATENATE ', end of t_' lv_name INTO <fs_buf>.
  ELSE.
    CONCATENATE ' end of t_' lv_name ',' INTO <fs_buf>.
  ENDIF.

  APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.

  IF p_parent NE 0.
    APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.
    IF p_normal IS INITIAL.
      CONCATENATE ', tt_' lv_name ' type t_' lv_name  ' OCCURS 0' INTO <fs_buf> RESPECTING BLANKS .
    ELSE.
      CONCATENATE ' tt_' lv_name ' type t_' lv_name   ' OCCURS 0,' INTO <fs_buf> RESPECTING BLANKS .
    ENDIF.

  ENDIF.

  APPEND INITIAL LINE TO lt_buf ASSIGNING <fs_buf>.


  ct_result = lt_buf.
ENDFORM.
