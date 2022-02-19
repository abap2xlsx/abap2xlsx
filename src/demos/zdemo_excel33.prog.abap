*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel33.

TYPES: ty_t005t_lines TYPE TABLE OF t005t.

DATA: lo_excel      TYPE REF TO zcl_excel,
      lo_worksheet  TYPE REF TO zcl_excel_worksheet,
      lo_converter  TYPE REF TO zcl_excel_converter,
      lo_autofilter TYPE REF TO zcl_excel_autofilter.

DATA lt_test TYPE ty_t005t_lines.

DATA: l_cell_value TYPE zexcel_cell_value,
      ls_area      TYPE zexcel_s_autofilter_area.
DATA: ls_option TYPE zexcel_s_converter_option.

CONSTANTS: c_airlines TYPE string VALUE 'Airlines'.

CONSTANTS: gc_save_file_name TYPE string VALUE '33_autofilter.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS p_convex AS CHECKBOX.

START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Internal table' ).

  PERFORM load_fixed_data CHANGING lt_test.

  CREATE OBJECT lo_converter.

  ls_option-conv_exit_length = p_convex.
  lo_converter->set_option( ls_option ).
  lo_converter->convert( EXPORTING
                            it_table     = lt_test
                            i_row_int    = 1
                            i_column_int = 1
                            io_worksheet = lo_worksheet
                         CHANGING
                            co_excel     = lo_excel ) .
  PERFORM set_column_headers USING lo_worksheet 'Client;Language;Country;Name;Nationality;Long name;Nationality'.

  lo_autofilter = lo_excel->add_new_autofilter( io_sheet = lo_worksheet ) .

  ls_area-row_start = 1.
  ls_area-col_start = 1.
  ls_area-row_end = lo_worksheet->get_highest_row( ).
  ls_area-col_end = lo_worksheet->get_highest_column( ).

  lo_autofilter->set_filter_area( is_area = ls_area ).

  lo_worksheet->get_cell( EXPORTING
                             ip_column    = 'C'
                             ip_row       = 2
                          IMPORTING
                             ep_value     = l_cell_value ).
  lo_autofilter->set_value( i_column = 3
                            i_value  = l_cell_value ).


*** Create output
  lcl_output=>output( lo_excel ).


FORM load_fixed_data CHANGING ct_test TYPE ty_t005t_lines.
  DATA: lt_lines  TYPE TABLE OF string,
        lv_line   TYPE string,
        lt_fields TYPE TABLE OF string,
        lv_comp   TYPE i,
        lv_field  TYPE string,
        ls_test   TYPE t005t.
  FIELD-SYMBOLS: <lv_field> TYPE simple.

  APPEND '001 E AD Andorra    Andorran    Andorra    Andorran   ' TO lt_lines.
  APPEND '001 E BE Belgium    Belgian     Belgium    Belgian    ' TO lt_lines.
  APPEND '001 E DE Germany    German      Germany    German     ' TO lt_lines.
  APPEND '001 E FM Micronesia Micronesian Micronesia Micronesian' TO lt_lines.
  LOOP AT lt_lines INTO lv_line.
    CONDENSE lv_line.
    SPLIT lv_line AT space INTO TABLE lt_fields.
    lv_comp = 1.
    LOOP AT lt_fields INTO lv_field.
      ASSIGN COMPONENT lv_comp OF STRUCTURE ls_test TO <lv_field>.
      <lv_field> = lv_field.
      lv_comp = lv_comp + 1.
    ENDLOOP.
    APPEND ls_test TO ct_test.
  ENDLOOP.
ENDFORM.

FORM set_column_headers
    USING io_worksheet TYPE REF TO zcl_excel_worksheet
          iv_headers   TYPE csequence
    RAISING zcx_excel.

  DATA: lt_headers TYPE TABLE OF string,
        lv_header  TYPE string,
        lv_tabix   TYPE i.

  SPLIT iv_headers AT ';' INTO TABLE lt_headers.
  LOOP AT lt_headers INTO lv_header.
    lv_tabix = sy-tabix.
    io_worksheet->set_cell( ip_row = 1 ip_column = lv_tabix ip_value = lv_header ).
  ENDLOOP.

ENDFORM.
