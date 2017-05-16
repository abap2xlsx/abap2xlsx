*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel33.

TYPE-POOLS: abap.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_converter            TYPE REF TO zcl_excel_converter,
      lo_autofilter           TYPE REF TO zcl_excel_autofilter.

DATA lt_test TYPE TABLE OF t005t.

DATA: l_cell_value TYPE zexcel_cell_value,
      ls_area      TYPE zexcel_s_autofilter_area.

CONSTANTS: c_airlines TYPE string VALUE 'Airlines'.

CONSTANTS: gc_save_file_name TYPE string VALUE '33_autofilter.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Internal table').

  SELECT * UP TO 2 ROWS FROM t005t INTO TABLE lt_test.  "#EC CI_NOWHERE

  CREATE OBJECT lo_converter.

  lo_converter->convert( EXPORTING
                            it_table     = lt_test
                            i_row_int    = 1
                            i_column_int = 1
                            io_worksheet = lo_worksheet
                         CHANGING
                            co_excel     = lo_excel ) .

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
