*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL8
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel8.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_range                TYPE REF TO zcl_excel_range.

DATA: lv_title          TYPE zexcel_sheet_title.

CONSTANTS: gc_save_file_name TYPE string VALUE '08_Range.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lv_title = 'Sheet1'.
  lo_worksheet->set_title( lv_title ).
  lo_range = lo_excel->add_new_range( ).
  lo_range->name = 'range'.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'C'
                       ip_start_row     = 4
                       ip_stop_column   = 'C'
                       ip_stop_row      = 8 ).


  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'Apple' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 'Banana' ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 'Blueberry' ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 'Ananas' ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 'Grapes' ).

  " Define another Range with a name longer than 40 characters that
  " tests the fix of issue #168 ranges namescan be only up to 20 chars

  lo_range = lo_excel->add_new_range( ).
  lo_range->name = 'A_range_with_a_name_that_is_longer_than_20_characters'.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'D'
                       ip_start_row     = 4
                       ip_stop_column   = 'D'
                       ip_stop_row      = 5 ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = 'Range Value 1' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'D' ip_value = 'Range Value 2' ).

  " issue #163
  " Define another Range with sheet visibility
  lo_range = lo_worksheet->add_new_range( ).
  lo_range->name = 'A_range_with_sheet_visibility'.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'E'
                       ip_start_row     = 4
                       ip_stop_column   = 'E'
                       ip_stop_row      = 5 ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'E' ip_value = 'Range Value 3' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'E' ip_value = 'Range Value 4' ).
  " issue #163

*** Create output
  lcl_output=>output( lo_excel ).
