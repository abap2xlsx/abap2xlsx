*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL19
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel19.

TYPE-POOLS: abap.

DATA: lo_excel                  TYPE REF TO zcl_excel,
      lo_worksheet              TYPE REF TO zcl_excel_worksheet.


CONSTANTS: gc_save_file_name TYPE string VALUE '19_SetActiveSheet.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_noout TYPE xfeld DEFAULT abap_true.


START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  " First Worksheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'First' ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'This is Sheet 1' ).

  " Second Worksheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Second' ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'This is Sheet 2' ).

  " Third Worksheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Third' ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'This is Sheet 3' ).

  IF p_noout EQ abap_false.
    " lo_excel->set_active_sheet_index_by_name( data_sheet_name ).
    DATA: active_sheet_index TYPE zexcel_active_worksheet.
    active_sheet_index = lo_excel->get_active_sheet_index( ).
    WRITE: 'Sheet Index before: ', active_sheet_index.
  ENDIF.
  lo_excel->set_active_sheet_index( '2' ).
  IF p_noout EQ abap_false.
    active_sheet_index = lo_excel->get_active_sheet_index( ).
    WRITE: 'Sheet Index after: ', active_sheet_index.
  ENDIF.


*** Create output
  lcl_output=>output( lo_excel ).
