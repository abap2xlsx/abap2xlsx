*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL44
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel44.

DATA: lo_excel_no_line_if_empty     TYPE REF TO zcl_excel,
      lo_excel                      TYPE REF TO zcl_excel,
      lo_worksheet_no_line_if_empty TYPE REF TO zcl_excel_worksheet,
      lo_worksheet                  TYPE REF TO zcl_excel_worksheet.

DATA: lt_field_catalog    TYPE zexcel_t_fieldcatalog.

DATA: gc_save_file_name TYPE string VALUE '44_iTabEmpty.csv'.
INCLUDE zdemo_excel_outputopt_incl.

SELECTION-SCREEN BEGIN OF BLOCK b44 WITH FRAME TITLE txt_b44.

* No line if internal table is empty
DATA: p_mtyfil TYPE flag VALUE abap_true.

SELECTION-SCREEN END OF BLOCK b44.

INITIALIZATION.
  txt_b44 = 'Testing empty file option'(b44).

START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel_no_line_if_empty.
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet_no_line_if_empty = lo_excel_no_line_if_empty->get_active_worksheet( ).
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet_no_line_if_empty->set_title( 'Internal table' ).
  lo_worksheet->set_title( 'Internal table' ).

  DATA lt_test TYPE TABLE OF sflight.

  lo_worksheet_no_line_if_empty->bind_table( ip_table            = lt_test
                            iv_no_line_if_empty = p_mtyfil ).

  p_mtyfil = abap_false.
  lo_worksheet->bind_table( ip_table            = lt_test
                            iv_no_line_if_empty = p_mtyfil ).

*** Create output
  lcl_output=>output( EXPORTING cl_excel            = lo_excel_no_line_if_empty
                                iv_writerclass_name = 'ZCL_EXCEL_WRITER_CSV' ).

  gc_save_file_name = '44_iTabNotEmpty.csv'.
  lcl_output=>output( EXPORTING cl_excel            = lo_excel
                              iv_writerclass_name = 'ZCL_EXCEL_WRITER_CSV' ).
