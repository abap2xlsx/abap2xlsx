*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL44
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel44.

DATA: lo_excel              TYPE REF TO zcl_excel,
      lo_worksheet          TYPE REF TO zcl_excel_worksheet.

DATA: lt_field_catalog    TYPE zexcel_t_fieldcatalog.

CONSTANTS: gc_save_file_name TYPE string VALUE '44_iTabEmptyOrNot.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

SELECTION-SCREEN BEGIN OF BLOCK b44 WITH FRAME TITLE txt_b44.

* No line if internal table is empty
PARAMETERS: p_mtyfil TYPE flag AS CHECKBOX DEFAULT 'X'.

SELECTION-SCREEN END OF BLOCK b44.

INITIALIZATION.
  txt_b44 = 'Testing empty file option'(b44).

START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Internal table' ).

  DATA lt_test TYPE TABLE OF sflight.

  lo_worksheet->bind_table( ip_table            = lt_test
                            iv_no_line_if_empty = p_mtyfil ).

*** Create output
  lcl_output=>output( EXPORTING cl_excel            = lo_excel
                                iv_writerclass_name = 'ZCL_EXCEL_WRITER_CSV' ).
