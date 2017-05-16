*&---------------------------------------------------------------------*
*& Report  ZDEMO_TECHED3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_teched3.

*******************************
*   Data Object declaration   *
*******************************

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_reader         TYPE REF TO zif_excel_reader,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet.

DATA: lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

DATA: lt_files          TYPE filetable,
      ls_file           TYPE file_table,
      lv_rc             TYPE i,
      lv_value          TYPE zexcel_cell_value.

CONSTANTS: gc_save_file_name TYPE string VALUE 'TechEd01.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.

*******************************
*    abap2xlsx read XLSX    *
*******************************
  CREATE OBJECT lo_excel_reader TYPE zcl_excel_reader_2007.
  lo_excel = lo_excel_reader->load_file( lv_full_path ).

  lo_excel->set_active_sheet_index( 1 ).
  lo_worksheet = lo_excel->get_active_worksheet( ).

  lo_worksheet->get_cell( EXPORTING ip_column = 'C'
                                    ip_row    = 10
                          IMPORTING ep_value  = lv_value ).

  WRITE: 'abap2xlsx total score is ', lv_value.

*** Create output
  lcl_output=>output( lo_excel ).
