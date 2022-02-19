*&---------------------------------------------------------------------*
*& Report zdemo_excel49
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel49.
DATA: lo_excel          TYPE REF TO zcl_excel,
      lo_worksheet      TYPE REF TO zcl_excel_worksheet,
      ls_table_settings TYPE zexcel_s_table_settings,
      ls_t002t          TYPE t002t,
      lt_t002t          TYPE TABLE OF t002t.
CONSTANTS: gc_save_file_name TYPE string VALUE '49_Bind_Table_Conversion_Exit.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.
  ls_t002t-spras = 'D'.
  ls_t002t-sprsl = 'D'.
  ls_t002t-sptxt = 'Deutsch'.
  APPEND ls_t002t TO lt_t002t.
  ls_t002t-spras = 'D'.
  ls_t002t-sprsl = 'E'.
  ls_t002t-sptxt = 'Englisch'.
  APPEND ls_t002t TO lt_t002t.
  ls_t002t-spras = 'E'.
  ls_t002t-sprsl = 'D'.
  ls_t002t-sptxt = 'German'.
  APPEND ls_t002t TO lt_t002t.
  ls_t002t-spras = 'E'.
  ls_t002t-sprsl = 'E'.
  ls_t002t-sptxt = 'English'.
  APPEND ls_t002t TO lt_t002t.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->get_active_worksheet( ).
  ls_table_settings-top_left_column = 'A'.
  ls_table_settings-top_left_row = 1.
  lo_worksheet->bind_table(
      ip_table            = lt_t002t
      is_table_settings   = ls_table_settings
      ip_conv_exit_length = abap_true ).
  lcl_output=>output( lo_excel ).
