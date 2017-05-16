*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL18
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel18.

DATA: lo_excel                  TYPE REF TO zcl_excel,
      lo_worksheet              TYPE REF TO zcl_excel_worksheet,
      lv_style_protection_guid  TYPE zexcel_cell_style.


CONSTANTS: gc_save_file_name TYPE string VALUE '18_BookProtection.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_excel->zif_excel_book_protection~protected     = zif_excel_book_protection=>c_protected.
  lo_excel->zif_excel_book_protection~lockrevision  = zif_excel_book_protection=>c_locked.
  lo_excel->zif_excel_book_protection~lockstructure = zif_excel_book_protection=>c_locked.
  lo_excel->zif_excel_book_protection~lockwindows   = zif_excel_book_protection=>c_locked.

  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'This cell is unlocked' ip_style = lv_style_protection_guid ).




*** Create output
  lcl_output=>output( lo_excel ).
