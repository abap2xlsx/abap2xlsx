*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL18
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel18.

CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_demo_excel_generator.

  PUBLIC SECTION.
    METHODS zif_demo_excel_generator~get_information REDEFINITION.
    METHODS zif_demo_excel_generator~generate_excel REDEFINITION.

ENDCLASS.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_excel_generator TYPE REF TO lcl_excel_generator.

CONSTANTS: gc_save_file_name TYPE string VALUE '18_BookProtection.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel_generator.
  lo_excel = lo_excel_generator->zif_demo_excel_generator~generate_excel( ).

*** Create output
  lcl_output=>output( lo_excel ).



CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD zif_demo_excel_generator~get_information.

    result-objid = sy-repid.
    result-text = 'abap2xlsx Demo: Book protection'.
    result-filename = gc_save_file_name.

  ENDMETHOD.

  METHOD zif_demo_excel_generator~generate_excel.

    DATA: lo_excel                 TYPE REF TO zcl_excel,
          lo_worksheet             TYPE REF TO zcl_excel_worksheet,
          lv_style_protection_guid TYPE zexcel_cell_style.

    CREATE OBJECT lo_excel.

    " Get active sheet
    lo_excel->zif_excel_book_protection~protected     = zif_excel_book_protection=>c_protected.
    lo_excel->zif_excel_book_protection~lockrevision  = zif_excel_book_protection=>c_locked.
    lo_excel->zif_excel_book_protection~lockstructure = zif_excel_book_protection=>c_locked.
    lo_excel->zif_excel_book_protection~lockwindows   = zif_excel_book_protection=>c_locked.

    lo_worksheet = lo_excel->get_active_worksheet( ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'This cell is unlocked' ip_style = lv_style_protection_guid ).

    result = lo_excel.

  ENDMETHOD.

ENDCLASS.
