*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL17
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel17.

CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_excel_demo_generator.

  PUBLIC SECTION.
    METHODS zif_excel_demo_generator~get_information REDEFINITION.
    METHODS zif_excel_demo_generator~generate_excel REDEFINITION.
    METHODS zif_excel_demo_generator~checker_initialization REDEFINITION.

ENDCLASS.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_excel_generator TYPE REF TO lcl_excel_generator.

CONSTANTS: gc_save_file_name TYPE string VALUE '17_SheetProtection.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_pwd   TYPE zexcel_aes_password LOWER CASE DEFAULT 'secret'.

START-OF-SELECTION.

  CREATE OBJECT lo_excel_generator.
  lo_excel = lo_excel_generator->zif_excel_demo_generator~generate_excel( ).

*** Create output
  lcl_output=>output( lo_excel ).



CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD zif_excel_demo_generator~get_information.

    result-objid = sy-repid.
    result-text = 'abap2xlsx Demo: Sheet Protection'.
    result-filename = gc_save_file_name.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~generate_excel.

    DATA: lo_excel                 TYPE REF TO zcl_excel,
          lo_worksheet             TYPE REF TO zcl_excel_worksheet,
          lo_style_protection      TYPE REF TO zcl_excel_style,
          lv_style_protection_guid TYPE zexcel_cell_style,
          lo_style                 TYPE REF TO zcl_excel_style,
          lv_style                 TYPE zexcel_cell_style.

    CREATE OBJECT lo_excel.

    " Get active sheet
    lo_worksheet = lo_excel->get_active_worksheet( ).
    lo_worksheet->zif_excel_sheet_protection~protected  = zif_excel_sheet_protection=>c_protected.
*  lo_worksheet->zif_excel_sheet_protection~password   = 'DAA7'. "it is the encoded word "secret"
    lo_worksheet->zif_excel_sheet_protection~password   = zcl_excel_common=>encrypt_password( p_pwd ).
    lo_worksheet->zif_excel_sheet_protection~sheet      = zif_excel_sheet_protection=>c_active.
    lo_worksheet->zif_excel_sheet_protection~objects    = zif_excel_sheet_protection=>c_active.
    lo_worksheet->zif_excel_sheet_protection~scenarios  = zif_excel_sheet_protection=>c_active.
    " First style to unlock a cell
    lo_style_protection = lo_excel->add_new_style( ).
    lo_style_protection->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.
    lv_style_protection_guid = lo_style_protection->get_guid( ).
    " Another style which should not affect the unlock style
    lo_style = lo_excel->add_new_style( ).
    lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb  = 'FFCC3333'.
    lv_style = lo_style->get_guid( ).
    lo_worksheet->set_cell( ip_row = 3 ip_column = 'C' ip_value = 'This cell is locked locked and has the second formating' ip_style = lv_style ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'This cell is unlocked' ip_style = lv_style_protection_guid ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 'This cell is locked as all the others empty cell' ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 'This cell is unlocked' ip_style = lv_style_protection_guid ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 'This cell is unlocked' ip_style = lv_style_protection_guid ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 'This cell is locked as all the others empty cell' ).

    result = lo_excel.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~checker_initialization.

    p_pwd = 'secret'.

  ENDMETHOD.

ENDCLASS.
