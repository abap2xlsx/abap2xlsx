*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL19
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel19.

TYPE-POOLS: abap.

CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_excel_demo_generator.

  PUBLIC SECTION.
    METHODS zif_excel_demo_generator~get_information REDEFINITION.
    METHODS zif_excel_demo_generator~generate_excel REDEFINITION.

ENDCLASS.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_excel_generator TYPE REF TO lcl_excel_generator.

CONSTANTS: gc_save_file_name TYPE string VALUE '19_SetActiveSheet.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_noout TYPE abap_bool DEFAULT abap_true.


START-OF-SELECTION.

  CREATE OBJECT lo_excel_generator.
  lo_excel = lo_excel_generator->zif_excel_demo_generator~generate_excel( ).

*** Create output
  lcl_output=>output( lo_excel ).



CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD zif_excel_demo_generator~get_information.

    result-objid = sy-repid.
    result-text = 'abap2xlsx Demo: Set active sheet'.
    result-filename = gc_save_file_name.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~generate_excel.

    DATA: lo_excel     TYPE REF TO zcl_excel,
          lo_worksheet TYPE REF TO zcl_excel_worksheet.

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

    result = lo_excel.

  ENDMETHOD.

ENDCLASS.
