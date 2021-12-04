*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL23
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel24.

TYPE-POOLS: abap.

CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_excel_demo_generator.

  PUBLIC SECTION.
    METHODS zif_excel_demo_generator~get_information REDEFINITION.
    METHODS zif_excel_demo_generator~generate_excel REDEFINITION.

ENDCLASS.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_excel_generator TYPE REF TO lcl_excel_generator.

CONSTANTS: gc_save_file_name TYPE string VALUE '24_Sheets_with_different_default_date_formats.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel_generator.
  lo_excel = lo_excel_generator->zif_excel_demo_generator~generate_excel( ).

*** Create output
  lcl_output=>output( lo_excel ).



CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD zif_excel_demo_generator~get_information.

    result-objid = sy-repid.
    result-text = 'abap2xlsx Demo: Multiple sheets with different default date formats'.
    result-filename = gc_save_file_name.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~generate_excel.

    DATA: lo_excel     TYPE REF TO zcl_excel,
          lo_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_column    TYPE REF TO zcl_excel_column,
          lo_hyperlink TYPE REF TO zcl_excel_hyperlink.

    DATA: lv_file      TYPE xstring,
          lv_bytecount TYPE i,
          lt_file_tab  TYPE solix_tab.

    DATA: lv_full_path      TYPE string,
          lv_workdir        TYPE string,
          lv_file_separator TYPE c.

    DATA: lv_value TYPE string,
          lv_datum TYPE d.

    " Creates active sheet
    CREATE OBJECT lo_excel.

    " Get active sheet
    lo_worksheet = lo_excel->get_active_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Sheet1' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Default Date Format' ).
    " Insert current date
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Current Date:' ).
    lv_datum = zcl_excel_demo_generator=>get_date_now( ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = lv_datum ).

    lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet2!A1' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = 'This is a link to the second sheet' ip_hyperlink = lo_hyperlink ).
    lo_column = lo_worksheet->get_column( ip_column = 'A' ).
    lo_column->set_auto_size( ip_auto_size = abap_true ).


    " Second sheet
    lo_worksheet = lo_excel->add_new_worksheet( ).
    lo_worksheet->set_default_excel_date_format( zcl_excel_style_number_format=>c_format_date_yyyymmdd ).
    lo_worksheet->set_title( ip_title = 'Sheet2' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Date Format set to YYYYMMDD' ).
    " Insert current date
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Current Date:' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = sy-datum ).

    lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet3!B2' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = 'This is link to the third sheet' ip_hyperlink = lo_hyperlink ).

    " Third sheet
    lo_worksheet = lo_excel->add_new_worksheet( ).
    " TODO: It seems that the zcl_excel_style_number_format=>c_format_date_yyyymmddslash
    " does not produce a valid output
    lo_worksheet->set_default_excel_date_format( zcl_excel_style_number_format=>c_format_date_yyyymmddslash ).
    lo_worksheet->set_title( ip_title = 'Sheet3' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Date Format set to YYYY/MM/DD' ).
    " Insert current date
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Current Date:' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = sy-datum ).

    lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet4!B2' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = 'This is link to the 4th sheet' ip_hyperlink = lo_hyperlink ).

    " 4th sheet
    lo_worksheet = lo_excel->add_new_worksheet( ).
    " Illustrate the Problem caused by:
    " Excel 2000 incorrectly assumes that the year 1900 is a leap year.
    " http://support.microsoft.com/kb/214326/en-us
    lo_worksheet->set_title( ip_title = 'Sheet4' ).
    " Loop from Start Date to the Max Date current data in daily steps
    CONSTANTS: lv_max TYPE d VALUE '19000302'.

    DATA: lv_date TYPE d VALUE '19000226',
          lv_row  TYPE i.
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'Formated date' ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Integer value for this date' ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 3 ip_value = 'Date as string' ).

    lv_row = 4.
    WHILE lv_date < lv_max.
      lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = lv_date ).
      lv_value = zcl_excel_common=>date_to_excel_string( lv_date ).
      lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = lv_value ).
      lv_value = lv_date.
      lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = lv_value ).
      lv_date = lv_date + 1.
      lv_row = lv_row + 1.
    ENDWHILE.

    lv_row = lv_row + 1.

    lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet1!B2' ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = 'This is link to the first sheet' ip_hyperlink = lo_hyperlink ).

    lo_excel->set_active_sheet_index_by_name( 'Sheet1' ).

    result = lo_excel.

  ENDMETHOD.

ENDCLASS.
