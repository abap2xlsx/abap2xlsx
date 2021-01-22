*&---------------------------------------------------------------------*
*& Report zdemo_excel47
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel47.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_hyperlink TYPE REF TO zcl_excel_hyperlink,
      lo_column    TYPE REF TO zcl_excel_column.

CONSTANTS: gc_save_file_name TYPE string VALUE '01_HelloWorld.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

CLASS lcl_app DEFINITION.
  PUBLIC SECTION.
    METHODS main
      RAISING
        zcx_excel.
ENDCLASS.
CLASS lcl_app IMPLEMENTATION.
  METHOD main.
    TYPES: BEGIN OF helper_type,
             carrid  TYPE sflight-carrid,
             connid  TYPE sflight-connid,
             fldate  TYPE sflight-fldate,
             price   TYPE sflight-price,
             formula TYPE sflight-price,
           END OF helper_type.
    DATA: itab          TYPE STANDARD TABLE OF helper_type,
          field_catalog TYPE zexcel_t_fieldcatalog,
          ls_catalog    TYPE zexcel_s_fieldcatalog.

    " Creates active sheet
    CREATE OBJECT lo_excel.

    " Get active sheet
    lo_worksheet = lo_excel->get_active_worksheet( ).
    SELECT carrid, connid, fldate, price FROM sflight INTO TABLE @itab.

    field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = itab ).

    ls_catalog-sformula     = 'D2+100'.
    ls_catalog-sformula_ref = |E2:E{ lines( itab ) + 1 }|.
    MODIFY field_catalog FROM ls_catalog TRANSPORTING sformula sformula_ref sformula_from WHERE fieldname = 'FORMULA'.

    lo_worksheet->bind_table(
      EXPORTING
        ip_table            = itab
        it_field_catalog    = field_catalog
        is_table_settings   = VALUE zexcel_s_table_settings(
                                table_style         = zcl_excel_table=>builtinstyle_medium2
                                table_name          = 'TblFlights'
                                top_left_column     = 'A'
                                top_left_row        = 1
                                show_row_stripes    = abap_true ) ).

*** Create output
    lcl_output=>output( lo_excel ).

  ENDMETHOD.
ENDCLASS.

START-OF-SELECTION.
  TRY.
      NEW lcl_app( )->main( ).
    CATCH zcx_excel INTO DATA(lx).
      MESSAGE lx TYPE 'I' DISPLAY LIKE 'E'.
  ENDTRY.
