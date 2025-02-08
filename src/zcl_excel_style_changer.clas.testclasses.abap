CLASS ltcl_test DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.

  PRIVATE SECTION.
    DATA mi_cut       TYPE REF TO zif_excel_style_changer.
    DATA mo_excel     TYPE REF TO zcl_excel.
    DATA mo_worksheet TYPE REF TO zcl_excel_worksheet.

    METHODS setup RAISING cx_static_check.
    METHODS apply FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltcl_test IMPLEMENTATION.

  METHOD setup.
    CREATE OBJECT mo_excel.
    mo_worksheet = mo_excel->get_active_worksheet( ).
    mi_cut = zcl_excel_style_changer=>create( mo_excel ).
  ENDMETHOD.

  METHOD apply.

    DATA lv_guid TYPE zexcel_cell_style.

    mo_worksheet->set_cell(
      ip_column = 'B'
      ip_row    = 2
      ip_value  = 'Hello' ).

    mi_cut->set_font_bold( abap_true ).

    lv_guid = mi_cut->apply(
      ip_worksheet = mo_worksheet
      ip_column    = 'B'
      ip_row       = 2 ).

    mo_excel->get_style_to_guid( lv_guid ).

  ENDMETHOD.

ENDCLASS.
