CLASS ltcl_test DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.

  PRIVATE SECTION.
    METHODS calculate FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltcl_test IMPLEMENTATION.

  METHOD calculate.

    DATA lv_width TYPE f.

    lv_width = zcl_excel_font=>calculate_text_width(
        iv_font_name   = 'foobar'
        iv_font_height = 20
        iv_flag_bold   = abap_false
        iv_flag_italic = abap_false
        iv_cell_value  = 'hello world' ).

    cl_abap_unit_assert=>assert_equals(
        act = lv_width
        exp = '2.75' ).

  ENDMETHOD.

ENDCLASS.
