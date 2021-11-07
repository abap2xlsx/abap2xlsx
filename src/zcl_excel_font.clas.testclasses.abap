CLASS ltcl_Test DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.

  PRIVATE SECTION.
    METHODS calculate FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltcl_Test IMPLEMENTATION.

  METHOD calculate.

    DATA lv_width TYPE f.

    lv_width = zcl_excel_font=>calculate(
        ld_font_name   = 'foobar'
        ld_font_height = 20
        ld_flag_bold   = abap_false
        ld_flag_italic = abap_false
        ld_cell_value  = 'hello world' ).

    cl_abap_unit_assert=>assert_equals(
        act = lv_width
        exp = '2.75' ).

  ENDMETHOD.

ENDCLASS.
