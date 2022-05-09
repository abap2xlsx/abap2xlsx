*"* use this source file for your ABAP unit test classes
CLASS lcl_test DEFINITION DEFERRED.
CLASS zcl_excel_reader_huge_file DEFINITION LOCAL FRIENDS lcl_test.

*
CLASS lcl_test DEFINITION FOR TESTING
      RISK LEVEL HARMLESS
      DURATION SHORT.

  PRIVATE SECTION.
    DATA:
      out       TYPE REF TO zcl_excel_reader_huge_file, " object under test
      excel     TYPE REF TO zcl_excel,
      worksheet TYPE REF TO zcl_excel_worksheet.
    METHODS:
      setup RAISING cx_static_check,
      test_number FOR TESTING RAISING cx_static_check,
      test_shared_string FOR TESTING RAISING cx_static_check,
      test_shared_string_missing FOR TESTING RAISING cx_static_check,
      test_inline_string FOR TESTING RAISING cx_static_check,
      test_empty_cells FOR TESTING RAISING cx_static_check,
      test_boolean FOR TESTING RAISING cx_static_check,
      test_style FOR TESTING RAISING cx_static_check,
      test_style_missing FOR TESTING RAISING cx_static_check,
      test_formula FOR TESTING RAISING cx_static_check,
      test_read_shared_strings FOR TESTING RAISING cx_static_check,
      test_shared_string_some_empty FOR TESTING RAISING cx_static_check,
      test_shared_string_multi_style FOR TESTING RAISING cx_static_check,
      test_skip_to_inexistent FOR TESTING RAISING cx_static_check,
      get_reader IMPORTING iv_xml TYPE string RETURNING VALUE(eo_reader) TYPE REF TO if_sxml_reader,
      assert_value_equals IMPORTING iv_row TYPE i DEFAULT 1 iv_col TYPE i DEFAULT 1 iv_value TYPE string,
      assert_formula_equals IMPORTING iv_row TYPE i DEFAULT 1 iv_col TYPE i DEFAULT 1 iv_formula TYPE string,
      assert_style_equals IMPORTING iv_row TYPE i DEFAULT 1 iv_col TYPE i DEFAULT 1 iv_style TYPE zexcel_cell_style,
      assert_datatype_equals IMPORTING iv_row TYPE i DEFAULT 1 iv_col TYPE i DEFAULT 1 iv_datatype TYPE string.

ENDCLASS.                    "lcl_test DEFINITION

*
CLASS lcl_test IMPLEMENTATION.

*
  METHOD test_number.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string.
    lo_reader = get_reader(
      `<c r="A1" t="n"><v>17</v></c>`
      ).
    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        assert_value_equals( `17` ).
        assert_datatype_equals( `n` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
        cl_abap_unit_assert=>fail( lv_text ).
    ENDTRY.
  ENDMETHOD.                    "test_shared_string

*
  METHOD test_shared_string.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string.
    DATA: ls_shared_string TYPE zcl_excel_reader_huge_file=>t_shared_string.
    ls_shared_string-value = `Test1`.
    APPEND ls_shared_string TO out->shared_strings.
    ls_shared_string-value = `Test2`.
    APPEND ls_shared_string TO out->shared_strings.
    lo_reader = get_reader(
      `<c r="A1" t="s"><v>1</v></c>`
      ).
    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        assert_value_equals( `Test2` ).
        assert_datatype_equals( `s` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
        cl_abap_unit_assert=>fail( lv_text ).
    ENDTRY.
  ENDMETHOD.                    "test_shared_string
*
  METHOD test_shared_string_missing.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string.
    DATA: ls_shared_string TYPE zcl_excel_reader_huge_file=>t_shared_string.
    ls_shared_string-value = `Test`.
    APPEND ls_shared_string TO out->shared_strings.
    lo_reader = get_reader(
      `<c r="A1" t="s"><v>1</v></c>`
      ).

    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        cl_abap_unit_assert=>fail( `Index to non-existent shared string should give an error` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
    ENDTRY.

  ENDMETHOD.
*
  METHOD test_inline_string.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string.
    lo_reader = get_reader(
      `<c r="A1" t="inlineStr"><is><t>Alpha</t></is></c>`
      ).
    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        assert_value_equals( `Alpha` ).
        assert_datatype_equals( `inlineStr` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
        cl_abap_unit_assert=>fail( lv_text ).
    ENDTRY.
  ENDMETHOD.                    "test_inline_string

*
  METHOD test_boolean.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string.
    lo_reader = get_reader(
      `<c r="A1" t="b"><v>1</v></c>`
      ).
    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        assert_value_equals( `1` ).
        assert_datatype_equals( `b` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
        cl_abap_unit_assert=>fail( lv_text ).
    ENDTRY.
  ENDMETHOD.                    "test_boolean

*
  METHOD test_formula.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string.
    lo_reader = get_reader(
      `<c r="A1" t="n"><f>A2*A2</f></c>`
      ).
    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        assert_formula_equals( `A2*A2` ).
        assert_datatype_equals( `n` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
        cl_abap_unit_assert=>fail( lv_text ).
    ENDTRY.
  ENDMETHOD.                    "test_formula

*
  METHOD test_empty_cells.

* There is no need to store an empty cell in the ABAP worksheet structure
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string.
    DATA: ls_shared_string TYPE zcl_excel_reader_huge_file=>t_shared_string.
    ls_shared_string-value = ``.
    APPEND ls_shared_string TO out->shared_strings.
    ls_shared_string-value = `t`.
    APPEND ls_shared_string TO out->shared_strings.
    lo_reader = get_reader(
      `<c r="A1" t="s"><v>0</v></c>` &
      `<c r="A2" t="inlineStr"><is><t></t></is></c>` &
      `<c r="A3" t="s"><v>1</v></c>`
      ).

    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).

        assert_value_equals( iv_row = 1 iv_col = 1 iv_value = `` ).
        assert_value_equals( iv_row = 2 iv_col = 1 iv_value = `` ).
        assert_value_equals( iv_row = 3 iv_col = 1 iv_value = `t` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
        cl_abap_unit_assert=>fail( lv_text ).
    ENDTRY.
  ENDMETHOD.

*
  METHOD test_style.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lo_ex     TYPE REF TO lcx_not_found,
          lv_text   TYPE string,
          lo_style  TYPE REF TO zcl_excel_style,
          lv_guid   TYPE zexcel_cell_style.
    CREATE OBJECT lo_style.
    APPEND lo_style TO out->styles.
    lv_guid = lo_style->get_guid( ).

    lo_reader = get_reader(
      `<c r="A1" s="0"><v>18</v></c>`
      ).
    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        assert_style_equals( lv_guid ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
        cl_abap_unit_assert=>fail( lv_text ).
    ENDTRY.

  ENDMETHOD.                    "test_style

*
  METHOD test_style_missing.

    DATA:
      lo_reader TYPE REF TO if_sxml_reader,
      lo_ex     TYPE REF TO lcx_not_found,
      lv_text   TYPE string.

    lo_reader = get_reader(
      `<c r="A1" s="0"><v>18</v></c>`
      ).

    TRY.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        cl_abap_unit_assert=>fail( `Reference to non-existent style should throw an lcx_not_found exception` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ).  " >>> May inspect the message in the debugger
    ENDTRY.

  ENDMETHOD.                    "test_style

*
  METHOD test_read_shared_strings.
    DATA: lo_c2x     TYPE REF TO cl_abap_conv_out_ce,
          lv_xstring TYPE xstring,
          lo_reader  TYPE REF TO if_sxml_reader,
          lt_act     TYPE stringtab,
          lt_exp     TYPE stringtab.

    lo_c2x = cl_abap_conv_out_ce=>create( ).
    lo_c2x->convert( EXPORTING data   = `<sst><si><t/></si><si><t>Alpha</t></si><si><t>Bravo</t></si></sst>`
                     IMPORTING buffer = lv_xstring ).
    lo_reader = cl_sxml_string_reader=>create( lv_xstring ).
    APPEND :
     `` TO lt_exp,
     `Alpha` TO lt_exp,
     `Bravo` TO lt_exp.

    lt_act = out->read_shared_strings( lo_reader ).

    cl_abap_unit_assert=>assert_equals( act = lt_act
                   exp = lt_exp ).

  ENDMETHOD.

*
  METHOD test_shared_string_some_empty.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lt_act    TYPE stringtab,
          lt_exp    TYPE stringtab.
    lo_reader = cl_sxml_string_reader=>create( cl_abap_codepage=>convert_to(
      `<sst><si><t/></si>` &
      `<si><t>Alpha</t></si>` &
      `<si><t/></si>` &
      `<si><t>Bravo</t></si></sst>`
      ) ).
    APPEND :
     `` TO lt_exp,
     `Alpha` TO lt_exp,
     ``      TO lt_exp,
     `Bravo` TO lt_exp.

    lt_act = out->read_shared_strings( lo_reader ).

    cl_abap_unit_assert=>assert_equals( act = lt_act
                   exp = lt_exp ).

  ENDMETHOD.

*
  METHOD test_shared_string_multi_style.
    DATA: lo_reader TYPE REF TO if_sxml_reader,
          lt_act    TYPE stringtab,
          lt_exp    TYPE stringtab.
    lo_reader = cl_sxml_string_reader=>create( cl_abap_codepage=>convert_to(
      `<sst>` &&
      `<si><t>Cell A2</t></si>` &&
      `<si><r><t>the following coloured part</t></r><r><rPr><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve"> will be preserved</t></r></si>` &&
      `<si><t>Cell A3</t></si>` &&
      `</sst>` ) ).
    APPEND :
     `Cell A2` TO lt_exp,
     `the following coloured part will be preserved` TO lt_exp,
     `Cell A3` TO lt_exp.

    lt_act = out->read_shared_strings( lo_reader ).

    cl_abap_unit_assert=>assert_equals( act = lt_act
                   exp = lt_exp ).

  ENDMETHOD.


*
  METHOD test_skip_to_inexistent.
    DATA: lo_c2x     TYPE REF TO cl_abap_conv_out_ce,
          lv_xstring TYPE xstring,
          lo_reader  TYPE REF TO if_sxml_reader,
          lo_ex      TYPE REF TO lcx_not_found,
          lv_text    TYPE string.

    lo_c2x = cl_abap_conv_out_ce=>create( ).
    lo_c2x->convert( EXPORTING data   = `<sst><si><t/></si><si><t>Alpha</t></si><si><t>Bravo</t></si></sst>`
                     IMPORTING buffer = lv_xstring ).
    lo_reader = cl_sxml_string_reader=>create( lv_xstring ).
    TRY.
        out->skip_to( iv_element_name = `nonExistingElement` io_reader = lo_reader ).
        cl_abap_unit_assert=>fail( `Skipping to non-existing element must raise lcx_not_found exception` ).
      CATCH lcx_not_found INTO lo_ex.
        lv_text = lo_ex->get_text( ). " May inspect exception text in debugger
    ENDTRY.
  ENDMETHOD.

*
  METHOD get_reader.
    DATA: lv_full    TYPE string,
          lo_c2x     TYPE REF TO cl_abap_conv_out_ce,
          lv_xstring TYPE xstring.
    CONCATENATE `<root><sheetData><row>` iv_xml `</row></sheetData></root>` INTO lv_full.
    lo_c2x = cl_abap_conv_out_ce=>create( ).
    lo_c2x->convert( EXPORTING data   = lv_full
                     IMPORTING buffer = lv_xstring ).
    eo_reader = cl_sxml_string_reader=>create( lv_xstring ).
  ENDMETHOD.                    "get_reader
*
  METHOD assert_value_equals.

    CONSTANTS: lc_empty_string TYPE string VALUE IS INITIAL.

    FIELD-SYMBOLS: <ls_cell_data> TYPE zexcel_s_cell_data,
                   <lv_value>     TYPE string.

    READ TABLE worksheet->sheet_content ASSIGNING <ls_cell_data>
      WITH TABLE KEY cell_row = iv_row cell_column = iv_col.
    IF sy-subrc EQ 0.
      ASSIGN <ls_cell_data>-cell_value TO <lv_value>.
    ELSE.
      ASSIGN lc_empty_string TO <lv_value>.
    ENDIF.

    cl_abap_unit_assert=>assert_equals( act = <lv_value>
                   exp = iv_value ).

  ENDMETHOD.                    "assert_value_equals
**
  METHOD assert_formula_equals.

    FIELD-SYMBOLS: <ls_cell_data> TYPE zexcel_s_cell_data.

    READ TABLE worksheet->sheet_content ASSIGNING <ls_cell_data>
      WITH TABLE KEY cell_row = iv_row cell_column = iv_col.
    cl_abap_unit_assert=>assert_subrc( sy-subrc ).

    cl_abap_unit_assert=>assert_equals( act = <ls_cell_data>-cell_formula
                   exp = iv_formula ).

  ENDMETHOD.                    "assert_formula_equals
*
  METHOD assert_style_equals.

    FIELD-SYMBOLS: <ls_cell_data> TYPE zexcel_s_cell_data.

    READ TABLE worksheet->sheet_content ASSIGNING <ls_cell_data>
      WITH TABLE KEY cell_row = iv_row cell_column = iv_col.
    cl_abap_unit_assert=>assert_subrc( sy-subrc ).

    cl_abap_unit_assert=>assert_equals( act = <ls_cell_data>-cell_style
                   exp = iv_style ).

  ENDMETHOD.
*
  METHOD assert_datatype_equals.

    FIELD-SYMBOLS: <ls_cell_data> TYPE zexcel_s_cell_data.

    READ TABLE worksheet->sheet_content ASSIGNING <ls_cell_data>
      WITH TABLE KEY cell_row = iv_row cell_column = iv_col.
    cl_abap_unit_assert=>assert_subrc( sy-subrc ).

    cl_abap_unit_assert=>assert_equals( act = <ls_cell_data>-data_type
                   exp = iv_datatype ).

  ENDMETHOD.                    "assert_datatype_equals
  METHOD setup.
    CREATE OBJECT out.
    CREATE OBJECT excel.
    CREATE OBJECT worksheet
      EXPORTING
        ip_excel = excel.
  ENDMETHOD.                    "setup
ENDCLASS.                    "lcl_test IMPLEMENTATION
