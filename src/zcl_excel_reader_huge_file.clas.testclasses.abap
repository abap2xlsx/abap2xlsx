*"* use this source file for your ABAP unit test classes
class lcl_test definition deferred.
class zcl_excel_reader_huge_file definition local friends lcl_test.

*
class lcl_test definition for testing  " #AU Risk_Level Harmless
      inheriting from cl_aunit_assert. " #AU Duration Short

  private section.
    data:
      out type ref to zcl_excel_reader_huge_file, " object under test
      excel type ref to zcl_excel,
      worksheet type ref to zcl_excel_worksheet.
    methods:
      setup,
      test_number for testing,
      test_shared_string for testing,
      test_shared_string_missing for testing,
      test_inline_string for testing,
      test_empty_cells for testing,
      test_boolean for testing,
      test_style for testing,
      test_style_missing for testing,
      test_formula for testing,
      test_read_shared_strings for testing,
      test_shared_string_some_empty for testing,
      test_skip_to_inexistent for testing,
      get_reader importing iv_xml type string returning value(eo_reader) type ref to if_sxml_reader,
      assert_value_equals importing iv_row type i default 1 iv_col type i default 1 iv_value type string,
      assert_formula_equals importing iv_row type i default 1 iv_col type i default 1 iv_formula type string,
      assert_style_equals importing iv_row type i default 1 iv_col type i default 1 iv_style type ZEXCEL_CELL_STYLE,
      assert_datatype_equals importing iv_row type i default 1 iv_col type i default 1 iv_datatype type string.

endclass.                    "lcl_test DEFINITION

*
class lcl_test implementation.

*
  method test_number.
    data lo_reader type ref to if_sxml_reader.
    lo_reader = get_reader(
      `<c r="A1" t="n"><v>17</v></c>`
      ).
    out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
    assert_value_equals( `17` ).
    assert_datatype_equals( `n` ).
  endmethod.                    "test_shared_string

*
  method test_shared_string.
    data lo_reader type ref to if_sxml_reader.
    append `Test1` to out->shared_strings.
    append `Test2` to out->shared_strings.
    lo_reader = get_reader(
      `<c r="A1" t="s"><v>1</v></c>`
      ).
    out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
    assert_value_equals( `Test2` ).
    assert_datatype_equals( `s` ).
  endmethod.                    "test_shared_string
*
  method test_shared_string_missing.

    data: lo_reader type ref to if_sxml_reader,
          lo_ex type ref to lcx_not_found,
          lv_text type string.
    append `Test` to out->shared_strings.
    lo_reader = get_reader(
      `<c r="A1" t="s"><v>1</v></c>`
      ).

    try.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        fail(`Index to non-existent shared string should give an error`).
      catch lcx_not_found into lo_ex.
        lv_text = lo_ex->get_text( ). " >>> May inspect the message in the debugger
    endtry.

  endmethod.
*
  method test_inline_string.
    data lo_reader type ref to if_sxml_reader.
    lo_reader = get_reader(
      `<c r="A1" t="inlineStr"><is><t>Alpha</t></is></c>`
      ).
    out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
    assert_value_equals( `Alpha` ).
    assert_datatype_equals( `inlineStr` ).
  endmethod.                    "test_inline_string

*
  method test_boolean.
    data lo_reader type ref to if_sxml_reader.
    lo_reader = get_reader(
      `<c r="A1" t="b"><v>1</v></c>`
      ).
    out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
    assert_value_equals( `1` ).
    assert_datatype_equals( `b` ).
  endmethod.                    "test_boolean

*
  method test_formula.
    data lo_reader type ref to if_sxml_reader.
    lo_reader = get_reader(
      `<c r="A1" t="n"><f>A2*A2</f></c>`
      ).
    out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
    assert_formula_equals( `A2*A2` ).
    assert_datatype_equals( `n` ).
  endmethod.                    "test_formula

*
  method test_empty_cells.

* There is no need to store an empty cell in the ABAP worksheet structure

    data: lo_reader type ref to if_sxml_reader.
    append `` to out->shared_strings.
    append `t` to out->shared_strings.
    lo_reader = get_reader(
      `<c r="A1" t="s"><v>0</v></c>` &
      `<c r="A2" t="inlineStr"><is><t></t></is></c>` &
      `<c r="A3" t="s"><v>1</v></c>`
      ).

    out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).

    assert_value_equals( iv_row = 1 iv_col = 1 iv_value = `` ).
    assert_value_equals( iv_row = 2 iv_col = 1 iv_value = `` ).
    assert_value_equals( iv_row = 3 iv_col = 1 iv_value = `t` ).

  endmethod.

*
  method test_style.
    data:
      lo_reader type ref to if_sxml_reader,
      lo_style  type ref to zcl_excel_style,
      lv_guid type ZEXCEL_CELL_STYLE.
      create object lo_style.
      append lo_style to out->styles.
      lv_guid = lo_style->get_guid( ).

    lo_reader = get_reader(
      `<c r="A1" s="0"><v>18</v></c>`
      ).
    out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).

    assert_style_equals( lv_guid ).

  endmethod.                    "test_style

*
  method test_style_missing.

    data:
      lo_reader type ref to if_sxml_reader,
      lo_ex type ref to lcx_not_found,
      lv_text type string.

    lo_reader = get_reader(
      `<c r="A1" s="0"><v>18</v></c>`
      ).

    try.
        out->read_worksheet_data( io_reader = lo_reader io_worksheet = worksheet ).
        fail(`Reference to non-existent style should throw an lcx_not_found exception`).
      catch lcx_not_found into lo_ex.
        lv_text = lo_ex->get_text( ).  " >>> May inspect the message in the debugger
    endtry.

  endmethod.                    "test_style

*
  method test_read_shared_strings.
    data: lo_c2x type ref to cl_abap_conv_out_ce,
          lv_xstring type xstring,
          lo_reader type ref to if_sxml_reader,
          lt_act type stringtab,
          lt_exp type stringtab.

    lo_c2x = cl_abap_conv_out_ce=>create( ).
    lo_c2x->convert( exporting data   = `<sst><si><t/></si><si><t>Alpha</t></si><si><t>Bravo</t></si></sst>`
                     importing buffer = lv_xstring ).
    lo_reader = cl_sxml_string_reader=>create( lv_xstring ).
    append :
     `` to lt_exp,
     `Alpha` to lt_exp,
     `Bravo` to lt_exp.

    lt_act = out->read_shared_strings( lo_reader ).

    assert_equals( act = lt_act
                   exp = lt_exp ).

  endmethod.

*
  method test_shared_string_some_empty.
    data: lo_reader type ref to if_sxml_reader,
          lt_act type stringtab,
          lt_exp type stringtab.
    lo_reader = cl_sxml_string_reader=>create( cl_abap_codepage=>convert_to(
      `<sst><si><t/></si>` &
      `<si><t>Alpha</t></si>` &
      `<si><t/></si>` &
      `<si><t>Bravo</t></si></sst>`
      ) ).
    append :
     `` to lt_exp,
     `Alpha` to lt_exp,
     ``      to lt_exp,
     `Bravo` to lt_exp.

    lt_act = out->read_shared_strings( lo_reader ).

    assert_equals( act = lt_act
                   exp = lt_exp ).

  endmethod.


*
  method test_skip_to_inexistent.
    data: lo_c2x type ref to cl_abap_conv_out_ce,
          lv_xstring type xstring,
          lo_reader type ref to if_sxml_reader,
          lo_ex type ref to lcx_not_found,
          lv_text type string.

    lo_c2x = cl_abap_conv_out_ce=>create( ).
    lo_c2x->convert( exporting data   = `<sst><si><t/></si><si><t>Alpha</t></si><si><t>Bravo</t></si></sst>`
                     importing buffer = lv_xstring ).
    lo_reader = cl_sxml_string_reader=>create( lv_xstring ).
    try.
        out->skip_to( iv_element_name = `nonExistingElement` io_reader = lo_reader ).
        fail(`Skipping to non-existing element must raise lcx_not_found exception`).
      catch lcx_not_found into lo_ex.
        lv_text = lo_ex->get_text( ). " May inspect exception text in debugger
    endtry.
  endmethod.

*
  method get_reader.
    data: lv_full type string,
          lo_c2x type ref to cl_abap_conv_out_ce,
          lv_xstring type xstring.
    concatenate `<root><sheetData><row>` iv_xml `</row></sheetData></root>` into lv_full.
    lo_c2x = cl_abap_conv_out_ce=>create( ).
    lo_c2x->convert( exporting data   = lv_full
                     importing buffer = lv_xstring ).
    eo_reader = cl_sxml_string_reader=>create( lv_xstring ).
  endmethod.                    "get_reader
*
  method assert_value_equals.

    constants: lc_empty_string type string value is initial.

    field-symbols: <ls_cell_data> type zexcel_s_cell_data,
                   <lv_value> type string.

    read table worksheet->sheet_content assigning <ls_cell_data>
      with table key cell_row = iv_row cell_column = iv_col.
    if sy-subrc eq 0.
      assign <ls_cell_data>-cell_value to <lv_value>.
    else.
      assign lc_empty_string to <lv_value>.
    endif.

    assert_equals( act = <lv_value>
                   exp = iv_value ).

  endmethod.                    "assert_value_equals
**
  method assert_formula_equals.

    field-symbols: <ls_cell_data> type zexcel_s_cell_data.

    read table worksheet->sheet_content assigning <ls_cell_data>
      with table key cell_row = iv_row cell_column = iv_col.
    assert_subrc( sy-subrc ).

    assert_equals( act = <ls_cell_data>-cell_formula
                   exp = iv_formula ).

  endmethod.                    "assert_formula_equals
*
  method assert_style_equals.

    field-symbols: <ls_cell_data> type zexcel_s_cell_data.

    read table worksheet->sheet_content assigning <ls_cell_data>
      with table key cell_row = iv_row cell_column = iv_col.
    assert_subrc( sy-subrc ).

    assert_equals( act = <ls_cell_data>-cell_style
                   exp = iv_style ).

  endmethod.
*
  method assert_datatype_equals.

    field-symbols: <ls_cell_data> type zexcel_s_cell_data.

    read table worksheet->sheet_content assigning <ls_cell_data>
      with table key cell_row = iv_row cell_column = iv_col.
    assert_subrc( sy-subrc ).

    assert_equals( act = <ls_cell_data>-data_type
                   exp = iv_datatype ).

  endmethod.                    "assert_datatype_equals
  method setup.
    create object out.
    create object excel.
    create object worksheet
      exporting
        ip_excel = excel.
  endmethod.                    "setup
endclass.                    "lcl_test IMPLEMENTATION
