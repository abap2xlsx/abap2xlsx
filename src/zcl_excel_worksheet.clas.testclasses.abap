class ltc_check_cell_column_formula definition deferred.
class zcl_excel_worksheet definition local friends
    ltc_check_cell_column_formula.

class lcl_excel_worksheet_test definition for testing
    risk level harmless
    duration short.

  private section.
* ================
    data:
      f_cut type ref to zcl_excel_worksheet.  "class under test

    class-methods: class_setup.
    class-methods: class_teardown.
    methods: setup.
    methods: teardown.
    methods: set_merge for testing raising cx_static_check.
    methods: delete_merge for testing raising cx_static_check.
    methods: get_dimension_range for testing raising cx_static_check.
endclass.       "lcl_Excel_Worksheet_Test


class ltc_check_cell_column_formula definition for testing
    risk level harmless
    duration short.

  private section.

    methods: success for testing raising cx_static_check.
    methods: fails_both_formula_id_value for testing raising cx_static_check.
    methods: fails_both_formula_id_formula for testing raising cx_static_check.
    methods: formula_id_not_found for testing raising cx_static_check.
    methods: outside_table_fails__above for testing raising cx_static_check.
    methods: outside_table_fails__below for testing raising cx_static_check.
    methods: outside_table_fails__left for testing raising cx_static_check.
    methods: outside_table_fails__right for testing raising cx_static_check.
    methods: must_be_in_same_column for testing raising cx_static_check.

    methods: setup.
    methods: should_fail
      importing
        ip_formula_id type zexcel_s_cell_data-column_formula_id
        ip_formula    type zexcel_s_cell_data-cell_formula optional
        ip_value      type zexcel_s_cell_data-cell_value optional
        ip_row        type zexcel_s_cell_data-cell_row
        ip_column     type zexcel_s_cell_data-cell_column
        ip_exp        type string
      raising
        zcx_excel.

    data: mt_column_formulas type zcl_excel_worksheet=>mty_th_column_formula,
          c_messages         like zcl_excel_worksheet=>c_messages.

endclass.


class lcl_excel_worksheet_test implementation.
* ==============================================

  method class_setup.
* ===================

  endmethod.       "class_Setup


  method class_teardown.
* ======================


  endmethod.       "class_Teardown


  method setup.
* =============

    data lo_excel type ref to zcl_excel.

    create object lo_excel.

    try.
        create object f_cut
          exporting
            ip_excel = lo_excel.
      catch zcx_excel.
      cl_abap_unit_assert=>fail( 'Could not create instance' ).
    endtry.

  endmethod.       "setup


  method teardown.
* ================


  endmethod.       "teardown

  method set_merge.
* ====================

    data lt_merge type string_table.
    data lv_merge type string.
    data lv_size type i.
    data lv_size_next type i.


*  Test 1. Simple test for initial value

    lt_merge = f_cut->get_merge( ).
    lv_size = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size
      exp   = 0
      msg   = 'Initial state of merge table is not empty'
      level = if_aunit_constants=>critical
    ).


* Test 2. Add merge

    f_cut->set_merge(
      ip_column_start = 2
      ip_column_end = 3
      ip_row = 2
      ip_row_to = 3
    ).

    lt_merge = f_cut->get_merge( ).
    lv_size_next = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size_next - lv_size
      exp   = 1
      msg   = 'Expect add 1 table line when 1 merge added'
      level = if_aunit_constants=>critical
    ).

* Test 2. Add same merge

    lv_size = lv_size_next.

    try.
        f_cut->set_merge(
          ip_column_start = 2
          ip_column_end = 3
          ip_row = 2
          ip_row_to = 3
        ).
      catch zcx_excel.
    endtry.

    lt_merge = f_cut->get_merge( ).
    lv_size_next = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size_next - lv_size
      exp   = 0
      msg   = 'Expect no change when add same merge'
      level = if_aunit_constants=>critical
    ).

* Test 3. Add one different merge

    lv_size = lv_size_next.

    f_cut->set_merge(
      ip_column_start = 4
      ip_column_end = 5
      ip_row = 2
      ip_row_to = 3
    ).

    lt_merge = f_cut->get_merge( ).
    lv_size_next = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size_next - lv_size
      exp   = 1
      msg   = 'Expect 1 change when add different merge'
      level = if_aunit_constants=>critical
    ).

* Test 4. Merge added with concrete value #1

    f_cut->delete_merge( ).

    f_cut->set_merge(
      ip_column_start = 2
      ip_column_end = 3
      ip_row = 2
      ip_row_to = 3
    ).

    lt_merge = f_cut->get_merge( ).
    read table lt_merge into lv_merge index 1.

    cl_abap_unit_assert=>assert_equals(
      act   = lv_merge
      exp   = 'B2:C3'
      msg   = 'Expect B2:C3'
      level = if_aunit_constants=>critical
    ).

* Test 5. Merge added with concrete value #2

    f_cut->delete_merge( ).

    f_cut->set_merge(
      ip_column_start = 4
      ip_column_end = 5
      ip_row = 4
      ip_row_to = 5
    ).

    lt_merge = f_cut->get_merge( ).
    read table lt_merge into lv_merge index 1.

    cl_abap_unit_assert=>assert_equals(
      act   = lv_merge
      exp   = 'D4:E5'
      msg   = 'Expect D4:E5'
      level = if_aunit_constants=>critical
    ).

  endmethod.

  method delete_merge.
* ====================
    data lt_merge type string_table.
    data lv_merge type string.
    data lv_size type i.
    data lv_index type i.

*  Test 1. Simple test delete all merges

    f_cut->set_merge(
      ip_column_start = 2
      ip_column_end = 3
      ip_row = 2
      ip_row_to = 3
    ).

    f_cut->delete_merge( ).
    lt_merge = f_cut->get_merge( ).
    lv_size = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size
      exp   = 0
      msg   = 'Expect merge table with 1 line fully cleared'
      level = if_aunit_constants=>critical
    ).

*  Test 2. Simple test delete all merges

    do 10 times.
      f_cut->set_merge(
        ip_column_start = 2 + sy-index * 2
        ip_column_end = 3 + sy-index * 2
        ip_row = 2 + sy-index * 2
        ip_row_to = 3 + sy-index * 2
      ).
    enddo.

    f_cut->delete_merge( ).
    lt_merge = f_cut->get_merge( ).
    lv_size = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size
      exp   = 0
      msg   = 'Expect merge table with few lines fully cleared'
      level = if_aunit_constants=>critical
    ).

*  Test 3. Delete concrete merge with success

    do 4 times.
      lv_index = sy-index.

      f_cut->delete_merge( ).

      f_cut->set_merge(
        ip_column_start = 2
        ip_column_end = 3
        ip_row = 2
        ip_row_to = 3
      ).

      case lv_index.
        when 1. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 2 ).
        when 2. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 3 ).
        when 3. f_cut->delete_merge( ip_cell_column = 3 ip_cell_row = 2 ).
        when 4. f_cut->delete_merge( ip_cell_column = 3 ip_cell_row = 3 ).
      endcase.

      lt_merge = f_cut->get_merge( ).
      lv_size = lines( lt_merge ).

      cl_abap_unit_assert=>assert_equals(
        act   = lv_size
        exp   = 0
        msg   = 'Expect merge table with 1 line fully cleared'
        level = if_aunit_constants=>critical
      ).
    enddo.

*  Test 4. Delete concrete merge with fail

    do 4 times.
      lv_index = sy-index.

      f_cut->delete_merge( ).

      f_cut->set_merge(
        ip_column_start = 2
        ip_column_end = 3
        ip_row = 2
        ip_row_to = 3
      ).

      case lv_index.
        when 1. f_cut->delete_merge( ip_cell_column = 1 ip_cell_row = 2 ).
        when 2. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 1 ).
        when 3. f_cut->delete_merge( ip_cell_column = 4 ip_cell_row = 2 ).
        when 4. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 4 ).
      endcase.

      lt_merge = f_cut->get_merge( ).
      lv_size = lines( lt_merge ).

      cl_abap_unit_assert=>assert_equals(
        act   = lv_size
        exp   = 1
        msg   = 'Expect no merge were deleted'
        level = if_aunit_constants=>critical
      ).
    enddo.

*  Test 5. Delete concrete merge #1

    f_cut->delete_merge( ).

    f_cut->set_merge(
      ip_column_start = 2
      ip_column_end = 3
      ip_row = 2
      ip_row_to = 3
    ).
    f_cut->set_merge(
      ip_column_start = 4
      ip_column_end = 5
      ip_row = 4
      ip_row_to = 5
    ).

    f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 2 ).
    lt_merge = f_cut->get_merge( ).
    lv_size = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size
      exp   = 1
      msg   = 'Expect we have the one merge'
      level = if_aunit_constants=>critical
    ).

    read table lt_merge into lv_merge index 1.

    cl_abap_unit_assert=>assert_equals(
      act   = lv_merge
      exp   = 'D4:E5'
      msg   = 'Expect delete B2:C3 merge'
      level = if_aunit_constants=>critical
    ).

*  Test 6. Delete concrete merge #2

    f_cut->delete_merge( ).

    f_cut->set_merge(
      ip_column_start = 2
      ip_column_end = 3
      ip_row = 2
      ip_row_to = 3
    ).
    f_cut->set_merge(
      ip_column_start = 4
      ip_column_end = 5
      ip_row = 4
      ip_row_to = 5
    ).

    f_cut->delete_merge( ip_cell_column = 4 ip_cell_row = 4 ).
    lt_merge = f_cut->get_merge( ).
    lv_size = lines( lt_merge ).

    cl_abap_unit_assert=>assert_equals(
      act   = lv_size
      exp   = 1
      msg   = 'Expect we have the one merge'
      level = if_aunit_constants=>critical
    ).

    read table lt_merge into lv_merge index 1.

    cl_abap_unit_assert=>assert_equals(
      act   = lv_merge
      exp   = 'B2:C3'
      msg   = 'Expect delete D4:E5 merge'
      level = if_aunit_constants=>critical
    ).

  endmethod.       "delete_Merge

  method get_dimension_range.
    cl_abap_unit_assert=>assert_equals(
      act   = f_cut->get_dimension_range( )
      exp   = 'A1'
      msg   = 'get_dimension_range inital value'
      level = if_aunit_constants=>critical
    ).

    f_cut->set_cell(
      ip_row       = 2
      ip_column    = 3
      ip_value     = 'Dummy'
    ).

    f_cut->set_cell(
      ip_row       = 5
      ip_column    = 6
      ip_value     = 'Dummy'
    ).

    cl_abap_unit_assert=>assert_equals(
      act   = f_cut->get_dimension_range( )
      exp   = 'C2:F5'
      msg   = 'get_dimension_range'
      level = if_aunit_constants=>critical
    ).
  endmethod.

endclass.       "lcl_Excel_Worksheet_Test


class ltc_check_cell_column_formula implementation.

  method setup.

    data: ls_column_formula  type zcl_excel_worksheet=>mty_s_column_formula.

    c_messages = zcl_excel_worksheet=>c_messages.

    " Column Formula in table A1:B4 (for unknown reason, bottom_right_row is last actual row minus 1)
    clear ls_column_formula.
    ls_column_formula-id = 1.
    ls_column_formula-column = 1.
    ls_column_formula-table_top_left_row = 1.
    ls_column_formula-table_bottom_right_row = 3.
    ls_column_formula-table_left_column_int = 1.
    ls_column_formula-table_right_column_int = 2.
    insert ls_column_formula into table mt_column_formulas.

    " Column Formula in table D1:E4 (for unknown reason, bottom_right_row is last actual row minus 1)
    clear ls_column_formula.
    ls_column_formula-id = 2.
    ls_column_formula-column = 4.
    ls_column_formula-table_top_left_row = 1.
    ls_column_formula-table_bottom_right_row = 3.
    ls_column_formula-table_left_column_int = 4.
    ls_column_formula-table_right_column_int = 5.
    insert ls_column_formula into table mt_column_formulas.

  endmethod.

  method success.

    zcl_excel_worksheet=>check_cell_column_formula(
        it_column_formulas   = mt_column_formulas
        ip_column_formula_id = 1
        ip_formula           = ''
        ip_value             = ''
        ip_row               = 2
        ip_column            = 1 ).

  endmethod.

  method fails_both_formula_id_value.
    should_fail( ip_formula_id = 1 ip_formula = '' ip_value = '3.14' ip_row = 2 ip_column = 1 ip_exp = c_messages-formula_id_only_is_possible ).
  endmethod.

  method fails_both_formula_id_formula.
    should_fail( ip_formula_id = 1 ip_formula = 'A2' ip_value = '' ip_row = 2 ip_column = 1 ip_exp = c_messages-formula_id_only_is_possible ).
  endmethod.

  method formula_id_not_found.
    should_fail( ip_formula_id = 3 ip_row = 1 ip_column = 1 ip_exp = c_messages-column_formula_id_not_found ).
  endmethod.

  method outside_table_fails__above.
    should_fail( ip_formula_id = 2 ip_row = 1 ip_column = 1 ip_exp = c_messages-formula_not_in_this_table ).
  endmethod.

  method outside_table_fails__below.
    should_fail( ip_formula_id = 2 ip_row = 5 ip_column = 1 ip_exp = c_messages-formula_not_in_this_table ).
  endmethod.

  method outside_table_fails__left.
    should_fail( ip_formula_id = 2 ip_row = 2 ip_column = 0 ip_exp = c_messages-formula_not_in_this_table ).
  endmethod.

  method outside_table_fails__right.
    should_fail( ip_formula_id = 2 ip_row = 2 ip_column = 3 ip_exp = c_messages-formula_not_in_this_table ).
  endmethod.

  method must_be_in_same_column.
    should_fail( ip_formula_id = 1 ip_row = 2 ip_column = 2 ip_exp = c_messages-formula_in_other_column ).
  endmethod.

  method should_fail.

    data: lo_exception type ref to zcx_excel.

    try.
        zcl_excel_worksheet=>check_cell_column_formula(
          exporting
            it_column_formulas   = mt_column_formulas
            ip_column_formula_id = ip_formula_id
            ip_formula           = ip_formula
            ip_value             = ip_value
            ip_row               = ip_row
            ip_column            = ip_column ).
        cl_abap_unit_assert=>fail( msg = |Should have failed with error "{ ip_exp }"| ).
      catch zcx_excel into lo_exception.
        cl_abap_unit_assert=>assert_equals( act = lo_exception->get_text( ) exp = ip_exp msg = ip_exp ).
    endtry.

  endmethod.

endclass.
