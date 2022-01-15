CLASS ltc_normalize_columnrow_param DEFINITION DEFERRED.
CLASS ltc_normalize_range_param DEFINITION DEFERRED.
CLASS ltc_normalize_style_param DEFINITION DEFERRED.
CLASS ltc_check_cell_column_formula DEFINITION DEFERRED.
CLASS zcl_excel_worksheet DEFINITION LOCAL FRIENDS
    ltc_normalize_columnrow_param
    ltc_normalize_range_param
    ltc_normalize_style_param
    ltc_check_cell_column_formula.

CLASS lcl_excel_worksheet_test DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.
* ================
    DATA:
      f_cut TYPE REF TO zcl_excel_worksheet.  "class under test

    CLASS-METHODS: class_setup.
    CLASS-METHODS: class_teardown.
    METHODS: setup.
    METHODS: teardown.
    METHODS: set_merge FOR TESTING RAISING cx_static_check.
    METHODS: delete_merge FOR TESTING RAISING cx_static_check.
    METHODS: get_dimension_range FOR TESTING RAISING cx_static_check.
ENDCLASS.       "lcl_Excel_Worksheet_Test


CLASS ltc_check_cell_column_formula DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.

    METHODS: success FOR TESTING RAISING cx_static_check.
    METHODS: fails_both_formula_id_value FOR TESTING RAISING cx_static_check.
    METHODS: fails_both_formula_id_formula FOR TESTING RAISING cx_static_check.
    METHODS: formula_id_not_found FOR TESTING RAISING cx_static_check.
    METHODS: outside_table_fails__above FOR TESTING RAISING cx_static_check.
    METHODS: outside_table_fails__below FOR TESTING RAISING cx_static_check.
    METHODS: outside_table_fails__left FOR TESTING RAISING cx_static_check.
    METHODS: outside_table_fails__right FOR TESTING RAISING cx_static_check.
    METHODS: must_be_in_same_column FOR TESTING RAISING cx_static_check.

    METHODS: setup.
    METHODS: should_fail
      IMPORTING
        ip_formula_id TYPE zexcel_s_cell_data-column_formula_id
        ip_formula    TYPE zexcel_s_cell_data-cell_formula OPTIONAL
        ip_value      TYPE zexcel_s_cell_data-cell_value OPTIONAL
        ip_row        TYPE zexcel_s_cell_data-cell_row
        ip_column     TYPE zexcel_s_cell_data-cell_column
        ip_exp        TYPE string
      RAISING
        zcx_excel.

    DATA: mt_column_formulas TYPE zcl_excel_worksheet=>mty_th_column_formula,
          c_messages         LIKE zcl_excel_worksheet=>c_messages.

ENDCLASS.


CLASS ltc_normalize_columnrow_param DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.
    TYPES : BEGIN OF ty_parameters,
              BEGIN OF input,
                columnrow TYPE string,
                column    TYPE string,
                row       TYPE zexcel_cell_row,
              END OF input,
              BEGIN OF output,
                fails  TYPE abap_bool,
                column TYPE zexcel_cell_column,
                row    TYPE zexcel_cell_row,
              END OF output,
            END OF ty_parameters.
    DATA:
      cut TYPE REF TO zcl_excel_worksheet.  "class under test

    METHODS setup.
    METHODS:
      test FOR TESTING RAISING cx_static_check,
      all_parameters_passed FOR TESTING RAISING cx_static_check,
      none_parameter_passed FOR TESTING RAISING cx_static_check.

    METHODS assert
      IMPORTING
        input TYPE ty_parameters-input
        exp   TYPE ty_parameters-output
      RAISING
        cx_static_check.

ENDCLASS.


CLASS ltc_normalize_range_param DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.
    TYPES : BEGIN OF ty_parameters,
              BEGIN OF input,
                range        TYPE string,
                column_start TYPE string,
                column_end   TYPE string,
                row          TYPE zexcel_cell_row,
                row_to       TYPE zexcel_cell_row,
              END OF input,
              BEGIN OF output,
                fails        TYPE abap_bool,
                column_start TYPE zexcel_cell_column,
                column_end   TYPE zexcel_cell_column,
                row_start    TYPE zexcel_cell_row,
                row_end      TYPE zexcel_cell_row,
              END OF output,
            END OF ty_parameters.
    DATA:
      cut TYPE REF TO zcl_excel_worksheet.  "class under test

    METHODS setup.
    METHODS:
      range_one_cell FOR TESTING RAISING cx_static_check,
      relative_range FOR TESTING RAISING cx_static_check,
      invalid_range FOR TESTING RAISING cx_static_check,
      absolute_range FOR TESTING RAISING cx_static_check,
      reverse_range_not_supported FOR TESTING RAISING cx_static_check,
      all_parameters_passed FOR TESTING RAISING cx_static_check,
      none_parameter_passed FOR TESTING RAISING cx_static_check,
      start_without_end FOR TESTING RAISING cx_static_check.

    METHODS assert
      IMPORTING
        input TYPE ty_parameters-input
        exp   TYPE ty_parameters-output
      RAISING
        cx_static_check.

ENDCLASS.


CLASS ltc_normalize_style_param DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.
    TYPES : BEGIN OF ty_parameters_output,
              fails TYPE abap_bool,
              guid  TYPE zexcel_cell_style,
            END OF ty_parameters_output.
    DATA:
      excel TYPE REF TO zcl_excel,
      cut   TYPE REF TO zcl_excel_worksheet,  "class under test
      exp   TYPE ty_parameters_output.


    METHODS setup.
    METHODS:
      ref_to_zcl_excel_style FOR TESTING RAISING cx_static_check,
      zexcel_cell_style FOR TESTING RAISING cx_static_check,
      raw_16_bytes FOR TESTING RAISING cx_static_check,
      other FOR TESTING RAISING cx_static_check.

    METHODS assert
      IMPORTING
        input TYPE any
        exp   TYPE ty_parameters_output
      RAISING
        cx_static_check.

ENDCLASS.


CLASS lcl_excel_worksheet_test IMPLEMENTATION.
* ==============================================

  METHOD class_setup.
* ===================

  ENDMETHOD.       "class_Setup


  METHOD class_teardown.
* ======================


  ENDMETHOD.       "class_Teardown


  METHOD setup.
* =============

    DATA lo_excel TYPE REF TO zcl_excel.

    CREATE OBJECT lo_excel.

    TRY.
        CREATE OBJECT f_cut
          EXPORTING
            ip_excel = lo_excel.
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail( 'Could not create instance' ).
    ENDTRY.

  ENDMETHOD.       "setup


  METHOD teardown.
* ================


  ENDMETHOD.       "teardown

  METHOD set_merge.
* ====================

    DATA lt_merge TYPE string_table.
    DATA lv_merge TYPE string.
    DATA lv_size TYPE i.
    DATA lv_size_next TYPE i.


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

    TRY.
        f_cut->set_merge(
          ip_column_start = 2
          ip_column_end = 3
          ip_row = 2
          ip_row_to = 3
        ).
      CATCH zcx_excel.
    ENDTRY.

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
    READ TABLE lt_merge INTO lv_merge INDEX 1.

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
    READ TABLE lt_merge INTO lv_merge INDEX 1.

    cl_abap_unit_assert=>assert_equals(
      act   = lv_merge
      exp   = 'D4:E5'
      msg   = 'Expect D4:E5'
      level = if_aunit_constants=>critical
    ).

  ENDMETHOD.

  METHOD delete_merge.
* ====================
    DATA lt_merge TYPE string_table.
    DATA lv_merge TYPE string.
    DATA lv_size TYPE i.
    DATA lv_index TYPE i.

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

    DO 10 TIMES.
      f_cut->set_merge(
        ip_column_start = 2 + sy-index * 2
        ip_column_end = 3 + sy-index * 2
        ip_row = 2 + sy-index * 2
        ip_row_to = 3 + sy-index * 2
      ).
    ENDDO.

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

    DO 4 TIMES.
      lv_index = sy-index.

      f_cut->delete_merge( ).

      f_cut->set_merge(
        ip_column_start = 2
        ip_column_end = 3
        ip_row = 2
        ip_row_to = 3
      ).

      CASE lv_index.
        WHEN 1. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 2 ).
        WHEN 2. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 3 ).
        WHEN 3. f_cut->delete_merge( ip_cell_column = 3 ip_cell_row = 2 ).
        WHEN 4. f_cut->delete_merge( ip_cell_column = 3 ip_cell_row = 3 ).
      ENDCASE.

      lt_merge = f_cut->get_merge( ).
      lv_size = lines( lt_merge ).

      cl_abap_unit_assert=>assert_equals(
        act   = lv_size
        exp   = 0
        msg   = 'Expect merge table with 1 line fully cleared'
        level = if_aunit_constants=>critical
      ).
    ENDDO.

*  Test 4. Delete concrete merge with fail

    DO 4 TIMES.
      lv_index = sy-index.

      f_cut->delete_merge( ).

      f_cut->set_merge(
        ip_column_start = 2
        ip_column_end = 3
        ip_row = 2
        ip_row_to = 3
      ).

      CASE lv_index.
        WHEN 1. f_cut->delete_merge( ip_cell_column = 1 ip_cell_row = 2 ).
        WHEN 2. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 1 ).
        WHEN 3. f_cut->delete_merge( ip_cell_column = 4 ip_cell_row = 2 ).
        WHEN 4. f_cut->delete_merge( ip_cell_column = 2 ip_cell_row = 4 ).
      ENDCASE.

      lt_merge = f_cut->get_merge( ).
      lv_size = lines( lt_merge ).

      cl_abap_unit_assert=>assert_equals(
        act   = lv_size
        exp   = 1
        msg   = 'Expect no merge were deleted'
        level = if_aunit_constants=>critical
      ).
    ENDDO.

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

    READ TABLE lt_merge INTO lv_merge INDEX 1.

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

    READ TABLE lt_merge INTO lv_merge INDEX 1.

    cl_abap_unit_assert=>assert_equals(
      act   = lv_merge
      exp   = 'B2:C3'
      msg   = 'Expect delete D4:E5 merge'
      level = if_aunit_constants=>critical
    ).

  ENDMETHOD.       "delete_Merge

  METHOD get_dimension_range.
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
  ENDMETHOD.

ENDCLASS.       "lcl_Excel_Worksheet_Test


CLASS ltc_check_cell_column_formula IMPLEMENTATION.

  METHOD setup.

    DATA: ls_column_formula  TYPE zcl_excel_worksheet=>mty_s_column_formula.

    c_messages = zcl_excel_worksheet=>c_messages.

    " Column Formula in table A1:B4 (for unknown reason, bottom_right_row is last actual row minus 1)
    CLEAR ls_column_formula.
    ls_column_formula-id = 1.
    ls_column_formula-column = 1.
    ls_column_formula-table_top_left_row = 1.
    ls_column_formula-table_bottom_right_row = 3.
    ls_column_formula-table_left_column_int = 1.
    ls_column_formula-table_right_column_int = 2.
    INSERT ls_column_formula INTO TABLE mt_column_formulas.

    " Column Formula in table D1:E4 (for unknown reason, bottom_right_row is last actual row minus 1)
    CLEAR ls_column_formula.
    ls_column_formula-id = 2.
    ls_column_formula-column = 4.
    ls_column_formula-table_top_left_row = 1.
    ls_column_formula-table_bottom_right_row = 3.
    ls_column_formula-table_left_column_int = 4.
    ls_column_formula-table_right_column_int = 5.
    INSERT ls_column_formula INTO TABLE mt_column_formulas.

  ENDMETHOD.

  METHOD success.

    zcl_excel_worksheet=>check_cell_column_formula(
        it_column_formulas   = mt_column_formulas
        ip_column_formula_id = 1
        ip_formula           = ''
        ip_value             = ''
        ip_row               = 2
        ip_column            = 1 ).

  ENDMETHOD.

  METHOD fails_both_formula_id_value.
    should_fail( ip_formula_id = 1 ip_formula = '' ip_value = '3.14' ip_row = 2 ip_column = 1 ip_exp = c_messages-formula_id_only_is_possible ).
  ENDMETHOD.

  METHOD fails_both_formula_id_formula.
    should_fail( ip_formula_id = 1 ip_formula = 'A2' ip_value = '' ip_row = 2 ip_column = 1 ip_exp = c_messages-formula_id_only_is_possible ).
  ENDMETHOD.

  METHOD formula_id_not_found.
    should_fail( ip_formula_id = 3 ip_row = 1 ip_column = 1 ip_exp = c_messages-column_formula_id_not_found ).
  ENDMETHOD.

  METHOD outside_table_fails__above.
    should_fail( ip_formula_id = 2 ip_row = 1 ip_column = 1 ip_exp = c_messages-formula_not_in_this_table ).
  ENDMETHOD.

  METHOD outside_table_fails__below.
    should_fail( ip_formula_id = 2 ip_row = 5 ip_column = 1 ip_exp = c_messages-formula_not_in_this_table ).
  ENDMETHOD.

  METHOD outside_table_fails__left.
    should_fail( ip_formula_id = 2 ip_row = 2 ip_column = 0 ip_exp = c_messages-formula_not_in_this_table ).
  ENDMETHOD.

  METHOD outside_table_fails__right.
    should_fail( ip_formula_id = 2 ip_row = 2 ip_column = 3 ip_exp = c_messages-formula_not_in_this_table ).
  ENDMETHOD.

  METHOD must_be_in_same_column.
    should_fail( ip_formula_id = 1 ip_row = 2 ip_column = 2 ip_exp = c_messages-formula_in_other_column ).
  ENDMETHOD.

  METHOD should_fail.

    DATA: lo_exception TYPE REF TO zcx_excel.

    TRY.
        zcl_excel_worksheet=>check_cell_column_formula(
          EXPORTING
            it_column_formulas   = mt_column_formulas
            ip_column_formula_id = ip_formula_id
            ip_formula           = ip_formula
            ip_value             = ip_value
            ip_row               = ip_row
            ip_column            = ip_column ).
        cl_abap_unit_assert=>fail( msg = |Should have failed with error "{ ip_exp }"| ).
      CATCH zcx_excel INTO lo_exception.
        cl_abap_unit_assert=>assert_equals( act = lo_exception->get_text( ) exp = ip_exp msg = ip_exp ).
    ENDTRY.

  ENDMETHOD.

ENDCLASS.


CLASS ltc_normalize_columnrow_param IMPLEMENTATION.

  METHOD setup.

    DATA: lo_excel TYPE REF TO zcl_excel.

    CREATE OBJECT lo_excel.

    TRY.
        CREATE OBJECT cut
          EXPORTING
            ip_excel = lo_excel.
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail( 'Could not create instance' ).
    ENDTRY.

  ENDMETHOD.

  METHOD test.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    input-columnrow = 'B4'.
    exp-column = 2.
    exp-row = 4.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD all_parameters_passed.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    input-columnrow = 'B4'.
    input-column = 'B'.
    input-row = 4.
    exp-fails = abap_true.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD none_parameter_passed.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    exp-fails = abap_true.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD assert.
    DATA: act        TYPE ty_parameters-output,
          error      TYPE REF TO zcx_excel,
          input_text TYPE string,
          message    TYPE string.

    TRY.

        cut->normalize_columnrow_parameter( EXPORTING ip_columnrow = input-columnrow
                                                      ip_column    = input-column
                                                      ip_row       = input-row
                                            IMPORTING ep_column    = act-column
                                                      ep_row       = act-row ).
        IF exp-fails = abap_true.
          message = |Should have failed for { input_text }|.
          cl_abap_unit_assert=>fail( msg = message ).
        ENDIF.

      CATCH zcx_excel INTO error.
        IF exp-fails = abap_false.
          RAISE EXCEPTION error.
        ENDIF.
        RETURN.
    ENDTRY.

    input_text = |input column/row { input-columnrow }, column { input-column }, row { input-row }|.
    message = |Invalid column for { input_text }|.
    cl_abap_unit_assert=>assert_equals( msg = message exp = exp-column act = act-column ).
    message = |Invalid row for { input_text }|.
    cl_abap_unit_assert=>assert_equals( msg = message exp = exp-row    act = act-row ).
  ENDMETHOD.

ENDCLASS.


CLASS ltc_normalize_range_param IMPLEMENTATION.

  METHOD setup.

    DATA: lo_excel TYPE REF TO zcl_excel.

    CREATE OBJECT lo_excel.

    TRY.
        CREATE OBJECT cut
          EXPORTING
            ip_excel = lo_excel.
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail( 'Could not create instance' ).
    ENDTRY.

  ENDMETHOD.

  METHOD range_one_cell.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    CLEAR: input, exp.
    input-range = 'B4:B4'.
    exp-column_start = 2.
    exp-column_end = 2.
    exp-row_start = 4.
    exp-row_end = 4.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD relative_range.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    CLEAR: input, exp.
    input-range = 'B4:AA10'.
    exp-column_start = 2.
    exp-column_end = 27.
    exp-row_start = 4.
    exp-row_end = 10.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD absolute_range.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    CLEAR: input, exp.
    input-range = '$B$4:$AA$10'.
    exp-column_start = 2.
    exp-column_end = 27.
    exp-row_start = 4.
    exp-row_end = 10.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD invalid_range.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    CLEAR: input, exp.
    input-range = 'B4'.
    exp-fails = abap_true.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD reverse_range_not_supported.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    CLEAR: input, exp.
    input-range = 'B4:A1'.
    exp-fails = abap_true.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD all_parameters_passed.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    input-range = 'B4:B4'.
    input-column_start = 'B'.
    input-row = 4.
    exp-fails = abap_true.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD none_parameter_passed.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    exp-fails = abap_true.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD start_without_end.
    DATA: input TYPE ty_parameters-input,
          exp   TYPE ty_parameters-output.

    input-column_start = 'B'.
    input-row = 4.
    exp-column_start = 2.
    exp-column_end = 2.
    exp-row_start = 4.
    exp-row_end = 4.
    assert( input = input exp = exp ).

  ENDMETHOD.

  METHOD assert.
    DATA: act        TYPE ty_parameters-output,
          error      TYPE REF TO zcx_excel,
          input_text TYPE string,
          message    TYPE string.

    input_text = |input range { input-range }, column start { input-column_start }, column end { input-column_end }|.

    TRY.

        cut->normalize_range_parameter( EXPORTING ip_range        = input-range
                                                  ip_column_start = input-column_start  ip_column_end = input-column_end
                                                  ip_row          = input-row           ip_row_to     = input-row_to
                                        IMPORTING ep_column_start = act-column_start    ep_column_end = act-column_end
                                                  ep_row          = act-row_start       ep_row_to     = act-row_end ).
        IF exp-fails = abap_true.
          message = |Should have failed for { input_text }|.
          cl_abap_unit_assert=>fail( msg = message ).
        ENDIF.

      CATCH zcx_excel INTO error.
        IF exp-fails = abap_false.
          RAISE EXCEPTION error.
        ENDIF.
        RETURN.
    ENDTRY.

    message = |Invalid column start for { input_text }|.
    cl_abap_unit_assert=>assert_equals( msg = message exp = exp-column_start  act = act-column_start ).
    message = |Invalid column end for { input_text }|.
    cl_abap_unit_assert=>assert_equals( msg = message exp = exp-column_end    act = act-column_end ).
    message = |Invalid row start for { input_text }|.
    cl_abap_unit_assert=>assert_equals( msg = message exp = exp-row_start     act = act-row_start ).
    message = |Invalid row start for { input_text }|.
    cl_abap_unit_assert=>assert_equals( msg = message exp = exp-row_end       act = act-row_end ).

  ENDMETHOD.

ENDCLASS.


CLASS ltc_normalize_style_param IMPLEMENTATION.

  METHOD setup.

    CREATE OBJECT excel.

    TRY.
        CREATE OBJECT cut
          EXPORTING
            ip_excel = excel.
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail( 'Could not create instance' ).
    ENDTRY.

  ENDMETHOD.

  METHOD ref_to_zcl_excel_style.
    DATA: style TYPE REF TO zcl_excel_style.

    style = excel->add_new_style( ).
    exp-fails = abap_false.
    exp-guid = style->get_guid( ).
    assert( input = style exp = exp ).

  ENDMETHOD.

  METHOD zexcel_cell_style.
    DATA: style TYPE REF TO zcl_excel_style.

    style = excel->add_new_style( ).
    exp-fails = abap_false.
    exp-guid = style->get_guid( ).
    assert( input = exp-guid exp = exp ).

  ENDMETHOD.

  METHOD raw_16_bytes.
    DATA: raw_16_bytes TYPE x LENGTH 16.

    raw_16_bytes = 'A1A1A1A1A1A1A1A1A1A1A1A1A1A1A1A1'.
    exp-fails = abap_false.
    exp-guid = raw_16_bytes.
    assert( input = raw_16_bytes exp = exp ).

  ENDMETHOD.

  METHOD other.

    exp-fails = abap_true.
    assert( input = 'A' exp = exp ).

  ENDMETHOD.

  METHOD assert.
    DATA: rtti       TYPE REF TO cl_abap_typedescr,
          message    TYPE string,
          act        TYPE zexcel_cell_style,
          error      TYPE REF TO zcx_excel,
          input_text TYPE string.

    rtti = cl_abap_typedescr=>describe_by_data( input ).
    IF rtti->type_kind = rtti->typekind_oref.
      rtti = cl_abap_typedescr=>describe_by_object_ref( input ).
      input_text = |REF TO { rtti->absolute_name }|.
    ELSE.
      input_text = rtti->absolute_name.
    ENDIF.

    TRY.

        act = cut->normalize_style_parameter( input ).

        IF exp-fails = abap_true.
          message = |Input type { input_text } accepted but only REF TO \\CLASS=ZCL_EXCEL_STYLE, \\TYPE=ZEXCEL_CELL_STYLE or any X type are supported|.
          cl_abap_unit_assert=>fail( msg = message ).
        ENDIF.

      CATCH zcx_excel INTO error.
        IF exp-fails = abap_false.
          message = |Input type { input_text } made it fail but it should succeed with result { exp-guid }|.
          cl_abap_unit_assert=>fail( msg = message ).
        ENDIF.
        RETURN.
    ENDTRY.

    cl_abap_unit_assert=>assert_equals( exp = exp-guid act = act ).

  ENDMETHOD.

ENDCLASS.
