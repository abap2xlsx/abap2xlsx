
CLASS lcl_excel_worksheet_test DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT
.
*?﻿<asx:abap xmlns:asx="http://www.sap.com/abapxml" version="1.0">
*?<asx:values>
*?<TESTCLASS_OPTIONS>
*?<TEST_CLASS>lcl_Excel_Worksheet_Test
*?</TEST_CLASS>
*?<TEST_MEMBER>f_Cut
*?</TEST_MEMBER>
*?<OBJECT_UNDER_TEST>ZCL_EXCEL_WORKSHEET
*?</OBJECT_UNDER_TEST>
*?<OBJECT_IS_LOCAL/>
*?<GENERATE_FIXTURE>X
*?</GENERATE_FIXTURE>
*?<GENERATE_CLASS_FIXTURE>X
*?</GENERATE_CLASS_FIXTURE>
*?<GENERATE_INVOCATION>X
*?</GENERATE_INVOCATION>
*?<GENERATE_ASSERT_EQUAL>X
*?</GENERATE_ASSERT_EQUAL>
*?</TESTCLASS_OPTIONS>
*?</asx:values>
*?</asx:abap>
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
        cl_abap_unit_assert=>fail( 'Instantiation failed' ).
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

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
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

      zcl_excel_aunit=>assert_equals(
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

      zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
      act   = lv_size
      exp   = 1
      msg   = 'Expect we have the one merge'
      level = if_aunit_constants=>critical
    ).

    READ TABLE lt_merge INTO lv_merge INDEX 1.

    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
      act   = lv_size
      exp   = 1
      msg   = 'Expect we have the one merge'
      level = if_aunit_constants=>critical
    ).

    READ TABLE lt_merge INTO lv_merge INDEX 1.

    zcl_excel_aunit=>assert_equals(
      act   = lv_merge
      exp   = 'B2:C3'
      msg   = 'Expect delete D4:E5 merge'
      level = if_aunit_constants=>critical
    ).

  ENDMETHOD.       "delete_Merge

  METHOD get_dimension_range.
    zcl_excel_aunit=>assert_equals(
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

    zcl_excel_aunit=>assert_equals(
      act   = f_cut->get_dimension_range( )
      exp   = 'C2:F5'
      msg   = 'get_dimension_range'
      level = if_aunit_constants=>critical
    ).
  ENDMETHOD.

ENDCLASS.       "lcl_Excel_Worksheet_Test
