CLASS ltc_check_cell_column_formula DEFINITION DEFERRED.
CLASS zcl_excel_worksheet DEFINITION LOCAL FRIENDS
    ltc_check_cell_column_formula.

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
      F_CUT TYPE REF TO ZCL_EXCEL_WORKSHEET.  "class under test

    CLASS-METHODS: CLASS_SETUP.
    CLASS-METHODS: CLASS_TEARDOWN.
    METHODS: SETUP.
    METHODS: TEARDOWN.
    METHODS: SET_MERGE FOR TESTING.
    METHODS: DELETE_MERGE FOR TESTING.
    METHODS: GET_DIMENSION_RANGE FOR TESTING.
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
          c_messages LIKE zcl_excel_worksheet=>c_messages.

ENDCLASS.


CLASS LCL_EXCEL_WORKSHEET_TEST IMPLEMENTATION.
* ==============================================

  METHOD CLASS_SETUP.
* ===================

  ENDMETHOD.       "class_Setup


  METHOD CLASS_TEARDOWN.
* ======================


  ENDMETHOD.       "class_Teardown


  METHOD SETUP.
* =============

    DATA LO_EXCEL TYPE REF TO ZCL_EXCEL.

    CREATE OBJECT LO_EXCEL.

    CREATE OBJECT F_CUT
      EXPORTING IP_EXCEL = LO_EXCEL.

  ENDMETHOD.       "setup


  METHOD TEARDOWN.
* ================


  ENDMETHOD.       "teardown

  METHOD SET_MERGE.
* ====================

    DATA LT_MERGE TYPE STRING_TABLE.
    DATA LV_MERGE TYPE STRING.
    DATA LV_SIZE TYPE I.
    DATA LV_SIZE_NEXT TYPE I.


*  Test 1. Simple test for initial value

    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE
      EXP   = 0
      MSG   = 'Initial state of merge table is not empty'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).


* Test 2. Add merge

    F_CUT->SET_MERGE(
      IP_COLUMN_START = 2
      IP_COLUMN_END = 3
      IP_ROW = 2
      IP_ROW_TO = 3
    ).

    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE_NEXT = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE_NEXT - LV_SIZE
      EXP   = 1
      MSG   = 'Expect add 1 table line when 1 merge added'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

* Test 2. Add same merge

    LV_SIZE = LV_SIZE_NEXT.

    TRY.
      F_CUT->SET_MERGE(
        IP_COLUMN_START = 2
        IP_COLUMN_END = 3
        IP_ROW = 2
        IP_ROW_TO = 3
      ).
    CATCH ZCX_EXCEL.
    ENDTRY.

    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE_NEXT = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE_NEXT - LV_SIZE
      EXP   = 0
      MSG   = 'Expect no change when add same merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

* Test 3. Add one different merge

    LV_SIZE = LV_SIZE_NEXT.

    F_CUT->SET_MERGE(
      IP_COLUMN_START = 4
      IP_COLUMN_END = 5
      IP_ROW = 2
      IP_ROW_TO = 3
    ).

    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE_NEXT = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE_NEXT - LV_SIZE
      EXP   = 1
      MSG   = 'Expect 1 change when add different merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

* Test 4. Merge added with concrete value #1

    F_CUT->DELETE_MERGE( ).

    F_CUT->SET_MERGE(
      IP_COLUMN_START = 2
      IP_COLUMN_END = 3
      IP_ROW = 2
      IP_ROW_TO = 3
    ).

    LT_MERGE = F_CUT->GET_MERGE( ).
    READ TABLE LT_MERGE INTO LV_MERGE INDEX 1.

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_MERGE
      EXP   = 'B2:C3'
      MSG   = 'Expect B2:C3'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

* Test 5. Merge added with concrete value #2

    F_CUT->DELETE_MERGE( ).

    F_CUT->SET_MERGE(
      IP_COLUMN_START = 4
      IP_COLUMN_END = 5
      IP_ROW = 4
      IP_ROW_TO = 5
    ).

    LT_MERGE = F_CUT->GET_MERGE( ).
    READ TABLE LT_MERGE INTO LV_MERGE INDEX 1.

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_MERGE
      EXP   = 'D4:E5'
      MSG   = 'Expect D4:E5'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

  ENDMETHOD.

  METHOD DELETE_MERGE.
* ====================
    DATA LT_MERGE TYPE STRING_TABLE.
    DATA LV_MERGE TYPE STRING.
    DATA LV_SIZE TYPE I.
    DATA LV_INDEX TYPE I.

*  Test 1. Simple test delete all merges

    F_CUT->SET_MERGE(
      IP_COLUMN_START = 2
      IP_COLUMN_END = 3
      IP_ROW = 2
      IP_ROW_TO = 3
    ).

    F_CUT->DELETE_MERGE( ).
    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE
      EXP   = 0
      MSG   = 'Expect merge table with 1 line fully cleared'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

*  Test 2. Simple test delete all merges

    DO 10 TIMES.
      F_CUT->SET_MERGE(
        IP_COLUMN_START = 2 + SY-INDEX * 2
        IP_COLUMN_END = 3 + SY-INDEX * 2
        IP_ROW = 2 + SY-INDEX * 2
        IP_ROW_TO = 3 + SY-INDEX * 2
      ).
    ENDDO.

    F_CUT->DELETE_MERGE( ).
    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE
      EXP   = 0
      MSG   = 'Expect merge table with few lines fully cleared'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

*  Test 3. Delete concrete merge with success

    DO 4 TIMES.
      LV_INDEX = SY-INDEX.

      F_CUT->DELETE_MERGE( ).

      F_CUT->SET_MERGE(
        IP_COLUMN_START = 2
        IP_COLUMN_END = 3
        IP_ROW = 2
        IP_ROW_TO = 3
      ).

      CASE LV_INDEX.
        WHEN 1. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 2 IP_CELL_ROW = 2 ).
        WHEN 2. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 2 IP_CELL_ROW = 3 ).
        WHEN 3. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 3 IP_CELL_ROW = 2 ).
        WHEN 4. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 3 IP_CELL_ROW = 3 ).
      ENDCASE.

      LT_MERGE = F_CUT->GET_MERGE( ).
      LV_SIZE = LINES( LT_MERGE ).

      CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
        ACT   = LV_SIZE
        EXP   = 0
        MSG   = 'Expect merge table with 1 line fully cleared'
        LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
      ).
    ENDDO.

*  Test 4. Delete concrete merge with fail

    DO 4 TIMES.
      LV_INDEX = SY-INDEX.

      F_CUT->DELETE_MERGE( ).

      F_CUT->SET_MERGE(
        IP_COLUMN_START = 2
        IP_COLUMN_END = 3
        IP_ROW = 2
        IP_ROW_TO = 3
      ).

      CASE LV_INDEX.
        WHEN 1. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 1 IP_CELL_ROW = 2 ).
        WHEN 2. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 2 IP_CELL_ROW = 1 ).
        WHEN 3. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 4 IP_CELL_ROW = 2 ).
        WHEN 4. F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 2 IP_CELL_ROW = 4 ).
      ENDCASE.

      LT_MERGE = F_CUT->GET_MERGE( ).
      LV_SIZE = LINES( LT_MERGE ).

      CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
        ACT   = LV_SIZE
        EXP   = 1
        MSG   = 'Expect no merge were deleted'
        LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
      ).
    ENDDO.

*  Test 5. Delete concrete merge #1

    F_CUT->DELETE_MERGE( ).

    F_CUT->SET_MERGE(
      IP_COLUMN_START = 2
      IP_COLUMN_END = 3
      IP_ROW = 2
      IP_ROW_TO = 3
    ).
    F_CUT->SET_MERGE(
      IP_COLUMN_START = 4
      IP_COLUMN_END = 5
      IP_ROW = 4
      IP_ROW_TO = 5
    ).

    F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 2 IP_CELL_ROW = 2 ).
    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE
      EXP   = 1
      MSG   = 'Expect we have the one merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

    READ TABLE LT_MERGE INTO LV_MERGE INDEX 1.

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_MERGE
      EXP   = 'D4:E5'
      MSG   = 'Expect delete B2:C3 merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

*  Test 6. Delete concrete merge #2

    F_CUT->DELETE_MERGE( ).

    F_CUT->SET_MERGE(
      IP_COLUMN_START = 2
      IP_COLUMN_END = 3
      IP_ROW = 2
      IP_ROW_TO = 3
    ).
    F_CUT->SET_MERGE(
      IP_COLUMN_START = 4
      IP_COLUMN_END = 5
      IP_ROW = 4
      IP_ROW_TO = 5
    ).

    F_CUT->DELETE_MERGE( IP_CELL_COLUMN = 4 IP_CELL_ROW = 4 ).
    LT_MERGE = F_CUT->GET_MERGE( ).
    LV_SIZE = LINES( LT_MERGE ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_SIZE
      EXP   = 1
      MSG   = 'Expect we have the one merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

    READ TABLE LT_MERGE INTO LV_MERGE INDEX 1.

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = LV_MERGE
      EXP   = 'B2:C3'
      MSG   = 'Expect delete D4:E5 merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

  ENDMETHOD.       "delete_Merge

  METHOD GET_DIMENSION_RANGE.
    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = F_CUT->GET_DIMENSION_RANGE( )
      EXP   = 'A1'
      MSG   = 'get_dimension_range inital value'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

    F_CUT->SET_CELL(
      IP_ROW       = 2
      IP_COLUMN    = 3
      IP_VALUE     = 'Dummy'
    ).

    F_CUT->SET_CELL(
      IP_ROW       = 5
      IP_COLUMN    = 6
      IP_VALUE     = 'Dummy'
    ).

    CL_ABAP_UNIT_ASSERT=>ASSERT_EQUALS(
      ACT   = F_CUT->GET_DIMENSION_RANGE( )
      EXP   = 'C2:F5'
      MSG   = 'get_dimension_range'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
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
