
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

      ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

      ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
      ACT   = LV_SIZE
      EXP   = 1
      MSG   = 'Expect we have the one merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

    READ TABLE LT_MERGE INTO LV_MERGE INDEX 1.

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
      ACT   = LV_SIZE
      EXP   = 1
      MSG   = 'Expect we have the one merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

    READ TABLE LT_MERGE INTO LV_MERGE INDEX 1.

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
      ACT   = LV_MERGE
      EXP   = 'B2:C3'
      MSG   = 'Expect delete D4:E5 merge'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).

  ENDMETHOD.       "delete_Merge

  METHOD GET_DIMENSION_RANGE.
    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
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

    ZCL_EXCEL_AUNIT=>ASSERT_EQUALS(
      ACT   = F_CUT->GET_DIMENSION_RANGE( )
      EXP   = 'C2:F5'
      MSG   = 'get_dimension_range'
      LEVEL = IF_AUNIT_CONSTANTS=>CRITICAL
    ).
  ENDMETHOD.

ENDCLASS.       "lcl_Excel_Worksheet_Test
