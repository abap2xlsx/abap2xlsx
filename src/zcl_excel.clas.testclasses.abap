CLASS zcl_tc_excel DEFINITION DEFERRED.
CLASS zcl_excel DEFINITION LOCAL FRIENDS zcl_tc_excel.

*----------------------------------------------------------------------*
*       CLASS zcl_Tc_Excel DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_tc_excel DEFINITION FOR TESTING
  DURATION SHORT
  RISK LEVEL HARMLESS
.
*?ï»¿<asx:abap xmlns:asx="http://www.sap.com/abapxml" version="1.0">
*?<asx:values>
*?<TESTCLASS_OPTIONS>
*?<TEST_CLASS>zcl_Tc_Excel
*?</TEST_CLASS>
*?<TEST_MEMBER>f_Cut
*?</TEST_MEMBER>
*?<OBJECT_UNDER_TEST>ZCL_EXCEL
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
      f_cut TYPE REF TO zcl_excel.  "class under test

    CLASS-METHODS: class_setup.
    CLASS-METHODS: class_teardown.
    METHODS: setup.
    METHODS: teardown.
    METHODS: create_empty_excel FOR TESTING.

ENDCLASS.       "zcl_Tc_Excel


*----------------------------------------------------------------------*
*       CLASS zcl_Tc_Excel IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_tc_excel IMPLEMENTATION.
* ==================================

  METHOD class_setup.
* ===================


  ENDMETHOD.       "class_Setup


  METHOD class_teardown.
* ======================


  ENDMETHOD.       "class_Teardown


  METHOD setup.
* =============

    CREATE OBJECT f_cut.
  ENDMETHOD.       "setup


  METHOD teardown.
* ================


  ENDMETHOD.       "teardown

*// START TEST METHODS

  METHOD create_empty_excel.
* ==================================

    DATA: lv_count TYPE i.
    lv_count = f_cut->get_worksheets_size( ).

    cl_abap_unit_assert=>assert_equals( act   = lv_count
                                        exp   = 1
                                        msg   = 'Testing number of sheet'
                                        level = if_aunit_constants=>tolerable ).
  ENDMETHOD.                    "create_empty_excel

*// END TEST METHODS


ENDCLASS.       "zcl_Tc_Excel
