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
    METHODS: get_worksheet_by_name_case_ins FOR TESTING RAISING cx_static_check.
    METHODS: get_worksheet_by_name_case_sen FOR TESTING RAISING cx_static_check.
    METHODS: get_worksheet_by_name_notfound FOR TESTING RAISING cx_static_check.

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

  METHOD get_worksheet_by_name_case_ins.
    DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet.

    f_cut->add_new_worksheet( ip_title = 'SheetOne' ).

    lo_worksheet = f_cut->get_worksheet_by_name( 'sheetone' ).
    cl_abap_unit_assert=>assert_bound(
      act = lo_worksheet
      msg = 'Should find worksheet case-insensitively with lower case' ).
    cl_abap_unit_assert=>assert_equals(
      act = lo_worksheet->get_title( )
      exp = 'SheetOne' ).

    lo_worksheet = f_cut->get_worksheet_by_name( 'SHEETONE' ).
    cl_abap_unit_assert=>assert_bound(
      act = lo_worksheet
      msg = 'Should find worksheet case-insensitively with upper case' ).

    lo_worksheet = f_cut->get_worksheet_by_name( 'SheetOne' ).
    cl_abap_unit_assert=>assert_bound(
      act = lo_worksheet
      msg = 'Should find worksheet with exact case' ).
  ENDMETHOD.

  METHOD get_worksheet_by_name_case_sen.
    DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet.

    f_cut->add_new_worksheet( ip_title = 'SheetOne' ).

    lo_worksheet = f_cut->get_worksheet_by_name(
      ip_sheet_name = 'sheetone'
      iv_match_case = abap_true ).
    cl_abap_unit_assert=>assert_not_bound(
      act = lo_worksheet
      msg = 'Should not find worksheet when match_case is true and case differs' ).

    lo_worksheet = f_cut->get_worksheet_by_name(
      ip_sheet_name = 'SheetOne'
      iv_match_case = abap_true ).
    cl_abap_unit_assert=>assert_bound(
      act = lo_worksheet
      msg = 'Should find worksheet when match_case is true and case matches' ).
  ENDMETHOD.

  METHOD get_worksheet_by_name_notfound.
    DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet.

    f_cut->add_new_worksheet( ip_title = 'SheetOne' ).

    lo_worksheet = f_cut->get_worksheet_by_name( 'NonExistent' ).
    cl_abap_unit_assert=>assert_not_bound(
      act = lo_worksheet
      msg = 'Should return unbound for non-existent sheet' ).
  ENDMETHOD.

*// END TEST METHODS


ENDCLASS.       "zcl_Tc_Excel
