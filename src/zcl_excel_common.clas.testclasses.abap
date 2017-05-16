CLASS lcl_excel_common_test DEFINITION DEFERRED.
CLASS zcl_excel_common DEFINITION LOCAL FRIENDS lcl_excel_common_test.

*----------------------------------------------------------------------*
*       CLASS lcl_Excel_Common_Test DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_excel_common_test DEFINITION FOR TESTING  "#AU Risk_Level Harmless
                                 .                  "#AU Duration Short
*?ï»¿<asx:abap xmlns:asx="http://www.sap.com/abapxml" version="1.0">
*?<asx:values>
*?<TESTCLASS_OPTIONS>
*?<TEST_CLASS>lcl_Excel_Common_Test
*?</TEST_CLASS>
*?<TEST_MEMBER>f_Cut
*?</TEST_MEMBER>
*?<OBJECT_UNDER_TEST>ZCL_EXCEL_COMMON
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
      lx_excel      TYPE REF TO zcx_excel,
      ls_symsg_act  TYPE symsg,                    " actual   messageinformation of exception
      ls_symsg_exp  TYPE symsg,                    " expected messageinformation of exception
      f_cut         TYPE REF TO zcl_excel_common.  "class under test

    CLASS-METHODS: class_setup.
    CLASS-METHODS: class_teardown.
    METHODS: setup.
    METHODS: teardown.
*    METHODS: char2hex FOR TESTING.
    METHODS: convert_column2alpha FOR TESTING.
    METHODS: convert_column2int FOR TESTING.
    METHODS: date_to_excel_string FOR TESTING.
    METHODS: encrypt_password FOR TESTING.
    METHODS: excel_string_to_date FOR TESTING.
    METHODS: excel_string_to_time FOR TESTING.
*    METHODS: number_to_excel_string FOR TESTING.
    METHODS: time_to_excel_string FOR TESTING.
    METHODS: split_file FOR TESTING.
    METHODS: convert_range2column_a_row FOR TESTING.
    METHODS: describe_structure FOR TESTING.
    METHODS: calculate_cell_distance FOR TESTING.
    METHODS: shift_formula FOR TESTING.
    METHODS: is_cell_in_range FOR TESTING.
ENDCLASS.       "lcl_Excel_Common_Test


*----------------------------------------------------------------------*
*       CLASS lcl_Excel_Common_Test IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_excel_common_test IMPLEMENTATION.
* ===========================================

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


  METHOD convert_column2alpha.
* ============================
    DATA ep_column TYPE zexcel_cell_column_alpha.

* Test 1. Simple test
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 1 ).

        zcl_excel_common=>assert_equals(
          act   = ep_column
          exp   = 'A'
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Max column for OXML #16,384 = XFD
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 16384 ).

        zcl_excel_common=>assert_equals(
          act   = ep_column
          exp   = 'XFD'
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 3. Index 0 is out of bounds
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 0 ).

        zcl_excel_common=>assert_equals(
          act   = ep_column
          exp   = 'A'
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>assert_equals(
          act   = lx_excel->error
          exp   = 'Index out of bounds'
          msg   = 'Colum index 0 is out of bounds, min column index is 1'
          level = if_aunit_constants=>fatal
        ).
    ENDTRY.

* Test 4. Exception should be thrown index out of bounds
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 16385 ).

        zcl_excel_common=>assert_differs(
          act   = ep_column
          exp   = 'XFE'
          msg   = 'Colum index 16385 is out of bounds, max column index is 16384'
          level = if_aunit_constants=>fatal
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>assert_equals(
          act   = lx_excel->error
          exp   = 'Index out of bounds'
          msg   = 'Wrong exception is thrown'
          level = if_aunit_constants=>tolerable
        ).
    ENDTRY.
  ENDMETHOD.       "convert_Column2alpha


  METHOD convert_column2int.
* ==========================
    DATA ep_column TYPE zexcel_cell_column.

* Test 1. Basic test
    TRY.
        ep_column = zcl_excel_common=>convert_column2int( 'A' ).

        zcl_excel_common=>assert_equals(
          act   = ep_column
          exp   = 1
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Max column
    TRY.
        ep_column = zcl_excel_common=>convert_column2int( 'XFD' ).

        zcl_excel_common=>assert_equals(
          act   = ep_column
          exp   = 16384
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 3. Out of bounds
    TRY.
        ep_column = zcl_excel_common=>convert_column2int( '' ).

        zcl_excel_common=>assert_differs( act   = ep_column
                                          exp   = '0'
                                          msg   = 'Wrong column conversion'
                                          level = if_aunit_constants=>critical ).
      CATCH zcx_excel INTO lx_excel.
        CLEAR: ls_symsg_act,
               ls_symsg_exp.
        ls_symsg_exp-msgid = 'ZABAP2XLSX'.
        ls_symsg_exp-msgno = '800'.
        ls_symsg_act-msgid = lx_excel->syst_at_raise-msgid.
        ls_symsg_act-msgno = lx_excel->syst_at_raise-msgno.
        zcl_excel_common=>assert_equals( act   = ls_symsg_act
                                         exp   = ls_symsg_exp
                                         msg   = 'Colum name should be a valid string'
                                         level = if_aunit_constants=>fatal ).
    ENDTRY.

* Test 4. Out of bounds
    TRY.
        ep_column = zcl_excel_common=>convert_column2int( 'XFE' ).

        zcl_excel_common=>assert_differs( act   = ep_column
                                          exp   = 16385
                                          msg   = 'Wrong column conversion'
                                          level = if_aunit_constants=>critical ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>assert_equals( act   = lx_excel->error
                                         exp   = 'Index out of bounds'
                                         msg   = 'Colum XFE is out of range'
                                         level = if_aunit_constants=>fatal ).
    ENDTRY.
  ENDMETHOD.       "convert_Column2int


  METHOD date_to_excel_string.
* ============================
    DATA ep_value TYPE zexcel_cell_value.

* Test 1. Basic conversion
    TRY.
        ep_value = zcl_excel_common=>date_to_excel_string( '19000101' ).

        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = 1
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
* Check around the "Excel Leap Year" 1900
    TRY.
        ep_value = zcl_excel_common=>date_to_excel_string( '19000228' ).

        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = 59
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
    TRY.
        ep_value = zcl_excel_common=>date_to_excel_string( '19000301' ).

        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = 61
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.


* Test 2. Basic conversion
    TRY.
        ep_value = zcl_excel_common=>date_to_excel_string( '99991212' ).

        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = 2958446
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 3. Initial date
    TRY.
        DATA: lv_date TYPE d.
        ep_value = zcl_excel_common=>date_to_excel_string( lv_date ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = ''
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Basic conversion
    TRY.
        DATA exp_value TYPE zexcel_cell_value VALUE 0.
        ep_value = zcl_excel_common=>date_to_excel_string( '18991231' ).

        zcl_excel_common=>assert_differs(
              act   = ep_value
              exp   = exp_value
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>assert_equals(
          act   = lx_excel->error
          exp   = 'Index out of bounds'
          msg   = 'Dates prior of 1900 are not available in excel'
          level = if_aunit_constants=>critical
        ).
    ENDTRY.

  ENDMETHOD.       "date_To_Excel_String


  METHOD encrypt_password.
* ========================
    DATA lv_encrypted_pwd TYPE zexcel_aes_password.

    TRY.
        lv_encrypted_pwd = zcl_excel_common=>encrypt_password( 'test' ).

        zcl_excel_common=>assert_equals(
              act   = lv_encrypted_pwd
              exp   = 'CBEB'
              msg   = 'Wrong password encryption'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
  ENDMETHOD.       "encrypt_Password


  METHOD excel_string_to_date.
* ============================
    DATA ep_value TYPE d.


* Test 1. Simple test -> ABAP Manage also date prior of 1900
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '0' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '18991231'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>tolerable
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
* Check around the "Excel Leap Year" 1900
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '59' ).

        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = '19000228'
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '61' ).

        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = '19000301'
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Simple test
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '1' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '19000101'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 3. Index 0 is out of bounds
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '2958446' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '99991212'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 4. Exception should be thrown index out of bounds
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '2958447' ).

        zcl_excel_common=>assert_differs(
          act   = ep_value
          exp   = '99991212'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>fatal
        ).

        zcl_excel_common=>assert_differs(
          act   = ep_value
          exp   = '00000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>fatal
        ).

      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>assert_equals(
          act   = lx_excel->error
          exp   = 'Index out of bounds'
          msg   = 'Wrong exception is thrown'
          level = if_aunit_constants=>tolerable
        ).
    ENDTRY.
  ENDMETHOD.       "excel_String_To_Date


  METHOD excel_string_to_time.
* ============================
    DATA ep_value TYPE t.

* Test 1. Simple test
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_time( '0' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>tolerable
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Simple test
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_time( '1' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 3. Simple test
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_time( '0.99999' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '235959'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 4. Also string greater than 1 should be managed
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_time( '4.1' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '022400'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 4. string is not a number
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_time( 'NaN' ).

        zcl_excel_common=>assert_differs(
          act   = ep_value
          exp   = '000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>assert_equals(
          act   = lx_excel->error
          exp   = 'Unable to interpret time'
          msg   = 'Time should be a valid string'
          level = if_aunit_constants=>fatal
        ).
    ENDTRY.
  ENDMETHOD.       "excel_String_To_Time


  METHOD time_to_excel_string.
* ============================
    DATA ep_value TYPE zexcel_cell_value.

* Test 1. Basic conversion
    TRY.
        ep_value = zcl_excel_common=>time_to_excel_string( '000001' ).
        " A test directly in Excel returns the value 0.0000115740740740741000
        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = '0.0000115740740741'
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Basic conversion
    TRY.
        ep_value = zcl_excel_common=>time_to_excel_string( '235959' ).
        " A test directly in Excel returns the value 0.9999884259259260000000
        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = '0.9999884259259260'
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 3. Initial date
    TRY.
        ep_value = zcl_excel_common=>time_to_excel_string( '000000' ).

        zcl_excel_common=>assert_equals(
          act   = ep_value
          exp   = '0'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Basic conversion
    TRY.
        ep_value = zcl_excel_common=>time_to_excel_string( '022400' ).

        zcl_excel_common=>assert_equals(
              act   = ep_value
              exp   = '0.1000000000000000'
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        zcl_excel_common=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

  ENDMETHOD.       "time_To_Excel_String

  METHOD split_file.
* ============================

    DATA: ep_file	TYPE text255,
    ep_extension  TYPE char10,
    ep_dotextension	TYPE char10.


* Test 1. Basic conversion
    zcl_excel_common=>split_file( EXPORTING ip_file         = 'filename.xml'
                                  IMPORTING ep_file         = ep_file
                                            ep_extension    = ep_extension
                                            ep_dotextension = ep_dotextension ).

    zcl_excel_common=>assert_equals(
          act   = ep_file
          exp   = 'filename'
          msg   = 'Split filename failed'
          level = if_aunit_constants=>critical ).

    zcl_excel_common=>assert_equals(
          act   = ep_extension
          exp   = 'xml'
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

    zcl_excel_common=>assert_equals(
          act   = ep_dotextension
          exp   = '.xml'
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

* Test 2. no extension
    zcl_excel_common=>split_file( EXPORTING ip_file         = 'filename'
                                  IMPORTING ep_file         = ep_file
                                            ep_extension    = ep_extension
                                            ep_dotextension = ep_dotextension ).

    zcl_excel_common=>assert_equals(
          act   = ep_file
          exp   = 'filename'
          msg   = 'Split filename failed'
          level = if_aunit_constants=>critical ).

    zcl_excel_common=>assert_equals(
          act   = ep_extension
          exp   = ''
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

    zcl_excel_common=>assert_equals(
          act   = ep_dotextension
          exp   = ''
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.  "split_file

  METHOD convert_range2column_a_row.
    DATA: lv_range        TYPE string.
    DATA: lv_column_start TYPE zexcel_cell_column_alpha,
          lv_column_end   TYPE zexcel_cell_column_alpha,
          lv_row_start    TYPE zexcel_cell_row,
          lv_row_end      TYPE zexcel_cell_row,
          lv_sheet        TYPE string.

* a) input empty --> nothing to do
    zcl_excel_common=>convert_range2column_a_row(
      EXPORTING
        i_range        = lv_range
      IMPORTING
        e_column_start = lv_column_start    " Cell Column Start
        e_column_end   = lv_column_end    " Cell Column End
        e_row_start    = lv_row_start    " Cell Row
        e_row_end      = lv_row_end    " Cell Row
        e_sheet        = lv_sheet    " Title
    ).

    zcl_excel_common=>assert_equals(
          act   = lv_column_start
          exp   = ''
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_column_end
          exp   = ''
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_start
          exp   = ''
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_end
          exp   = ''
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_sheet
          exp   = ''
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
* b) sheetname existing - starts with '            example 'Sheet 1'!$B$6:$D$13
    lv_range = `'Sheet 1'!$B$6:$D$13`.
    zcl_excel_common=>convert_range2column_a_row(
      EXPORTING
        i_range        = lv_range
      IMPORTING
        e_column_start = lv_column_start    " Cell Column Start
        e_column_end   = lv_column_end    " Cell Column End
        e_row_start    = lv_row_start    " Cell Row
        e_row_end      = lv_row_end    " Cell Row
        e_sheet        = lv_sheet    " Title
    ).

    zcl_excel_common=>assert_equals(
          act   = lv_column_start
          exp   = 'B'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_column_end
          exp   = 'D'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_start
          exp   = '6'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_end
          exp   = '13'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_sheet
          exp   = 'Sheet 1'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
* c) sheetname existing - does not start with '    example Sheet1!$B$6:$D$13
    lv_range = `Sheet1!B6:$D$13`.
    zcl_excel_common=>convert_range2column_a_row(
      EXPORTING
        i_range        = lv_range
      IMPORTING
        e_column_start = lv_column_start    " Cell Column Start
        e_column_end   = lv_column_end    " Cell Column End
        e_row_start    = lv_row_start    " Cell Row
        e_row_end      = lv_row_end    " Cell Row
        e_sheet        = lv_sheet    " Title
    ).

    zcl_excel_common=>assert_equals(
          act   = lv_column_start
          exp   = 'B'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_column_end
          exp   = 'D'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_start
          exp   = '6'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_end
          exp   = '13'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_sheet
          exp   = 'Sheet1'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
* d) no sheetname - just area                      example $B$6:$D$13
    lv_range = `$B$6:D13`.
    zcl_excel_common=>convert_range2column_a_row(
      EXPORTING
        i_range        = lv_range
      IMPORTING
        e_column_start = lv_column_start    " Cell Column Start
        e_column_end   = lv_column_end    " Cell Column End
        e_row_start    = lv_row_start    " Cell Row
        e_row_end      = lv_row_end    " Cell Row
        e_sheet        = lv_sheet    " Title
    ).

    zcl_excel_common=>assert_equals(
          act   = lv_column_start
          exp   = 'B'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_column_end
          exp   = 'D'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_start
          exp   = '6'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_row_end
          exp   = '13'
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
    zcl_excel_common=>assert_equals(
          act   = lv_sheet
          exp   = ''
          msg   = 'Conversion of range failed'
          level = if_aunit_constants=>critical ).
  ENDMETHOD.                    "convert_range2column_a_row


  METHOD describe_structure.
    DATA: ls_test TYPE scarr.
    DATA: lo_structdescr TYPE REF TO cl_abap_structdescr.
    DATA: lt_structure TYPE ddfields.
    FIELD-SYMBOLS: <line> LIKE LINE OF lt_structure.

    " Test with DDIC Type
    lo_structdescr ?= cl_abap_structdescr=>describe_by_data( p_data = ls_test ).
    lt_structure = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).
    READ TABLE lt_structure ASSIGNING <line> INDEX 1.
    zcl_excel_common=>assert_equals(
          act   = <line>-fieldname
          exp   = 'MANDT'
          msg   = 'Describe structure failed'
          level = if_aunit_constants=>critical ).

    " Test with local defined structure having DDIC and non DDIC elements
    TYPES:
      BEGIN OF t_test,
        carrid   TYPE s_carr_id,
        carrname TYPE s_carrname,
        carrdesc TYPE string,
      END OF t_test.
    DATA: ls_ttest TYPE t_test.

    lo_structdescr ?= cl_abap_structdescr=>describe_by_data( p_data = ls_ttest ).
    lt_structure = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).
    READ TABLE lt_structure ASSIGNING <line> INDEX 1.
    zcl_excel_common=>assert_equals(
          act   = <line>-fieldname
          exp   = 'CARRID'
          msg   = 'Describe structure failed'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.                    "describe_structure


  METHOD calculate_cell_distance.
    DATA: lv_offset_rows             TYPE i,
          lv_offset_cols             TYPE i,
          lv_message                 TYPE string.

    DEFINE macro_calculate_cell_distance.
      zcl_excel_common=>calculate_cell_distance( exporting iv_reference_cell = &1
                                                           iv_current_cell   = &2
                                                 importing ev_row_difference = lv_offset_rows
                                                           ev_col_difference = lv_offset_cols ).
* Check delta columns
      concatenate 'Error calculating column difference in test:'
                  &1
                  '->'
                  &2
           into lv_message separated by space.
      zcl_excel_common=>assert_equals(  act   = lv_offset_cols
                                        exp   = &3
                                        msg   = lv_message
                                        quit  = 0  " continue tests
                                        level = if_aunit_constants=>critical ).
* Check delta rows
      concatenate 'Error calculating row difference in test:'
                  &1
                  '->'
                  &2
           into lv_message separated by space.
      zcl_excel_common=>assert_equals(  act   = lv_offset_rows
                                        exp   = &4
                                        msg   = lv_message
                                        quit  = 0  " continue tests
                                        level = if_aunit_constants=>critical ).
    END-OF-DEFINITION.


    macro_calculate_cell_distance:
          'C12'        'C12'          0           0        ,  " Same cell
          'C12'        'C13'          0           1        ,  " Shift down 1 place
          'C12'        'C25'          0          13        ,  " Shift down some places
          'C12'        'C11'          0          -1        ,  " Shift up 1 place
          'C12'        'C1'           0         -11        ,  " Shift up some place
          'C12'        'D12'          1           0        ,  " Shift right 1 place
          'C12'        'AA12'        24           0        ,  " Shift right some places
          'C12'        'B12'         -1           0        ,  " Shift left 1 place
          'AA12'       'C12'        -24           0        ,  " Shift left some place
          'AA121'      'C12'        -24        -109        .  " The full package.

  ENDMETHOD.                    "CALCULATE_CELL_DISTANCE

  METHOD shift_formula.
    DATA: lv_resulting_formula       TYPE string,
          lv_message                 TYPE string,
          lv_counter                 TYPE num8.

    DEFINE macro_shift_formula.
      add 1 to lv_counter.
      clear lv_resulting_formula.
      try.
          lv_resulting_formula = zcl_excel_common=>shift_formula( iv_reference_formula = &1
                                                                  iv_shift_cols        = &2
                                                                  iv_shift_rows        = &3 ).
          concatenate 'Wrong result in test'
                      lv_counter
                      'shifting formula '
                      &1
               into lv_message separated by space.
          zcl_excel_common=>assert_equals(  act   = lv_resulting_formula
                                            exp   = &4
                                            msg   = lv_message
                                            quit  = 0  " continue tests
                                            level = if_aunit_constants=>critical ).
        catch zcx_excel.
          concatenate 'Unexpected exception occurred in test'
                      lv_counter
                      'shifting formula '
                      &1
               into lv_message separated by space.
          zcl_excel_common=>assert_equals(  act   = lv_resulting_formula
                                            exp   = &4
                                            msg   = lv_message
                                            quit  = 0  " continue tests
                                            level = if_aunit_constants=>critical ).
      endtry.
    END-OF-DEFINITION.

* Test shifts that should result in a valid output
    macro_shift_formula:
          'C17'                                  0   0       'C17',                       " Very basic check
          'C17'                                  2   3       'E20',                       " Check shift right and down
          'C17'                                 -2  -3       'A14',                       " Check shift left and up
          '$C$17'                                1   1       '$C$17',                     " Fixed columns/rows
          'SUM($C17:C$23)+C30'                   1  11       'SUM($C28:D$23)+D41',        " Operators and Ranges, mixed fixed rows or columns
          'RNGNAME1+C7'                         -1  -4       'RNGNAME1+B3',               " Operators and Rangename
          '"Date:"&TEXT(B2)'                     1   1       '"Date:"&TEXT(C3)',          " String literals and string concatenation
          '[TEST6.XLSX]SHEET1!A1'                1  11       '[TEST6.XLSX]SHEET1!B12',    " External sheet reference
          `X(B13, "KK" )  `                      1   1       `X(C14,"KK")`,               " superflous blanks, multi-argument functions, literals in function, unknown functions
*          'SIN((((((B2))))))'                    1   1       'SIN((((((C3))))))',        " Deep nesting
*          'SIN(SIN(SIN(SIN(E22))))'              0   1       'SIN(SIN(SIN(SIN(E23))))',   " Different type of deep nesting
          `SIN(SIN(SIN(SIN(E22))))`              0   1       'SIN(SIN(SIN(SIN(E23))))',   " same as above - but with string input instead of Char-input
          'HEUTE()'                              2   5       'HEUTE()',                   " Functions w/o arguments, No cellreferences
          '"B2"'                                 2   5       '"B2"',                      " No cellreferences
          ''                                     2   5       '',                          " Empty
          'A1+$A1+A$1+$A$1+B2'                  -1   0       '#REF!+$A1+#REF!+$A$1+A2',   " Referencing error , column only    , underflow
          'A1+$A1+A$1+$A$1+B2'                   0  -1       '#REF!+#REF!+A$1+$A$1+B1',   " Referencing error , row only       , underflow
          'A1+$A1+A$1+$A$1+B2'                  -1  -1       '#REF!+#REF!+#REF!+$A$1+A1'. " Referencing error , row and column , underflow
  ENDMETHOD.                    "SHIFT_FORMULA

  METHOD is_cell_in_range.
    DATA ep_cell_in_range TYPE abap_bool.

* Test 1: upper left corner (in range)
    TRY.
      ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
          ip_column   = 'B'
          ip_row      = 2
          ip_range    = 'B2:D4' ).

      zcl_excel_common=>assert_equals(
          act   = ep_cell_in_range
          exp   = abap_true
          msg   = 'Check cell in range failed'
          level = if_aunit_constants=>critical ).
     CATCH zcx_excel.
        zcl_excel_common=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.

* Test 2: lower right corner (in range)
    TRY.
      ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
          ip_column   = 'D'
          ip_row      = 4
          ip_range    = 'B2:D4' ).

      zcl_excel_common=>assert_equals(
          act   = ep_cell_in_range
          exp   = abap_true
          msg   = 'Check cell in range failed'
          level = if_aunit_constants=>critical ).
     CATCH zcx_excel.
        zcl_excel_common=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.

* Test 3: left side (out of range)
    TRY.
      ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
          ip_column   = 'A'
          ip_row      = 3
          ip_range    = 'B2:D4' ).

      zcl_excel_common=>assert_equals(
          act   = ep_cell_in_range
          exp   = abap_false
          msg   = 'Check cell in range failed'
          level = if_aunit_constants=>critical ).
     CATCH zcx_excel.
        zcl_excel_common=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.

* Test 4: upper side (out of range)
    TRY.
      ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
          ip_column   = 'C'
          ip_row      = 1
          ip_range    = 'B2:D4' ).

      zcl_excel_common=>assert_equals(
          act   = ep_cell_in_range
          exp   = abap_false
          msg   = 'Check cell in range failed'
          level = if_aunit_constants=>critical ).
     CATCH zcx_excel.
        zcl_excel_common=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.

* Test 5: right side (out of range)
    TRY.
      ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
          ip_column   = 'E'
          ip_row      = 3
          ip_range    = 'B2:D4' ).

      zcl_excel_common=>assert_equals(
          act   = ep_cell_in_range
          exp   = abap_false
          msg   = 'Check cell in range failed'
          level = if_aunit_constants=>critical ).
     CATCH zcx_excel.
        zcl_excel_common=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.

* Test 6: lower side (out of range)
    TRY.
      ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
          ip_column   = 'C'
          ip_row      = 5
          ip_range    = 'B2:D4' ).

      zcl_excel_common=>assert_equals(
          act   = ep_cell_in_range
          exp   = abap_false
          msg   = 'Check cell in range failed'
          level = if_aunit_constants=>critical ).
     CATCH zcx_excel.
        zcl_excel_common=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.
  ENDMETHOD.

ENDCLASS.       "lcl_Excel_Common_Test
