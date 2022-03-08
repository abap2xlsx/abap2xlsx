CLASS lcl_excel_common_test DEFINITION DEFERRED.
CLASS zcl_excel_common DEFINITION LOCAL FRIENDS lcl_excel_common_test.

*----------------------------------------------------------------------*
*       CLASS lcl_Excel_Common_Test DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_excel_common_test DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.
* ================
    TYPES: BEGIN OF ty_convert_range2column_a_row,
             column_start     TYPE zexcel_cell_column_alpha,
             column_start_int TYPE zexcel_cell_column,
             column_end       TYPE zexcel_cell_column_alpha,
             column_end_int   TYPE zexcel_cell_column,
             row_start        TYPE zexcel_cell_row,
             row_end          TYPE zexcel_cell_row,
             sheet            TYPE string,
           END OF ty_convert_range2column_a_row.
    DATA:
      lx_excel     TYPE REF TO zcx_excel,
      ls_symsg_act LIKE sy,                    " actual   messageinformation of exception
      ls_symsg_exp LIKE sy,                    " expected messageinformation of exception
      f_cut        TYPE REF TO zcl_excel_common.  "class under test

    METHODS: setup.
    METHODS: convert_column2alpha_simple FOR TESTING.
    METHODS: convert_column2alpha_maxcol FOR TESTING.
    METHODS: convert_column2alpha_last FOR TESTING.
    METHODS: convert_column2alpha_oob FOR TESTING.
    METHODS convert_column2int_basic FOR TESTING.
    METHODS convert_column2int_from_int FOR TESTING RAISING cx_static_check.
    METHODS convert_column2int_maxcol FOR TESTING.
    METHODS convert_column2int_oob_empty FOR TESTING.
    METHODS convert_column2int_oob_invalid FOR TESTING.
    METHODS convert_column_a_row2columnrow FOR TESTING RAISING cx_static_check.
    METHODS convert_columnrow2column_a_row FOR TESTING RAISING cx_static_check.
    METHODS date_to_excel_string1 FOR TESTING RAISING cx_static_check.
    METHODS date_to_excel_string2 FOR TESTING RAISING cx_static_check.
    METHODS date_to_excel_string3 FOR TESTING RAISING cx_static_check.
    METHODS date_to_excel_string4 FOR TESTING RAISING cx_static_check.
    METHODS date_to_excel_string5 FOR TESTING RAISING cx_static_check.
    METHODS date_to_excel_string6 FOR TESTING RAISING cx_static_check.
    METHODS: encrypt_password FOR TESTING.
    METHODS: excel_string_to_date FOR TESTING.
    METHODS excel_string_to_time1 FOR TESTING RAISING cx_static_check.
    METHODS excel_string_to_time2 FOR TESTING RAISING cx_static_check.
    METHODS excel_string_to_time3 FOR TESTING RAISING cx_static_check.
    METHODS excel_string_to_time4 FOR TESTING RAISING cx_static_check.
    METHODS excel_string_to_time5 FOR TESTING RAISING cx_static_check.
    METHODS time_to_excel_string1 FOR TESTING RAISING cx_static_check.
    METHODS time_to_excel_string2 FOR TESTING RAISING cx_static_check.
    METHODS time_to_excel_string3 FOR TESTING RAISING cx_static_check.
    METHODS time_to_excel_string4 FOR TESTING RAISING cx_static_check.
    METHODS: split_file FOR TESTING.
    METHODS: convert_range2column_a_row FOR TESTING RAISING cx_static_check.
    METHODS: assert_convert_range2column_a_
      IMPORTING
        i_range            TYPE clike
        i_allow_1dim_range TYPE abap_bool DEFAULT abap_false
        is_exp             TYPE ty_convert_range2column_a_row
      RAISING
        cx_static_check.
    METHODS: describe_structure FOR TESTING.
    METHODS macro_calculate_cell_distance
      IMPORTING
        iv_reference_cell  TYPE clike
        iv_current_cell    TYPE clike
        iv_expected_column TYPE i
        iv_expected_row    TYPE i
      RAISING
        cx_static_check.
    METHODS: calc_cell_dist_samecell FOR TESTING RAISING cx_static_check,
      calc_cell_dist_down1pl FOR TESTING RAISING cx_static_check,
      calc_cell_dist_downsome FOR TESTING RAISING cx_static_check,
      calc_cell_dist_up1pl FOR TESTING RAISING cx_static_check,
      calc_cell_dist_upsome FOR TESTING RAISING cx_static_check,
      calc_cell_dist_right1pl FOR TESTING RAISING cx_static_check,
      calc_cell_dist_rightsome FOR TESTING RAISING cx_static_check,
      calc_cell_dist_left1pl FOR TESTING RAISING cx_static_check,
      calc_cell_dist_leftsome FOR TESTING RAISING cx_static_check,
      calc_cell_dist_fullpack FOR TESTING RAISING cx_static_check.
    METHODS macro_shift_formula
      IMPORTING
        iv_reference_formula TYPE clike
        iv_shift_cols        TYPE i
        iv_shift_rows        TYPE i
        iv_expected          TYPE string.
    METHODS: shift_formula_basic FOR TESTING,
      shift_formula_rightdown FOR TESTING,
      shift_formula_leftup FOR TESTING,
      shift_formula_fixedcolrows FOR TESTING,
      shift_formula_mixedfixedrows FOR TESTING,
      shift_formula_rangename FOR TESTING,
      shift_formula_stringlitconc FOR TESTING,
      shift_formula_extref FOR TESTING,
      shift_formula_charblanks FOR TESTING,
      shift_formula_stringblanks FOR TESTING,
      shift_formula_funcnoargs FOR TESTING,
      shift_formula_nocellref FOR TESTING,
      shift_formula_empty FOR TESTING,
      shift_formula_referr_colunder FOR TESTING,
      shift_formula_referr_rowunder FOR TESTING,
      shift_formula_referr_rowcolund FOR TESTING,
      shift_formula_sheet_nodigit FOR TESTING,
      shift_formula_sheet_nodig FOR TESTING,
      shift_formula_sheet_special FOR TESTING,
      shift_formula_resp_blanks_1 FOR TESTING,
      shift_formula_resp_blanks_2 FOR TESTING,
      shift_formula_range FOR TESTING,
      shift_formula_notcols FOR TESTING,
      shift_formula_name FOR TESTING,
      shift_formula_refcolumn1 FOR TESTING,
      shift_formula_refcolumn2 FOR TESTING.
    METHODS is_cell_in_range_ulc_in FOR TESTING.
    METHODS is_cell_in_range_lrc_in FOR TESTING.
    METHODS is_cell_in_range_leftside_out FOR TESTING.
    METHODS is_cell_in_range_upperside_out FOR TESTING.
    METHODS is_cell_in_range_rightside_out FOR TESTING.
    METHODS is_cell_in_range_lowerside_out FOR TESTING.
ENDCLASS.


*----------------------------------------------------------------------*
*       CLASS lcl_Excel_Common_Test IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_excel_common_test IMPLEMENTATION.
* ===========================================

  METHOD setup.
* =============

    CREATE OBJECT f_cut.
  ENDMETHOD.       "setup


  METHOD convert_column2alpha_simple.
* ============================
    DATA ep_column TYPE zexcel_cell_column_alpha.

* Test 1. Simple test
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 1 ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_column
          exp   = 'A'
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
  ENDMETHOD. "convert_column2alpha_simple


  METHOD convert_column2alpha_maxcol.
* ============================
    DATA ep_column TYPE zexcel_cell_column_alpha.

* Test 2. Max column for OXML #16,384 = XFD
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 16384 ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_column
          exp   = 'XFD'
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
  ENDMETHOD. "convert_column2alpha_maxcol


  METHOD convert_column2alpha_last.
* ============================
    DATA ep_column TYPE zexcel_cell_column_alpha.

* Test 3. Index 0 is out of bounds
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 0 ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_column
          exp   = 'A'
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>assert_equals(
          act   = lx_excel->error
          exp   = 'Index out of bounds'
          msg   = 'Colum index 0 is out of bounds, min column index is 1'
          level = if_aunit_constants=>fatal
        ).
    ENDTRY.
  ENDMETHOD. "convert_column2alpha_last


  METHOD convert_column2alpha_oob.
* ============================
    DATA ep_column TYPE zexcel_cell_column_alpha.

* Test 4. Exception should be thrown index out of bounds
    TRY.
        ep_column = zcl_excel_common=>convert_column2alpha( 16385 ).

        cl_abap_unit_assert=>assert_differs(
          act   = ep_column
          exp   = 'XFE'
          msg   = 'Colum index 16385 is out of bounds, max column index is 16384'
          level = if_aunit_constants=>fatal
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>assert_equals(
          act   = lx_excel->error
          exp   = 'Index out of bounds'
          msg   = 'Wrong exception is thrown'
          level = if_aunit_constants=>tolerable
        ).
    ENDTRY.
  ENDMETHOD.       "convert_Column2alpha_oob


  METHOD convert_column2int_basic.
* ==========================
* Test 1. Basic test
    DATA ep_column TYPE zexcel_cell_column.

    TRY.
        ep_column = zcl_excel_common=>convert_column2int( 'A' ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_column
          exp   = 1
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
  ENDMETHOD. "convert_column2int_basic.


  METHOD convert_column2int_from_int.

    DATA ep_column TYPE zexcel_cell_column.

    ep_column = zcl_excel_common=>convert_column2int( 5 ).

    cl_abap_unit_assert=>assert_equals( act = ep_column exp = 5 ).

  ENDMETHOD.


  METHOD convert_column2int_maxcol.
* ==========================
* Test 2. Max column
    DATA ep_column TYPE zexcel_cell_column.

    TRY.
        ep_column = zcl_excel_common=>convert_column2int( 'XFD' ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_column
          exp   = 16384
          msg   = 'Wrong column conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
  ENDMETHOD. "convert_column2int_maxcol


  METHOD convert_column2int_oob_empty.
* ==========================
* Test 3. Out of bounds
    DATA ep_column TYPE zexcel_cell_column.

    TRY.
        ep_column = zcl_excel_common=>convert_column2int( '' ).

        cl_abap_unit_assert=>assert_differs( act   = ep_column
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
        cl_abap_unit_assert=>assert_equals( act   = ls_symsg_act
                                         exp   = ls_symsg_exp
                                         msg   = 'Colum name should be a valid string'
                                         level = if_aunit_constants=>fatal ).
    ENDTRY.
  ENDMETHOD. "convert_column2int_oob_empty.


  METHOD convert_column2int_oob_invalid.
* ==========================
* Test 4. Out of bounds
    DATA ep_column TYPE zexcel_cell_column.

    TRY.
        ep_column = zcl_excel_common=>convert_column2int( 'XFE' ).

        cl_abap_unit_assert=>assert_differs( act   = ep_column
                                          exp   = 16385
                                          msg   = 'Wrong column conversion'
                                          level = if_aunit_constants=>critical ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>assert_equals( act   = lx_excel->error
                                         exp   = 'Index out of bounds'
                                         msg   = 'Colum XFE is out of range'
                                         level = if_aunit_constants=>fatal ).
    ENDTRY.
  ENDMETHOD.       "convert_column2int_oob_invalid.


  METHOD convert_column_a_row2columnrow.

   DATA: cell_coords TYPE string.

   cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = 'B' i_row = 6 ).

   cl_abap_unit_assert=>assert_equals( act = cell_coords exp = 'B6' ).


   cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = 2 i_row = 6 ).

   cl_abap_unit_assert=>assert_equals( act = cell_coords exp = 'B6' ).

  ENDMETHOD.


  METHOD convert_columnrow2column_a_row.

   DATA: column     TYPE zexcel_cell_column_alpha,
         column_int TYPE zexcel_cell_column,
         row        TYPE zexcel_cell_row.

   zcl_excel_common=>convert_columnrow2column_a_row(
     EXPORTING
       i_columnrow  = 'B6'
     IMPORTING
       e_column     = column
       e_column_int = column_int
       e_row        = row ).

   cl_abap_unit_assert=>assert_equals( act = column     exp = 'B'   msg = 'Invalid column (alpha)' ).
   cl_abap_unit_assert=>assert_equals( act = column_int exp = 2     msg = 'Invalid column (numeric)' ).
   cl_abap_unit_assert=>assert_equals( act = row        exp = 6     msg = 'Invalid row' ).

  ENDMETHOD.


  METHOD date_to_excel_string1.
    DATA ep_value TYPE zexcel_cell_value.

* Test 1. Basic conversion
    ep_value = zcl_excel_common=>date_to_excel_string( '19000101' ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = 1
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD date_to_excel_string2.
    DATA ep_value TYPE zexcel_cell_value.

* Check around the "Excel Leap Year" 1900
    ep_value = zcl_excel_common=>date_to_excel_string( '19000228' ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = 59
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD date_to_excel_string3.
    DATA ep_value TYPE zexcel_cell_value.

    ep_value = zcl_excel_common=>date_to_excel_string( '19000301' ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = 61
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD date_to_excel_string4.
    DATA ep_value TYPE zexcel_cell_value.

* Test 2. Basic conversion
    ep_value = zcl_excel_common=>date_to_excel_string( '99991212' ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = 2958446
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD date_to_excel_string5.
    DATA ep_value TYPE zexcel_cell_value.

* Test 3. Initial date
    DATA: lv_date TYPE d.
    ep_value = zcl_excel_common=>date_to_excel_string( lv_date ).

    cl_abap_unit_assert=>assert_equals(
      act   = ep_value
      exp   = ''
      msg   = 'Wrong date conversion'
      level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD date_to_excel_string6.
    DATA ep_value TYPE zexcel_cell_value.

* Test 2. Basic conversion
    DATA exp_value TYPE zexcel_cell_value VALUE 0.
    ep_value = zcl_excel_common=>date_to_excel_string( '18991231' ).

    cl_abap_unit_assert=>assert_differs(
          act   = ep_value
          exp   = exp_value
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.


  METHOD encrypt_password.
* ========================
    DATA lv_encrypted_pwd TYPE zexcel_aes_password.

    TRY.
        lv_encrypted_pwd = zcl_excel_common=>encrypt_password( 'test' ).

        cl_abap_unit_assert=>assert_equals(
              act   = lv_encrypted_pwd
              exp   = 'CBEB'
              msg   = 'Wrong password encryption'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
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

        cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '00000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
* Check empty content
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '' ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '00000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
* Check space character
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( ` ` ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '00000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
* Check first Excel date 1/1/1900
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '1' ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '19000101'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
* Check around the "Excel Leap Year" 1900
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '59' ).

        cl_abap_unit_assert=>assert_equals(
              act   = ep_value
              exp   = '19000228'
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '61' ).

        cl_abap_unit_assert=>assert_equals(
              act   = ep_value
              exp   = '19000301'
              msg   = 'Wrong date conversion'
              level = if_aunit_constants=>critical
            ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 2. Simple test
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '1' ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '19000101'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 3. Last possible date
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '2958465' ).

        cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '99991231'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical
        ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'unexpected exception'
            level  = if_aunit_constants=>critical    " Error Severity
        ).
    ENDTRY.

* Test 4. Exception should be thrown index out of bounds
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_date( '2958466' ).

        cl_abap_unit_assert=>fail(
          msg   = |Unexpected result '{ ep_value }'|
          level = if_aunit_constants=>critical
        ).

      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>assert_equals(
          act   = lx_excel->error
          exp   = 'Unable to interpret date'
          msg   = 'Time should be a valid date'
          level = if_aunit_constants=>fatal
        ).
    ENDTRY.
  ENDMETHOD.       "excel_String_To_Date


  METHOD excel_string_to_time1.
    DATA ep_value TYPE t.

* Test 1. Simple test
    ep_value = zcl_excel_common=>excel_string_to_time( '0' ).

    cl_abap_unit_assert=>assert_equals(
      act   = ep_value
      exp   = '000000'
      msg   = 'Wrong date conversion'
      level = if_aunit_constants=>tolerable ).

  ENDMETHOD.

  METHOD excel_string_to_time2.
    DATA ep_value TYPE t.
* Test 2. Simple test

    ep_value = zcl_excel_common=>excel_string_to_time( '1' ).

    cl_abap_unit_assert=>assert_equals(
      act   = ep_value
      exp   = '000000'
      msg   = 'Wrong date conversion'
      level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD excel_string_to_time3.
    DATA ep_value TYPE t.
* Test 3. Simple test

    ep_value = zcl_excel_common=>excel_string_to_time( '0.99999' ).

    cl_abap_unit_assert=>assert_equals(
      act   = ep_value
      exp   = '235959'
      msg   = 'Wrong date conversion'
      level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD excel_string_to_time4.
    DATA ep_value TYPE t.
* Test 4. Also string greater than 1 should be managed

    ep_value = zcl_excel_common=>excel_string_to_time( '4.1' ).

    cl_abap_unit_assert=>assert_equals(
      act   = ep_value
      exp   = '022400'
      msg   = 'Wrong date conversion'
      level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD excel_string_to_time5.
    DATA ep_value TYPE t.
* Test 4. string is not a number
    TRY.
        ep_value = zcl_excel_common=>excel_string_to_time( 'NaN' ).

        cl_abap_unit_assert=>assert_differs(
          act   = ep_value
          exp   = '000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).
      CATCH zcx_excel INTO lx_excel.
        cl_abap_unit_assert=>assert_equals(
          act   = lx_excel->error
          exp   = 'Unable to interpret time'
          msg   = 'Time should be a valid string'
          level = if_aunit_constants=>fatal ).
    ENDTRY.
  ENDMETHOD.

  METHOD time_to_excel_string1.
    DATA ep_value TYPE zexcel_cell_value.

* Test 1. Basic conversion
    ep_value = zcl_excel_common=>time_to_excel_string( '000001' ).
    " A test directly in Excel returns the value 0.0000115740740740741000
    cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '0.0000115740740741'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD time_to_excel_string2.
    DATA ep_value TYPE zexcel_cell_value.

* Test 2. Basic conversion
    ep_value = zcl_excel_common=>time_to_excel_string( '235959' ).
    " A test directly in Excel returns the value 0.9999884259259260000000
    cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '0.9999884259259260'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD time_to_excel_string3.
    DATA ep_value TYPE zexcel_cell_value.

* Test 3. Initial date
    ep_value = zcl_excel_common=>time_to_excel_string( '000000' ).

    cl_abap_unit_assert=>assert_equals(
      act   = ep_value
      exp   = '0'
      msg   = 'Wrong date conversion'
      level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD time_to_excel_string4.

    DATA ep_value TYPE zexcel_cell_value.

* Test 2. Basic conversion
    ep_value = zcl_excel_common=>time_to_excel_string( '022400' ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_value
          exp   = '0.1000000000000000'
          msg   = 'Wrong date conversion'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD split_file.
* ============================

    DATA: ep_file         TYPE text255,
          ep_extension    TYPE char10,
          ep_dotextension TYPE char10.


* Test 1. Basic conversion
    zcl_excel_common=>split_file( EXPORTING ip_file         = 'filename.xml'
                                  IMPORTING ep_file         = ep_file
                                            ep_extension    = ep_extension
                                            ep_dotextension = ep_dotextension ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_file
          exp   = 'filename'
          msg   = 'Split filename failed'
          level = if_aunit_constants=>critical ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_extension
          exp   = 'xml'
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_dotextension
          exp   = '.xml'
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

* Test 2. no extension
    zcl_excel_common=>split_file( EXPORTING ip_file         = 'filename'
                                  IMPORTING ep_file         = ep_file
                                            ep_extension    = ep_extension
                                            ep_dotextension = ep_dotextension ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_file
          exp   = 'filename'
          msg   = 'Split filename failed'
          level = if_aunit_constants=>critical ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_extension
          exp   = ''
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

    cl_abap_unit_assert=>assert_equals(
          act   = ep_dotextension
          exp   = ''
          msg   = 'Split extension failed'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.  "split_file

  METHOD convert_range2column_a_row.
    DATA: lv_range TYPE string,
          ls_exp   TYPE ty_convert_range2column_a_row.

* a) input empty --> nothing to do
    lv_range = ''.

    CLEAR ls_exp.

    assert_convert_range2column_a_( i_range = lv_range is_exp = ls_exp ).

* b) sheetname existing - starts with '            example 'Sheet 1'!$B$6:$D$13
    lv_range = `'Sheet 1'!$B$6:$D$13`.

    CLEAR ls_exp.
    ls_exp-column_start     = 'B'.
    ls_exp-column_start_int = 2.
    ls_exp-column_end       = 'D'.
    ls_exp-column_end_int   = 4.
    ls_exp-row_start        = 6.
    ls_exp-row_end          = 13.
    ls_exp-sheet            = 'Sheet 1'.

    assert_convert_range2column_a_( i_range = lv_range is_exp = ls_exp ).

* c) sheetname existing - does not start with '    example Sheet1!$B$6:$D$13
    lv_range = `Sheet1!B6:$D$13`.

    CLEAR ls_exp.
    ls_exp-column_start     = 'B'.
    ls_exp-column_start_int = 2.
    ls_exp-column_end       = 'D'.
    ls_exp-column_end_int   = 4.
    ls_exp-row_start        = 6.
    ls_exp-row_end          = 13.
    ls_exp-sheet            = 'Sheet1'.

    assert_convert_range2column_a_( i_range = lv_range is_exp = ls_exp ).

* d) no sheetname - just area                      example $B$6:$D$13
    lv_range = `$B$6:D13`.

    CLEAR ls_exp.
    ls_exp-column_start     = 'B'.
    ls_exp-column_start_int = 2.
    ls_exp-column_end       = 'D'.
    ls_exp-column_end_int   = 4.
    ls_exp-row_start        = 6.
    ls_exp-row_end          = 13.

    assert_convert_range2column_a_( i_range = lv_range is_exp = ls_exp ).

**********************************************************************
* 1 Dimensional Ranges - Ros or Cols Only (eg Print Tiles)
*
    lv_range = `$2:$7`.

    CLEAR ls_exp.

    assert_convert_range2column_a_( i_range = lv_range i_allow_1dim_range = abap_false is_exp = ls_exp ).

***
    lv_range = `$2:$7`.

    CLEAR ls_exp.
    ls_exp-row_start        = 2.
    ls_exp-row_end          = 7.

    assert_convert_range2column_a_( i_range = lv_range i_allow_1dim_range = abap_true is_exp = ls_exp ).

***
    lv_range = `Sheet3!$D:$I`.

    CLEAR ls_exp.
    ls_exp-column_start     = 'D'.
    ls_exp-column_start_int = 4.
    ls_exp-column_end       = 'I'.
    ls_exp-column_end_int   = 9.
    ls_exp-sheet            = 'Sheet3'.

    assert_convert_range2column_a_( i_range = lv_range i_allow_1dim_range = abap_true is_exp = ls_exp ).

  ENDMETHOD.                    "convert_range2column_a_row


  METHOD assert_convert_range2column_a_.

    DATA: ls_act     TYPE ty_convert_range2column_a_row,
          lv_message TYPE string.

    zcl_excel_common=>convert_range2column_a_row(
      EXPORTING
        i_range            = i_range
        i_allow_1dim_range = i_allow_1dim_range
      IMPORTING
        e_column_start     = ls_act-column_start
        e_column_start_int = ls_act-column_start_int
        e_column_end       = ls_act-column_end
        e_column_end_int   = ls_act-column_end_int
        e_row_start        = ls_act-row_start
        e_row_end          = ls_act-row_end
        e_sheet            = ls_act-sheet ).

    lv_message = `Invalid column start (alpha) for ` && i_range.
    cl_abap_unit_assert=>assert_equals(
          act   = ls_act-column_start
          exp   = is_exp-column_start
          msg   = lv_message ).
    lv_message = `Invalid column start (numeric) for ` && i_range.
    cl_abap_unit_assert=>assert_equals(
          act   = ls_act-column_start_int
          exp   = is_exp-column_start_int
          msg   = lv_message ).
    lv_message = `Invalid column end (alpha) for ` && i_range.
    cl_abap_unit_assert=>assert_equals(
          act   = ls_act-column_end
          exp   = is_exp-column_end
          msg   = lv_message ).
    lv_message = `Invalid column end (numeric) for ` && i_range.
    cl_abap_unit_assert=>assert_equals(
          act   = ls_act-column_end_int
          exp   = is_exp-column_end_int
          msg   = lv_message ).
    lv_message = `Invalid row start for ` && i_range.
    cl_abap_unit_assert=>assert_equals(
          act   = ls_act-row_start
          exp   = is_exp-row_start
          msg   = lv_message ).
    lv_message = `Invalid row end for ` && i_range.
    cl_abap_unit_assert=>assert_equals(
          act   = ls_act-row_end
          exp   = is_exp-row_end
          msg   = lv_message ).
    lv_message = `Invalid sheet for ` && i_range.
    cl_abap_unit_assert=>assert_equals(
          act   = ls_act-sheet
          exp   = is_exp-sheet
          msg   = lv_message ).

  ENDMETHOD.


  METHOD describe_structure.
    DATA: ls_test TYPE zexcel_pane.
    DATA: lo_structdescr TYPE REF TO cl_abap_structdescr.
    DATA: lt_structure TYPE ddfields.
    FIELD-SYMBOLS: <line> LIKE LINE OF lt_structure.

    " Test with DDIC Type
    lo_structdescr ?= cl_abap_structdescr=>describe_by_data( p_data = ls_test ).
    lt_structure = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).
    READ TABLE lt_structure ASSIGNING <line> INDEX 1.
    cl_abap_unit_assert=>assert_equals(
          act   = <line>-fieldname
          exp   = 'YSPLIT'
          msg   = 'Describe structure failed'
          level = if_aunit_constants=>critical ).

    " Test with local defined structure having DDIC and non DDIC elements
    TYPES:
      BEGIN OF t_test,
        carrid   TYPE string,
        carrname TYPE string,
        carrdesc TYPE string,
      END OF t_test.
    DATA: ls_ttest TYPE t_test.

    lo_structdescr ?= cl_abap_structdescr=>describe_by_data( p_data = ls_ttest ).
    lt_structure = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).
    READ TABLE lt_structure ASSIGNING <line> INDEX 1.
    cl_abap_unit_assert=>assert_equals(
          act   = <line>-fieldname
          exp   = 'CARRID'
          msg   = 'Describe structure failed'
          level = if_aunit_constants=>critical ).

  ENDMETHOD.                    "describe_structure

  METHOD macro_calculate_cell_distance.

    DATA: lv_offset_rows TYPE i,
          lv_offset_cols TYPE i,
          lv_message     TYPE string.

    zcl_excel_common=>calculate_cell_distance( EXPORTING iv_reference_cell = iv_reference_cell
                                                         iv_current_cell   = iv_current_cell
                                               IMPORTING ev_row_difference = lv_offset_rows
                                                         ev_col_difference = lv_offset_cols ).
* Check delta columns
    CONCATENATE 'Error calculating column difference in test:'
                iv_reference_cell
                '->'
                iv_current_cell
         INTO lv_message SEPARATED BY space.
    cl_abap_unit_assert=>assert_equals(  act   = lv_offset_cols
                                      exp   = iv_expected_column
                                      msg   = lv_message
                                      quit  = 0  " continue tests
                                      level = if_aunit_constants=>critical ).
* Check delta rows
    CONCATENATE 'Error calculating row difference in test:'
                iv_reference_cell
                '->'
                iv_current_cell
         INTO lv_message SEPARATED BY space.
    cl_abap_unit_assert=>assert_equals(  act   = lv_offset_rows
                                      exp   = iv_expected_row
                                      msg   = lv_message
                                      quit  = 0  " continue tests
                                      level = if_aunit_constants=>critical ).

  ENDMETHOD.

  METHOD calc_cell_dist_samecell.

    " Same cell
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'C12'
      iv_expected_column = 0
      iv_expected_row    = 0 ).
  ENDMETHOD.

  METHOD calc_cell_dist_down1pl.

    " Shift down 1 place
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'C13'
      iv_expected_column = 0
      iv_expected_row    = 1 ).
  ENDMETHOD.

  METHOD calc_cell_dist_downsome.

    " Shift down some places
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'C25'
      iv_expected_column = 0
      iv_expected_row    = 13 ).
  ENDMETHOD.

  METHOD calc_cell_dist_up1pl.

    " Shift up 1 place
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'C11'
      iv_expected_column = 0
      iv_expected_row    = -1 ).
  ENDMETHOD.

  METHOD calc_cell_dist_upsome.

    " Shift up some place
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'C1'
      iv_expected_column = 0
      iv_expected_row    = -11 ).
  ENDMETHOD.

  METHOD calc_cell_dist_right1pl.

    " Shift right 1 place
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'D12'
      iv_expected_column = 1
      iv_expected_row    = 0 ).
  ENDMETHOD.

  METHOD calc_cell_dist_rightsome.

    " Shift right some places
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'AA12'
      iv_expected_column = 24
      iv_expected_row    = 0 ).
  ENDMETHOD.

  METHOD calc_cell_dist_left1pl.

    " Shift left 1 place
    macro_calculate_cell_distance(
      iv_reference_cell  = 'C12'
      iv_current_cell    = 'B12'
      iv_expected_column = -1
      iv_expected_row    = 0 ).
  ENDMETHOD.

  METHOD calc_cell_dist_leftsome.

    " Shift left some place
    macro_calculate_cell_distance(
      iv_reference_cell  = 'AA12'
      iv_current_cell    = 'C12'
      iv_expected_column = -24
      iv_expected_row    = 0 ).
  ENDMETHOD.

  METHOD calc_cell_dist_fullpack.

    " The full package.
    macro_calculate_cell_distance(
      iv_reference_cell  = 'AA121'
      iv_current_cell    = 'C12'
      iv_expected_column = -24
      iv_expected_row    = -109 ).
  ENDMETHOD.

  METHOD macro_shift_formula.

    DATA: lv_resulting_formula TYPE string,
          lv_message           TYPE string,
          lv_counter           TYPE n LENGTH 8.

    ADD 1 TO lv_counter.
    CLEAR lv_resulting_formula.
    TRY.
        lv_resulting_formula = zcl_excel_common=>shift_formula( iv_reference_formula = iv_reference_formula
                                                                iv_shift_cols        = iv_shift_cols
                                                                iv_shift_rows        = iv_shift_rows ).
        CONCATENATE 'Wrong result in test'
                    lv_counter
                    'shifting formula '
                    iv_reference_formula
             INTO lv_message SEPARATED BY space.
        cl_abap_unit_assert=>assert_equals(  act   = lv_resulting_formula
                                          exp   = iv_expected
                                          msg   = lv_message
                                          quit  = 0  " continue tests
                                          level = if_aunit_constants=>critical ).
      CATCH zcx_excel.
        CONCATENATE 'Unexpected exception occurred in test'
                    lv_counter
                    'shifting formula '
                    iv_reference_formula
             INTO lv_message SEPARATED BY space.
        cl_abap_unit_assert=>assert_equals(  act   = lv_resulting_formula
                                          exp   = iv_expected
                                          msg   = lv_message
                                          quit  = 0  " continue tests
                                          level = if_aunit_constants=>critical ).
    ENDTRY.

  ENDMETHOD.

  METHOD shift_formula_basic.

    " Very basic check
    macro_shift_formula(
      iv_reference_formula = 'C17'
      iv_shift_cols        = 0
      iv_shift_rows        = 0
      iv_expected          = 'C17' ).

  ENDMETHOD.

  METHOD shift_formula_rightdown.

    " Check shift right and down
    macro_shift_formula(
      iv_reference_formula = 'C17'
      iv_shift_cols        = 2
      iv_shift_rows        = 3
      iv_expected          = 'E20' ).

  ENDMETHOD.

  METHOD shift_formula_leftup.

    " Check shift left and up
    macro_shift_formula(
      iv_reference_formula = 'C17'
      iv_shift_cols        = -2
      iv_shift_rows        = -3
      iv_expected          = 'A14' ).

  ENDMETHOD.

  METHOD shift_formula_fixedcolrows.

    " Fixed columns/rows
    macro_shift_formula(
      iv_reference_formula = '$C$17'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = '$C$17' ).

  ENDMETHOD.

  METHOD shift_formula_mixedfixedrows.

    " Operators and Ranges, mixed fixed rows or columns
    macro_shift_formula(
      iv_reference_formula = 'SUM($C17:C$23)+C30'
      iv_shift_cols        = 1
      iv_shift_rows        = 11
      iv_expected          = 'SUM($C28:D$23)+D41' ).

  ENDMETHOD.

  METHOD shift_formula_rangename.

    " Operators and Rangename
    macro_shift_formula(
      iv_reference_formula = 'RNGNAME1+C7'
      iv_shift_cols        = -1
      iv_shift_rows        = -4
      iv_expected          = 'RNGNAME1+B3' ).

  ENDMETHOD.

  METHOD shift_formula_stringlitconc.

    " String literals and string concatenation
    macro_shift_formula(
      iv_reference_formula = '"Date:"&TEXT(B2)'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = '"Date:"&TEXT(C3)' ).

  ENDMETHOD.

  METHOD shift_formula_extref.

    " External sheet reference
    macro_shift_formula(
      iv_reference_formula = '[TEST6.XLSX]SHEET1!A1'
      iv_shift_cols        = 1
      iv_shift_rows        = 11
      iv_expected          = '[TEST6.XLSX]SHEET1!B12' ).

  ENDMETHOD.

  METHOD shift_formula_charblanks.

    " superflous blanks, multi-argument functions, literals in function, unknown functions
    macro_shift_formula(
      iv_reference_formula = `X(B13, "KK" )  `
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = `X(C14, "KK" )  ` ).

  ENDMETHOD.

  METHOD shift_formula_stringblanks.

    " same as above - but with string input instead of Char-input
    macro_shift_formula(
      iv_reference_formula = `SIN(SIN(SIN(SIN(E22))))`
      iv_shift_cols        = 0
      iv_shift_rows        = 1
      iv_expected          = 'SIN(SIN(SIN(SIN(E23))))' ).

  ENDMETHOD.

  METHOD shift_formula_funcnoargs.

    " Functions w/o arguments, No cellreferences
    macro_shift_formula(
      iv_reference_formula = 'HEUTE()'
      iv_shift_cols        = 2
      iv_shift_rows        = 5
      iv_expected          = 'HEUTE()' ).

  ENDMETHOD.

  METHOD shift_formula_nocellref.

    " No cellreferences
    macro_shift_formula(
      iv_reference_formula = '"B2"'
      iv_shift_cols        = 2
      iv_shift_rows        = 5
      iv_expected          = '"B2"' ).

  ENDMETHOD.

  METHOD shift_formula_empty.

    " Empty
    macro_shift_formula(
      iv_reference_formula = ''
      iv_shift_cols        = 2
      iv_shift_rows        = 5
      iv_expected          = '' ).

  ENDMETHOD.

  METHOD shift_formula_referr_colunder.

    " Referencing error , column only    , underflow
    macro_shift_formula(
      iv_reference_formula = 'A1+$A1+A$1+$A$1+B2'
      iv_shift_cols        = -1
      iv_shift_rows        = 0
      iv_expected          = '#REF!+$A1+#REF!+$A$1+A2' ).

  ENDMETHOD.

  METHOD shift_formula_referr_rowunder.

    " Referencing error , row only       , underflow
    macro_shift_formula(
      iv_reference_formula = 'A1+$A1+A$1+$A$1+B2'
      iv_shift_cols        = 0
      iv_shift_rows        = -1
      iv_expected          = '#REF!+#REF!+A$1+$A$1+B1' ).

  ENDMETHOD.

  METHOD shift_formula_referr_rowcolund.

    " Referencing error , row and column , underflow
    macro_shift_formula(
      iv_reference_formula = 'A1+$A1+A$1+$A$1+B2'
      iv_shift_cols        = -1
      iv_shift_rows        = -1
      iv_expected          = '#REF!+#REF!+#REF!+$A$1+A1' ).

  ENDMETHOD.

  METHOD shift_formula_sheet_nodigit.

    " Sheet name not ending with digit
    macro_shift_formula(
      iv_reference_formula = 'Sheet!A1'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = 'Sheet!B2' ).

  ENDMETHOD.

  METHOD shift_formula_sheet_nodig.

    " Sheet name ending with digit
    macro_shift_formula(
      iv_reference_formula = 'Sheet2!A1'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = 'Sheet2!B2' ).

  ENDMETHOD.

  METHOD shift_formula_sheet_special.

    " Sheet name with special characters
    macro_shift_formula(
      iv_reference_formula = |'Sheet name'!A1|
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = |'Sheet name'!B2| ).

  ENDMETHOD.

  METHOD shift_formula_resp_blanks_1.

    " Respecting blanks
    macro_shift_formula(
      iv_reference_formula = 'SUBTOTAL(109,Table1[SUM 1])'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = 'SUBTOTAL(109,Table1[SUM 1])' ).

  ENDMETHOD.

  METHOD shift_formula_resp_blanks_2.

    " Respecting blanks
    macro_shift_formula(
      iv_reference_formula = 'B4 & C4'
      iv_shift_cols        = 0
      iv_shift_rows        = 1
      iv_expected          = 'B5 & C5' ).

  ENDMETHOD.

  METHOD shift_formula_range.

    " F_1 is a range name, not a cell address
    macro_shift_formula(
      iv_reference_formula = 'SUM(F_1,F_2)'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = 'SUM(F_1,F_2)' ).

  ENDMETHOD.

  METHOD shift_formula_notcols.
    " RC are not columns
    macro_shift_formula(
      iv_reference_formula = 'INDIRECT("RC[4]",FALSE)'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = 'INDIRECT("RC[4]",FALSE)' ).

  ENDMETHOD.

  METHOD shift_formula_name.

    " A1 is a sheet name
    macro_shift_formula(
      iv_reference_formula = |'A1'!$A$1|
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = |'A1'!$A$1| ).

  ENDMETHOD.

  METHOD shift_formula_refcolumn1.

    " Reference to another column in the same row of a Table, with a space in the column name
    macro_shift_formula(
      iv_reference_formula = 'Tbl[[#This Row],[Air fare]]'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = 'Tbl[[#This Row],[Air fare]]' ).

  ENDMETHOD.

  METHOD shift_formula_refcolumn2.

    " Reference to another column in the same row of a Table, inside more complex expression
    macro_shift_formula(
      iv_reference_formula = 'Tbl[[#This Row],[Air]]+A1'
      iv_shift_cols        = 1
      iv_shift_rows        = 1
      iv_expected          = 'Tbl[[#This Row],[Air]]+B2' ).

  ENDMETHOD.

  METHOD is_cell_in_range_ulc_in.
* Test 1: upper left corner (in range)
    DATA ep_cell_in_range TYPE abap_bool.

    TRY.
        ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
            ip_column   = 'B'
            ip_row      = 2
            ip_range    = 'B2:D4' ).

        cl_abap_unit_assert=>assert_equals(
            act   = ep_cell_in_range
            exp   = abap_true
            msg   = 'Check cell in range failed'
            level = if_aunit_constants=>critical ).
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.
  ENDMETHOD. "is_cell_in_range_ulc_in

  METHOD is_cell_in_range_lrc_in.
* Test 2: lower right corner (in range)
    DATA ep_cell_in_range TYPE abap_bool.

    TRY.
        ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
            ip_column   = 'D'
            ip_row      = 4
            ip_range    = 'B2:D4' ).

        cl_abap_unit_assert=>assert_equals(
            act   = ep_cell_in_range
            exp   = abap_true
            msg   = 'Check cell in range failed'
            level = if_aunit_constants=>critical ).
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.
  ENDMETHOD. "is_cell_in_range_lrc_in

  METHOD is_cell_in_range_leftside_out.
* Test 3: left side (out of range)
    DATA ep_cell_in_range TYPE abap_bool.

    TRY.
        ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
            ip_column   = 'A'
            ip_row      = 3
            ip_range    = 'B2:D4' ).

        cl_abap_unit_assert=>assert_equals(
            act   = ep_cell_in_range
            exp   = abap_false
            msg   = 'Check cell in range failed'
            level = if_aunit_constants=>critical ).
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.
  ENDMETHOD. "is_cell_in_range_leftside_out

  METHOD is_cell_in_range_upperside_out.
* Test 4: upper side (out of range)
    DATA ep_cell_in_range TYPE abap_bool.

    TRY.
        ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
            ip_column   = 'C'
            ip_row      = 1
            ip_range    = 'B2:D4' ).

        cl_abap_unit_assert=>assert_equals(
            act   = ep_cell_in_range
            exp   = abap_false
            msg   = 'Check cell in range failed'
            level = if_aunit_constants=>critical ).
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.
  ENDMETHOD. "is_cell_in_range_upperside_out

  METHOD is_cell_in_range_rightside_out.
* Test 5: right side (out of range)
    DATA ep_cell_in_range TYPE abap_bool.

    TRY.
        ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
            ip_column   = 'E'
            ip_row      = 3
            ip_range    = 'B2:D4' ).

        cl_abap_unit_assert=>assert_equals(
            act   = ep_cell_in_range
            exp   = abap_false
            msg   = 'Check cell in range failed'
            level = if_aunit_constants=>critical ).
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.
  ENDMETHOD. "is_cell_in_range_rightside_out

  METHOD is_cell_in_range_lowerside_out.
* Test 6: lower side (out of range)
    DATA ep_cell_in_range TYPE abap_bool.

    TRY.
        ep_cell_in_range = zcl_excel_common=>is_cell_in_range(
            ip_column   = 'C'
            ip_row      = 5
            ip_range    = 'B2:D4' ).

        cl_abap_unit_assert=>assert_equals(
            act   = ep_cell_in_range
            exp   = abap_false
            msg   = 'Check cell in range failed'
            level = if_aunit_constants=>critical ).
      CATCH zcx_excel.
        cl_abap_unit_assert=>fail(
            msg    = 'Unexpected exception'
            level  = if_aunit_constants=>critical ).
    ENDTRY.
  ENDMETHOD. "is_cell_in_range_lowerside_out.

ENDCLASS.
