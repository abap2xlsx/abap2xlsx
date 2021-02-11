class ZCL_EXCEL_COMMON definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
public section.

  constants C_EXCEL_BASELINE_DATE type D value '19000101'. "#EC NOTEXT
  class-data C_EXCEL_NUMFMT_OFFSET type INT1 value 164. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  constants C_EXCEL_SHEET_MAX_COL type INT4 value 16384. "#EC NOTEXT
  constants C_EXCEL_SHEET_MIN_COL type INT4 value 1. "#EC NOTEXT
  constants C_EXCEL_SHEET_MAX_ROW type INT4 value 1048576. "#EC NOTEXT
  constants C_EXCEL_SHEET_MIN_ROW type INT4 value 1. "#EC NOTEXT
  class-data C_SPRAS_EN type SPRAS value 'E'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  class-data O_CONV type ref to CL_ABAP_CONV_OUT_CE .
  constants C_EXCEL_1900_LEAP_YEAR type D value '19000228'. "#EC NOTEXT
  class-data C_XLSX_FILE_FILTER type STRING value 'Excel Workbook (*.xlsx)|*.xlsx|'. "#EC NOTEXT .  .  .  .  .  .  . " .

  class-methods CLASS_CONSTRUCTOR .
  class-methods DESCRIBE_STRUCTURE
    importing
      !IO_STRUCT type ref to CL_ABAP_STRUCTDESCR
    returning
      value(RT_DFIES) type DDFIELDS .
  class-methods CONVERT_COLUMN2ALPHA
    importing
      !IP_COLUMN type SIMPLE
    returning
      value(EP_COLUMN) type ZEXCEL_CELL_COLUMN_ALPHA
    raising
      ZCX_EXCEL .
  class-methods CONVERT_COLUMN2INT
    importing
      !IP_COLUMN type SIMPLE
    returning
      value(EP_COLUMN) type ZEXCEL_CELL_COLUMN
    raising
      ZCX_EXCEL .
  class-methods CONVERT_COLUMNROW2COLUMN_A_ROW
    importing
      !I_COLUMNROW type CLIKE
    exporting
      !E_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA
      !E_ROW type ZEXCEL_CELL_ROW .
  class-methods CONVERT_RANGE2COLUMN_A_ROW
    importing
      !I_RANGE type CLIKE
    exporting
      !E_COLUMN_START type ZEXCEL_CELL_COLUMN_ALPHA
      !E_COLUMN_END type ZEXCEL_CELL_COLUMN_ALPHA
      !E_ROW_START type ZEXCEL_CELL_ROW
      !E_ROW_END type ZEXCEL_CELL_ROW
      !E_SHEET type CLIKE
    raising
      ZCX_EXCEL .
  class-methods DATE_TO_EXCEL_STRING
    importing
      !IP_VALUE type D
    returning
      value(EP_VALUE) type ZEXCEL_CELL_VALUE .
  class-methods ENCRYPT_PASSWORD
    importing
      !I_PWD type ZEXCEL_AES_PASSWORD
    returning
      value(R_ENCRYPTED_PWD) type ZEXCEL_AES_PASSWORD .
  class-methods ESCAPE_STRING
    importing
      !IP_VALUE type CLIKE
    returning
      value(EP_ESCAPED_VALUE) type STRING .
  class-methods UNESCAPE_STRING
    importing
      !IV_ESCAPED type CLIKE
    returning
      value(EV_UNESCAPED_STRING) type STRING
    raising
      ZCX_EXCEL .
  class-methods EXCEL_STRING_TO_DATE
    importing
      !IP_VALUE type ZEXCEL_CELL_VALUE
    returning
      value(EP_VALUE) type D
    raising
      ZCX_EXCEL .
  class-methods EXCEL_STRING_TO_TIME
    importing
      !IP_VALUE type ZEXCEL_CELL_VALUE
    returning
      value(EP_VALUE) type T
    raising
      ZCX_EXCEL .
  class-methods EXCEL_STRING_TO_NUMBER
    importing
      !IP_VALUE type ZEXCEL_CELL_VALUE
    returning
      value(EP_VALUE) type F
    raising
      ZCX_EXCEL .
  class-methods GET_FIELDCATALOG
    importing
      !IP_TABLE type STANDARD TABLE
      !IV_HIDE_MANDT type ABAP_BOOL default ABAP_TRUE
    returning
      value(EP_FIELDCATALOG) type ZEXCEL_T_FIELDCATALOG .
  class-methods NUMBER_TO_EXCEL_STRING
    importing
      value(IP_VALUE) type NUMERIC
    returning
      value(EP_VALUE) type ZEXCEL_CELL_VALUE .
  class-methods RECURSIVE_CLASS_TO_STRUCT
    importing
      !I_SOURCE type ANY
    changing
      !E_TARGET type DATA
      !E_TARGETX type DATA .
  class-methods RECURSIVE_STRUCT_TO_CLASS
    importing
      !I_SOURCE type DATA
      !I_SOURCEX type DATA
    changing
      !E_TARGET type ANY .
  class-methods TIME_TO_EXCEL_STRING
    importing
      !IP_VALUE type T
    returning
      value(EP_VALUE) type ZEXCEL_CELL_VALUE .
  type-pools ABAP .
  class-methods SPLIT_FILE
    importing
      !IP_FILE type TEXT255
    exporting
      !EP_FILE type TEXT255
      !EP_EXTENSION type CHAR10
      !EP_DOTEXTENSION type CHAR10 .
  class-methods CALCULATE_CELL_DISTANCE
    importing
      !IV_REFERENCE_CELL type CLIKE
      !IV_CURRENT_CELL type CLIKE
    exporting
      !EV_ROW_DIFFERENCE type I
      !EV_COL_DIFFERENCE type I
    raising
      ZCX_EXCEL .
  class-methods DETERMINE_RESULTING_FORMULA
    importing
      !IV_REFERENCE_CELL type CLIKE
      !IV_REFERENCE_FORMULA type CLIKE
      !IV_CURRENT_CELL type CLIKE
    returning
      value(EV_RESULTING_FORMULA) type STRING
    raising
      ZCX_EXCEL .
  class-methods SHIFT_FORMULA
    importing
      !IV_REFERENCE_FORMULA type CLIKE
      value(IV_SHIFT_COLS) type I
      value(IV_SHIFT_ROWS) type I
    returning
      value(EV_RESULTING_FORMULA) type STRING
    raising
      ZCX_EXCEL .
  class-methods IS_CELL_IN_RANGE
    importing
      !IP_COLUMN type SIMPLE
      !IP_ROW type ZEXCEL_CELL_ROW
      !IP_RANGE type CLIKE
    returning
      value(RP_IN_RANGE) type ABAP_BOOL
    raising
      ZCX_EXCEL .
*"* protected components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
protected section.
private section.

  class-data C_EXCEL_COL_MODULE type INT2 value 64. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .

  class-methods STRUCTURE_CASE
    importing
      !IS_COMPONENT type ABAP_COMPONENTDESCR
    changing
      !XT_COMPONENTS type ABAP_COMPONENT_TAB .
  class-methods STRUCTURE_RECURSIVE
    importing
      !IS_COMPONENT type ABAP_COMPONENTDESCR
    returning
      value(RT_COMPONENTS) type ABAP_COMPONENT_TAB .
  class-methods CHAR2HEX
    importing
      !I_CHAR type CHAR1
    returning
      value(R_HEX) type ZEXCEL_PWD_HASH .
  class-methods SHL01
    importing
      !I_PWD_HASH type ZEXCEL_PWD_HASH
    returning
      value(R_PWD_HASH) type ZEXCEL_PWD_HASH .
  class-methods SHR14
    importing
      !I_PWD_HASH type ZEXCEL_PWD_HASH
    returning
      value(R_PWD_HASH) type ZEXCEL_PWD_HASH .
ENDCLASS.



CLASS ZCL_EXCEL_COMMON IMPLEMENTATION.

METHOD calculate_cell_distance.

  DATA:       lv_reference_row                TYPE i,
              lv_reference_col_alpha          TYPE zexcel_cell_column_alpha,
              lv_reference_col                TYPE i,
              lv_current_row                  TYPE i,
              lv_current_col_alpha            TYPE zexcel_cell_column_alpha,
              lv_current_col                  TYPE i.

*--------------------------------------------------------------------*
* Split reference  cell into numerical row/column representation
*--------------------------------------------------------------------*
  convert_columnrow2column_a_row( EXPORTING
                                    i_columnrow = iv_reference_cell
                                  IMPORTING
                                    e_column    = lv_reference_col_alpha
                                    e_row       = lv_reference_row ).
  lv_reference_col = convert_column2int( lv_reference_col_alpha ).

*--------------------------------------------------------------------*
* Split current  cell into numerical row/column representation
*--------------------------------------------------------------------*
  convert_columnrow2column_a_row( EXPORTING
                                    i_columnrow = iv_current_cell
                                  IMPORTING
                                    e_column    = lv_current_col_alpha
                                    e_row       = lv_current_row ).
  lv_current_col = convert_column2int( lv_current_col_alpha ).

*--------------------------------------------------------------------*
* Calculate row and column difference
* Positive:   Current cell below    reference cell
*         or  Current cell right of reference cell
* Negative:   Current cell above    reference cell
*         or  Current cell left  of reference cell
*--------------------------------------------------------------------*
  ev_row_difference = lv_current_row - lv_reference_row.
  ev_col_difference = lv_current_col - lv_reference_col.

ENDMETHOD.


method CHAR2HEX.

  IF o_conv IS NOT BOUND.
    o_conv = cl_abap_conv_out_ce=>create( endian   = 'L'
                                          ignore_cerr = abap_true
                                          replacement = '#' ).
  ENDIF.

  CALL METHOD o_conv->reset( ).
  CALL METHOD o_conv->write( data = i_char ).
  r_hex+1 = o_conv->get_buffer( ). " x'65' must be x'0065'

endmethod.


METHOD class_constructor.
  c_xlsx_file_filter = 'Excel Workbook (*.xlsx)|*.xlsx|'(005).
ENDMETHOD.


METHOD convert_column2alpha.

  DATA:       lv_uccpi                        TYPE i,
              lv_text                         TYPE sychar02,
              lv_module                       TYPE int4,
              lv_column                       TYPE zexcel_cell_column.

  STATICS:    sv_prev_in  LIKE lv_column,
              sv_prev_out LIKE ep_column.

* Propagate zcx_excel if error occurs           " issue #155 - less restrictive typing for ip_column
  lv_column = convert_column2int( ip_column ).  " issue #155 - less restrictive typing for ip_column

*--------------------------------------------------------------------*
* Check whether column is in allowed range for EXCEL to handle ( 1-16384 )
*--------------------------------------------------------------------*
  IF   lv_column > 16384
    OR lv_column < 1.
    zcx_excel=>raise_text( 'Index out of bounds' ).
  ENDIF.

*--------------------------------------------------------------------*
* Look up for previous succesfull cached result
*--------------------------------------------------------------------*
  IF lv_column = sv_prev_in AND sv_prev_out IS NOT INITIAL.
    ep_column = sv_prev_out.
    RETURN.
  ELSE.
    CLEAR sv_prev_out.
    sv_prev_in  = lv_column.
  ENDIF.

*--------------------------------------------------------------------*
* Build alpha representation of column
*--------------------------------------------------------------------*
  WHILE lv_column GT 0.

    lv_module = ( lv_column - 1 ) MOD 26.
    lv_uccpi  = 65 + lv_module.

    lv_column = ( lv_column - lv_module ) / 26.

    lv_text   = cl_abap_conv_in_ce=>uccpi( lv_uccpi ).
    CONCATENATE lv_text ep_column INTO ep_column.

  ENDWHILE.

*--------------------------------------------------------------------*
* Save succesfull output into cache
*--------------------------------------------------------------------*
  sv_prev_out = ep_column.

ENDMETHOD.


METHOD convert_column2int.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-29
*              - ...
* changes: renaming variables to naming conventions
*          removing unused variables
*          removing commented out code that is inactive for more then half a year
*          message made to support multilinguality
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#246 - error converting lower case column names
*              - Stefan Schmoecker,                          2012-12-29
* changes: translating the correct variable to upper dase
*          adding missing exception if input is a number
*          that is out of bounds
*          adding missing exception if input contains
*          illegal characters like german umlauts
*--------------------------------------------------------------------*

  DATA:       lv_column                       TYPE zexcel_cell_column_alpha,
              lv_column_c                     TYPE char10,
              lv_column_s                     TYPE string,
              lv_errormessage                 TYPE string,                          " Can't pass '...'(abc) to exception-class
              lv_modulo                       TYPE i.

  STATICS:    sv_prev_in  LIKE lv_column_c,
              sv_prev_out LIKE ep_column.
*--------------------------------------------------------------------*
* This module tries to identify which column a user wants to access
* Numbers as input are just passed back, anything else will be converted
* using EXCEL nomenclatura A = 1, AA = 27, ..., XFD = 16384
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* Normalize input ( upper case , no gaps )
*--------------------------------------------------------------------*
  lv_column_c = ip_column.
*  TRANSLATE lv_column TO UPPER CASE.                       " Fix #246
  TRANSLATE lv_column_c TO UPPER CASE.                      " Fix #246
  CONDENSE lv_column_c NO-GAPS.
  IF lv_column_c EQ ''.
*    lv_errormessage = 'Unable to interpret input as column'(003).
    MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
    zcx_excel=>raise_symsg( ).
  ENDIF.

*--------------------------------------------------------------------*
* Look up for previous succesfull cached result
*--------------------------------------------------------------------*
  IF lv_column_c = sv_prev_in AND sv_prev_out IS NOT INITIAL.
    ep_column = sv_prev_out.
    RETURN.
  ELSE.
    CLEAR sv_prev_out.
    sv_prev_in  = lv_column_c.
  ENDIF.

*--------------------------------------------------------------------*
* If a number gets passed, just convert it to an integer and return
* the converted value
*--------------------------------------------------------------------*
  TRY.
      IF lv_column_c CO '1234567890 '.                      " Fix #164
        ep_column = lv_column_c.                            " Fix #164
*--------------------------------------------------------------------*
* Maximum column for EXCEL:  XFD = 16384    " if anyone has a reference for this information - please add here instead of this comment
*--------------------------------------------------------------------*
        IF ep_column > 16384 OR ep_column < 1.
          lv_errormessage = 'Index out of bounds'(004).
          zcx_excel=>raise_text( lv_errormessage ).
        ENDIF.
        EXIT.
      ENDIF.
    CATCH cx_sy_conversion_no_number.                   "#EC NO_HANDLER
      " Try the character-approach if approach via number has failed
  ENDTRY.

*--------------------------------------------------------------------*
* Raise error if unexpected characters turns up
*--------------------------------------------------------------------*
  lv_column_s = lv_column_c.
  IF lv_column_s CN sy-abcde.
*    lv_errormessage = 'Unable to interpret input as column'(003).
    MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
    zcx_excel=>raise_symsg( ).
  ENDIF.

  DO 1 TIMES. "Because of using CHECK
*--------------------------------------------------------------------*
* Interpret input as number to base 26 with A=1, ... Z=26
* Raise error if unexpected character turns up
*--------------------------------------------------------------------*
* 1st character
*--------------------------------------------------------------------*
    lv_column = lv_column_c.
    lv_modulo = cl_abap_conv_out_ce=>uccpi( lv_column+0(1) ) MOD zcl_excel_common=>c_excel_col_module.
    IF lv_modulo < 1 OR lv_modulo > 26.
*    lv_errormessage = 'Unable to interpret input as column'(003).
      MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
      zcx_excel=>raise_symsg( ).
    ENDIF.
    ep_column = lv_modulo.                    " Leftmost digit

*--------------------------------------------------------------------*
* 2nd character if present
*--------------------------------------------------------------------*
    CHECK lv_column+1(1) IS NOT INITIAL.      " No need to continue if string ended
    lv_modulo = cl_abap_conv_out_ce=>uccpi( lv_column+1(1) ) MOD zcl_excel_common=>c_excel_col_module.
    IF lv_modulo < 1 OR lv_modulo > 26.
*    lv_errormessage = 'Unable to interpret input as column'(003).
      MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
      zcx_excel=>raise_symsg( ).
    ENDIF.
    ep_column = 26 * ep_column + lv_modulo.   " if second digit is present first digit is for 26^1

*--------------------------------------------------------------------*
* 3rd character if present
*--------------------------------------------------------------------*
    CHECK lv_column+2(1) IS NOT INITIAL.      " No need to continue if string ended
    lv_modulo = cl_abap_conv_out_ce=>uccpi( lv_column+2(1) ) MOD zcl_excel_common=>c_excel_col_module.
    IF lv_modulo < 1 OR lv_modulo > 26.
*    lv_errormessage = 'Unable to interpret input as column'(003).
      MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
      zcx_excel=>raise_symsg( ).
    ENDIF.
    ep_column = 26 * ep_column + lv_modulo.   " if third digit is present first digit is for 26^2 and second digit for 26^1
  ENDDO.

*--------------------------------------------------------------------*
* Maximum column for EXCEL:  XFD = 16384    " if anyone has a reference for this information - please add here instead of this comment
*--------------------------------------------------------------------*
  IF ep_column > 16384 OR ep_column < 1.
    lv_errormessage = 'Index out of bounds'(004).
    zcx_excel=>raise_text( lv_errormessage ).
  ENDIF.

*--------------------------------------------------------------------*
* Save succesfull output into cache
*--------------------------------------------------------------------*
  sv_prev_out = ep_column.

ENDMETHOD.


method CONVERT_COLUMNROW2COLUMN_A_ROW.
*--------------------------------------------------------------------*
  "issue #256 - replacing char processing with regex
*--------------------------------------------------------------------*
* Stefan Schmoecker, 2013-08-11
*    Allow input to be CLIKE instead of STRING
*--------------------------------------------------------------------*

  DATA: pane_cell_row_a     TYPE string,
        lv_columnrow        type string.

  lv_columnrow = i_columnrow.    " Get rid of trailing blanks

  FIND REGEX '^(\D+)(\d+)$' IN lv_columnrow SUBMATCHES e_column
                                                       pane_cell_row_a.
  e_row = pane_cell_row_a.

endmethod.


method CONVERT_RANGE2COLUMN_A_ROW.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-07
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          added exceptionclass
*          added errorhandling for invalid range
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#241 - error when sheetname contains "!"
*           - sheetname should be returned unescaped
*              - Stefan Schmoecker,                          2012-12-07
* changes: changed coding to support sheetnames with "!"
*          unescaping sheetname
*--------------------------------------------------------------------*
* issue#155 - lessening restrictions of input parameters
*              - Stefan Schmoecker,                          2012-12-07
* changes: i_range changed to clike
*          e_sheet changed to clike
*--------------------------------------------------------------------*

  DATA:       lv_sheet                        TYPE string,
              lv_range                        TYPE string,
              lv_columnrow_start              TYPE string,
              lv_columnrow_end                TYPE string,
              lv_errormessage                 TYPE string.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Split input range into sheetname and Area
* 4 cases - a) input empty --> nothing to do
*         - b) sheetname existing - starts with '            example 'Sheet 1'!$B$6:$D$13
*         - c) sheetname existing - does not start with '    example Sheet1!$B$6:$D$13
*         - d) no sheetname - just area                      example $B$6:$D$13
*--------------------------------------------------------------------*
* Initialize output parameters
  CLEAR: e_column_start,
         e_column_end,
         e_row_start,
         e_row_end,
         e_sheet.

  IF i_range IS INITIAL.                                " a) input empty --> nothing to do
    EXIT.

  ELSEIF i_range(1) = `'`.                              " b) sheetname existing - starts with '
    FIND REGEX '\![^\!]*$' IN i_range MATCH OFFSET sy-fdpos.  " Find last !
    IF sy-subrc = 0.
      lv_sheet = i_range(sy-fdpos).
      ADD 1 TO sy-fdpos.
      lv_range = i_range.
      SHIFT lv_range LEFT BY sy-fdpos PLACES.
    ELSE.
      lv_errormessage = 'Invalid range'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

  ELSEIF i_range CS '!'.                                " c) sheetname existing - does not start with '
    SPLIT i_range AT '!' INTO lv_sheet lv_range.
      " begin Dennis Schaaf
      IF lv_range CP '*#REF*'.
        lv_errormessage = 'Invalid range'(001).
        zcx_excel=>raise_text( lv_errormessage ).
      ENDIF.
      " end Dennis Schaaf
  ELSE.                                                 " d) no sheetname - just area
    lv_range = i_range.
  ENDIF.

  REPLACE ALL OCCURRENCES OF '$' IN lv_range WITH ''.
  SPLIT lv_range AT ':' INTO lv_columnrow_start lv_columnrow_end.

  convert_columnrow2column_a_row( EXPORTING
                                    i_columnrow = lv_columnrow_start
                                  IMPORTING
                                    e_column    = e_column_start
                                    e_row       = e_row_start ).
  convert_columnrow2column_a_row( EXPORTING
                                    i_columnrow = lv_columnrow_end
                                  IMPORTING
                                    e_column    = e_column_end
                                    e_row       = e_row_end ).

  e_sheet = unescape_string( lv_sheet ).                  " Return in unescaped form
endmethod.


method DATE_TO_EXCEL_STRING.
  DATA: lv_date_diff         TYPE i.

  CHECK ip_value IS NOT INITIAL
    AND ip_value <> space.
  " Needed hack caused by the problem that:
  " Excel 2000 incorrectly assumes that the year 1900 is a leap year
  " http://support.microsoft.com/kb/214326/en-us
  IF ip_value > c_excel_1900_leap_year.
    lv_date_diff = ip_value - c_excel_baseline_date + 2.
  ELSE.
    lv_date_diff = ip_value - c_excel_baseline_date + 1.
  ENDIF.
  ep_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_date_diff ).
endmethod.


method DESCRIBE_STRUCTURE.
  DATA:   lt_components      TYPE abap_component_tab,
          lt_comps           TYPE abap_component_tab,
          ls_component       TYPE abap_componentdescr,
          lo_elemdescr       TYPE REF TO cl_abap_elemdescr,
          ls_dfies           TYPE dfies,
          l_position         TYPE tabfdpos.

  "for DDIC structure get the info directly
  IF io_struct->is_ddic_type( ) = abap_true.
    rt_dfies = io_struct->get_ddic_field_list( ).
  ELSE.
    lt_components = io_struct->get_components( ).

    LOOP AT lt_components INTO ls_component.
      structure_case( EXPORTING is_component  = ls_component
                      CHANGING  xt_components = lt_comps   ) .
    ENDLOOP.
    LOOP AT lt_comps INTO ls_component.
      CLEAR ls_dfies.
      IF ls_component-type->kind = cl_abap_typedescr=>kind_elem. "E Elementary Type
        ADD 1 TO l_position.
        lo_elemdescr ?= ls_component-type.
        IF lo_elemdescr->is_ddic_type( ) = abap_true.
          ls_dfies           = lo_elemdescr->get_ddic_field( ).
          ls_dfies-fieldname = ls_component-name.
          ls_dfies-position  = l_position.
        ELSE.
          ls_dfies-fieldname = ls_component-name.
          ls_dfies-position  = l_position.
          ls_dfies-inttype   = lo_elemdescr->type_kind.
          ls_dfies-leng      = lo_elemdescr->length.
          ls_dfies-outputlen = lo_elemdescr->length.
          ls_dfies-decimals  = lo_elemdescr->decimals.
          ls_dfies-fieldtext = ls_component-name.
          ls_dfies-reptext   = ls_component-name.
          ls_dfies-scrtext_s = ls_component-name.
          ls_dfies-scrtext_m = ls_component-name.
          ls_dfies-scrtext_l = ls_component-name.
          ls_dfies-dynpfld   = abap_true.
        ENDIF.
        INSERT ls_dfies INTO TABLE rt_dfies.
      ENDIF.
    ENDLOOP.
  ENDIF.
endmethod.


METHOD determine_resulting_formula.

  DATA:       lv_row_difference               TYPE i,
              lv_col_difference               TYPE i.

*--------------------------------------------------------------------*
* Calculate distance of reference and current cell
*--------------------------------------------------------------------*
  calculate_cell_distance( EXPORTING
                             iv_reference_cell = iv_reference_cell
                             iv_current_cell   = iv_current_cell
                           IMPORTING
                             ev_row_difference = lv_row_difference
                             ev_col_difference = lv_col_difference ).

*--------------------------------------------------------------------*
* and shift formula by using the row- and columndistance
*--------------------------------------------------------------------*
  ev_resulting_formula = shift_formula( iv_reference_formula = iv_reference_formula
                                        iv_shift_rows        = lv_row_difference
                                        iv_shift_cols        = lv_col_difference ).

ENDMETHOD.                    "determine_resulting_formula


method ENCRYPT_PASSWORD.

  DATA lv_curr_offset            TYPE i.
  DATA lv_curr_char              TYPE c LENGTH 1.
  DATA lv_curr_hex               TYPE zexcel_pwd_hash.
  DATA lv_pwd_len                TYPE zexcel_pwd_hash.
  DATA lv_pwd_hash               TYPE zexcel_pwd_hash.

  CONSTANTS:
    lv_0x7fff                    TYPE zexcel_pwd_hash VALUE '7FFF',
    lv_0x0001                    TYPE zexcel_pwd_hash VALUE '0001',
    lv_0xce4b                    TYPE zexcel_pwd_hash VALUE 'CE4B'.

  DATA lv_pwd            TYPE zexcel_aes_password.

  lv_pwd = i_pwd(15).

  lv_pwd_len = STRLEN( lv_pwd ).
  lv_curr_offset = lv_pwd_len - 1.

  WHILE lv_curr_offset GE 0.

    lv_curr_char = lv_pwd+lv_curr_offset(1).
    lv_curr_hex = char2hex( lv_curr_char ).

    lv_pwd_hash = (  shr14( lv_pwd_hash ) BIT-AND lv_0x0001 ) BIT-OR ( shl01( lv_pwd_hash ) BIT-AND lv_0x7fff ).

    lv_pwd_hash = lv_pwd_hash BIT-XOR lv_curr_hex.
    SUBTRACT 1 FROM lv_curr_offset.
  ENDWHILE.

  lv_pwd_hash = (  shr14( lv_pwd_hash ) BIT-AND lv_0x0001 ) BIT-OR ( shl01( lv_pwd_hash ) BIT-AND lv_0x7fff ).
  lv_pwd_hash = lv_pwd_hash BIT-XOR lv_0xce4b.
  lv_pwd_hash = lv_pwd_hash BIT-XOR lv_pwd_len.

  WRITE lv_pwd_hash TO r_encrypted_pwd.

endmethod.


method ESCAPE_STRING.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-08
*              - ...
* changes: aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#242 - Support escaping for white-spaces
*           - Escaping also necessary when ' encountered in input
*              - Stefan Schmoecker,                          2012-12-08
* changes: switched check if escaping is necessary to regular expression
*          and moved the "REPLACE"
*--------------------------------------------------------------------*
* issue#155 - lessening restrictions of input parameters
*              - Stefan Schmoecker,                          2012-12-08
* changes: ip_value changed to clike
*--------------------------------------------------------------------*
  DATA:       lv_value                        TYPE string.

*--------------------------------------------------------------------*
* There exist various situations when a space will be used to separate
* different parts of a string. When we have a string consisting spaces
* that will cause errors unless we "escape" the string by putting ' at
* the beginning and at the end of the string.
*--------------------------------------------------------------------*


*--------------------------------------------------------------------*
* When allowing clike-input parameters we might encounter trailing
* "real" blanks .  These are automatically eliminated when moving
* the input parameter to a string.
* Now any remaining spaces ( white-spaces or normal spaces ) should
* trigger the escaping as well as any '
*--------------------------------------------------------------------*
  lv_value = ip_value.


  FIND REGEX `\s|'` IN lv_value.  " \s finds regular and white spaces
  IF sy-subrc = 0.
    REPLACE ALL OCCURRENCES OF `'` IN lv_value WITH `''`.
    CONCATENATE `'` lv_value `'` INTO lv_value .
  ENDIF.

  ep_escaped_value = lv_value.

endmethod.


method EXCEL_STRING_TO_DATE.
  DATA: lv_date_int TYPE i.

  TRY.
      lv_date_int = ip_value.
      ep_value = lv_date_int + c_excel_baseline_date - 2.
      " Needed hack caused by the problem that:
      " Excel 2000 incorrectly assumes that the year 1900 is a leap year
      " http://support.microsoft.com/kb/214326/en-us
      IF ep_value < c_excel_1900_leap_year.
        ep_value = ep_value + 1.
      ENDIF.
    CATCH cx_sy_conversion_error.
      zcx_excel=>raise_text( 'Index out of bounds' ).
  ENDTRY.
endmethod.


method EXCEL_STRING_TO_NUMBER.

* If we encounter anything more complicated in EXCEL we might have to extend this
* But currently this works fine - even for numbers in scientific notation

  ep_value = ip_value.

endmethod.


method EXCEL_STRING_TO_TIME.
  DATA: lv_seconds_in_day TYPE i,
        lv_day_fraction   TYPE f,
        lc_seconds_in_day TYPE i VALUE 86400.

  TRY.

      lv_day_fraction = ip_value.
      lv_seconds_in_day = lv_day_fraction * lc_seconds_in_day.

      ep_value = lv_seconds_in_day.

    CATCH cx_sy_conversion_error.
      zcx_excel=>raise_text( 'Unable to interpret time' ).
  ENDTRY.
endmethod.

METHOD get_fieldcatalog.
  DATA: lr_dref_tab           TYPE REF TO data,
        lo_salv_table         TYPE REF TO cl_salv_table,
        lo_salv_columns_table TYPE REF TO cl_salv_columns_table,
        lt_salv_t_column_ref  TYPE salv_t_column_ref,
        ls_salv_t_column_ref  LIKE LINE OF lt_salv_t_column_ref,
        lo_salv_column_table  TYPE REF TO cl_salv_column_table.

  FIELD-SYMBOLS: <tab>          TYPE STANDARD TABLE.
  FIELD-SYMBOLS: <fcat>         LIKE LINE OF ep_fieldcatalog.

* Get copy of IP_TABLE-structure <-- must be changeable to create salv
  CREATE DATA lr_dref_tab LIKE ip_table.
  ASSIGN lr_dref_tab->* TO <tab>.
* Create salv --> implicitly create fieldcat
  TRY.
      cl_salv_table=>factory( IMPORTING
                                r_salv_table   = lo_salv_table
                              CHANGING
                                t_table        = <tab>  ).
      lo_salv_columns_table = lo_salv_table->get_columns( ).
      lt_salv_t_column_ref  = lo_salv_columns_table->get( ).
    CATCH cx_root.
* maybe some errorhandling here - just haven't made up my mind yet
  ENDTRY.

* Loop through columns and set relevant fields ( fieldname, texts )
  LOOP AT lt_salv_t_column_ref INTO ls_salv_t_column_ref.

    lo_salv_column_table ?= ls_salv_t_column_ref-r_column.
    APPEND INITIAL LINE TO ep_fieldcatalog ASSIGNING <fcat>.
    <fcat>-position  = sy-tabix.
    <fcat>-fieldname = ls_salv_t_column_ref-columnname.
    <fcat>-scrtext_s = ls_salv_t_column_ref-r_column->get_short_text( ).
    <fcat>-scrtext_m = ls_salv_t_column_ref-r_column->get_medium_text( ).
    <fcat>-scrtext_l = ls_salv_t_column_ref-r_column->get_long_text( ).
    <fcat>-abap_type = lo_salv_column_table->get_ddic_inttype( ).

    <fcat>-dynpfld   = 'X'.  " What in the world would we exclude here?
    " except for the MANDT-field of most tables ( 1st column that is )
    IF <fcat>-position = 1 AND lo_salv_column_table->get_ddic_datatype( ) = 'CLNT' AND iv_hide_mandt = abap_true.
      CLEAR <fcat>-dynpfld.
    ENDIF.

* For fields that don't a description (  i.e. defined by  "field type i," )
* just use the fieldname as description - that is better than nothing
    IF    <fcat>-scrtext_s IS INITIAL
      AND <fcat>-scrtext_m IS INITIAL
      AND <fcat>-scrtext_l IS INITIAL.
      CONCATENATE 'Col:' <fcat>-fieldname INTO <fcat>-scrtext_l  SEPARATED BY space.
      <fcat>-scrtext_m = <fcat>-scrtext_l.
      <fcat>-scrtext_s = <fcat>-scrtext_l.
    ENDIF.

  ENDLOOP.

ENDMETHOD.


method IS_CELL_IN_RANGE.
  DATA lv_column_start    TYPE zexcel_cell_column_alpha.
  DATA lv_column_end      TYPE zexcel_cell_column_alpha.
  DATA lv_row_start       TYPE zexcel_cell_row.
  DATA lv_row_end         TYPE zexcel_cell_row.
  DATA lv_column_start_i  TYPE zexcel_cell_column.
  DATA lv_column_end_i    TYPE zexcel_cell_column.
  DATA lv_column_i        TYPE zexcel_cell_column.


* Split range and convert columns
  convert_range2column_a_row(
    exporting
      i_range        = ip_range
    IMPORTING
      e_column_start = lv_column_start
      e_column_end   = lv_column_end
      e_row_start    = lv_row_start
      e_row_end      = lv_row_end ).

  lv_column_start_i = convert_column2int( ip_column = lv_column_start ).
  lv_column_end_i   = convert_column2int( ip_column = lv_column_end ).

  lv_column_i = convert_column2int( ip_column = ip_column ).

* Check if cell is in range
  IF lv_column_i >= lv_column_start_i AND
     lv_column_i <= lv_column_end_i   AND
     ip_row      >= lv_row_start      AND
     ip_row      <= lv_row_end.
    rp_in_range = abap_true.
  ENDIF.
endmethod.


method NUMBER_TO_EXCEL_STRING.
  DATA: lv_value_c TYPE c LENGTH 100.

  WRITE ip_value TO lv_value_c EXPONENT 0 NO-GROUPING NO-SIGN.
  REPLACE ALL OCCURRENCES OF ',' IN lv_value_c WITH '.'.

  ep_value = lv_value_c.
  CONDENSE ep_value.

  IF ip_value < 0.
    CONCATENATE '-' ep_value INTO ep_value.
  ELSEIF ip_value EQ 0.
    ep_value = '0'.
  ENDIF.
endmethod.


method RECURSIVE_CLASS_TO_STRUCT.
  " # issue 139
* is working for me - but after looking through this coding I guess
* I'll rewrite this to a version w/o recursion
* This is private an no one using it so far except me, so no need to hurry
  DATA: descr           TYPE REF TO cl_abap_structdescr,
        wa_component    LIKE LINE OF descr->components,
        attribute_name  TYPE fieldname,
        flag_class      TYPE flag.

  FIELD-SYMBOLS: <field>      TYPE any,
                 <fieldx>     TYPE any,
                 <attribute>  TYPE any.


  descr ?= cl_abap_structdescr=>describe_by_data( e_target ).

  LOOP AT descr->components INTO wa_component.

* Assign structure and X-structure
    ASSIGN COMPONENT wa_component-name OF STRUCTURE e_target  TO <field>.
    ASSIGN COMPONENT wa_component-name OF STRUCTURE e_targetx TO <fieldx>.
* At least one field in the structure should be marked - otherwise continue with next field
    CLEAR flag_class.
* maybe source is just a structure - try assign component...
    ASSIGN COMPONENT wa_component-name OF STRUCTURE i_source  TO <attribute>.
    IF sy-subrc <> 0.
* not - then it is an attribute of the class - use different assign then
      CONCATENATE 'i_source->' wa_component-name INTO attribute_name.
      ASSIGN (attribute_name) TO <attribute>.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.  " Should not happen if structure is built properly - otherwise just exit to create no dumps
      flag_class = abap_true.
    ENDIF.

    CASE wa_component-type_kind.
      WHEN cl_abap_structdescr=>typekind_struct1 OR cl_abap_structdescr=>typekind_struct2.  " Structure --> use recursio
*        IF flag_class = abap_true.
** Only borders will be passed as unbound references.  But since we want to set a value we have to create an instance
*        ENDIF.
        zcl_excel_common=>recursive_class_to_struct( EXPORTING i_source  = <attribute>
                                                     CHANGING  e_target  = <field>
                                                               e_targetx = <fieldx> ).
      WHEN OTHERS.
        <field> = <attribute>.
        <fieldx> = abap_true.

    ENDCASE.
  ENDLOOP.

endmethod.


METHOD recursive_struct_to_class.
  " # issue 139
* is working for me - but after looking through this coding I guess
* I'll rewrite this to a version w/o recursion
* This is private an no one using it so far except me, so no need to hurry
  DATA: descr           TYPE REF TO cl_abap_structdescr,
        wa_component    LIKE LINE OF descr->components,
        attribute_name  TYPE fieldname,
        flag_class      TYPE flag,
        o_border        TYPE REF TO zcl_excel_style_border.

  FIELD-SYMBOLS: <field>        TYPE any,
                 <fieldx>       TYPE any,
                 <attribute>    TYPE any.


  descr ?= cl_abap_structdescr=>describe_by_data( i_source ).

  LOOP AT descr->components INTO wa_component.

* Assign structure and X-structure
    ASSIGN COMPONENT wa_component-name OF STRUCTURE i_source  TO <field>.
    ASSIGN COMPONENT wa_component-name OF STRUCTURE i_sourcex TO <fieldx>.
* At least one field in the structure should be marked - otherwise continue with next field
    CHECK <fieldx> CA abap_true.
    CLEAR flag_class.
* maybe target is just a structure - try assign component...
    ASSIGN COMPONENT wa_component-name OF STRUCTURE e_target  TO <attribute>.
    IF sy-subrc <> 0.
* not - then it is an attribute of the class - use different assign then
      CONCATENATE 'E_TARGET->' wa_component-name INTO attribute_name.
      ASSIGN (attribute_name) TO <attribute>.
      IF sy-subrc <> 0.EXIT.ENDIF.  " Should not happen if structure is built properly - otherwise just exit to create no dumps
      flag_class = abap_true.
    ENDIF.

    CASE wa_component-type_kind.
      WHEN cl_abap_structdescr=>typekind_struct1 OR cl_abap_structdescr=>typekind_struct2.  " Structure --> use recursion
        " To avoid dump with attribute GRADTYPE of class ZCL_EXCEL_STYLE_FILL
        " quick and really dirty fix -> check the attribute name
        " Border has to be initialized somewhere else
        IF wa_component-name EQ 'GRADTYPE'.
          flag_class = abap_false.
        ENDIF.

        IF flag_class = abap_true AND <attribute> IS INITIAL.
* Only borders will be passed as unbound references.  But since we want to set a value we have to create an instance
          CREATE OBJECT o_border.
          <attribute> = o_border.
        ENDIF.
        zcl_excel_common=>recursive_struct_to_class( EXPORTING i_source  = <field>
                                                               i_sourcex = <fieldx>
                                                     CHANGING  e_target  = <attribute> ).
*      WHEN cl_abap_structdescr=>typekind_struct2.  " String
*        CHECK <fieldx> = abap_true.  " Marked for change
*        <attribute_s> = <field>.
      WHEN OTHERS.
        CHECK <fieldx> = abap_true.  " Marked for change
        <attribute> = <field>.

    ENDCASE.
  ENDLOOP.

ENDMETHOD.


METHOD shift_formula.

  CONSTANTS:  lcv_operators                   TYPE string VALUE '+-/*^%=<>&, !',
              lcv_letters                     TYPE string VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ$',
              lcv_digits                      TYPE string VALUE '0123456789',
              lcv_cell_reference_error        TYPE string VALUE '#REF!'.

  DATA:       lv_tcnt                         TYPE i,         " Counter variable
              lv_tlen                         TYPE i,         " Temp variable length
              lv_cnt                          TYPE i,         " Counter variable
              lv_cnt2                         TYPE i,         " Counter variable
              lv_offset1                      TYPE i,         " Character offset
              lv_numchars                     TYPE i,         " Number of characters counter
              lv_tchar(1)                     TYPE c,         " Temp character
              lv_tchar2(1)                    TYPE c,         " Temp character
              lv_cur_form                     TYPE string,    " Formula for current cell
              lv_ref_cell_addr                TYPE string,    " Reference cell address
              lv_tcol1                        TYPE string,    " Temp column letter
              lv_tcol2                        TYPE string,    " Temp column letter
              lv_tcoln                        TYPE i,         " Temp column number
              lv_trow1                        TYPE string,    " Temp row number
              lv_trow2                        TYPE string,    " Temp row number
              lv_flen                         TYPE i,         " Length of reference formula
              lv_tlen2                        TYPE i,         " Temp variable length
              lv_substr1                      TYPE string,    " Substring variable
              lv_abscol                       TYPE string,    " Absolute column symbol
              lv_absrow                       TYPE string,    " Absolute row symbol
              lv_compare_1                    TYPE string,
              lv_compare_2                    TYPE string,

              lv_errormessage                 TYPE string.

*--------------------------------------------------------------------*
* When copying a cell in EXCEL to another cell any inherent formulas
* are copied as well.  Cell-references in the formula are being adjusted
* by the distance of the new cell to the original one
*--------------------------------------------------------------------*
* §1 Parse reference formula character by character
* §2 Identify Cell-references
* §3 Shift cell-reference
* §4 Build resulting formula
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* No distance --> Reference = resulting cell/formula
*--------------------------------------------------------------------*
  IF    iv_shift_cols = 0
    AND iv_shift_rows = 0.
    ev_resulting_formula = iv_reference_formula.
    EXIT. " done
  ENDIF.


  lv_flen     = STRLEN( iv_reference_formula ).
  lv_numchars = 1.

*--------------------------------------------------------------------*
* §1 Parse reference formula character by character
*--------------------------------------------------------------------*
  DO lv_flen TIMES.

    CLEAR: lv_tchar,
           lv_substr1,
           lv_ref_cell_addr.
    lv_cnt2 = lv_cnt + 1.
    IF lv_cnt2 > lv_flen.
      EXIT. " Done
    ENDIF.

*--------------------------------------------------------------------*
* Here we have the current character in the formula
*--------------------------------------------------------------------*
    lv_tchar = iv_reference_formula+lv_cnt(1).

*--------------------------------------------------------------------*
* Operators or opening parenthesis will separate possible cellreferences
*--------------------------------------------------------------------*
    IF    (    lv_tchar CA lcv_operators
            OR lv_tchar CA '(' )
      AND lv_cnt2 = 1.
      lv_substr1  = iv_reference_formula+lv_offset1(1).
      CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
      lv_cnt      = lv_cnt + 1.
      lv_offset1  = lv_cnt.
      lv_numchars = 1.
      CONTINUE.       " --> next character in formula can be analyzed
    ENDIF.

*--------------------------------------------------------------------*
* Quoted literal text holds no cell reference --> advance to end of text
*--------------------------------------------------------------------*
    IF lv_tchar EQ '"'.
      lv_cnt      = lv_cnt + 1.
      lv_numchars = lv_numchars + 1.
      lv_tchar     = iv_reference_formula+lv_cnt(1).
      WHILE lv_tchar NE '"'.

        lv_cnt      = lv_cnt + 1.
        lv_numchars = lv_numchars + 1.
        lv_tchar    = iv_reference_formula+lv_cnt(1).

      ENDWHILE.
      lv_cnt2    = lv_cnt + 1.
      lv_substr1 = iv_reference_formula+lv_offset1(lv_numchars).
      CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
      lv_cnt     = lv_cnt + 1.
      IF lv_cnt = lv_flen.
        EXIT.
      ENDIF.
      lv_offset1  = lv_cnt.
      lv_numchars = 1.
      lv_tchar    = iv_reference_formula+lv_cnt(1).
      lv_cnt2     = lv_cnt + 1.
      CONTINUE.       " --> next character in formula can be analyzed
    ENDIF.


*--------------------------------------------------------------------*
* Operators or parenthesis or last character in formula will separate possible cellreferences
*--------------------------------------------------------------------*
    IF   lv_tchar CA lcv_operators
      OR lv_tchar CA '():'
      OR lv_cnt2  =  lv_flen.
      IF lv_cnt > 0.
        lv_substr1 = iv_reference_formula+lv_offset1(lv_numchars).
*--------------------------------------------------------------------*
* Check for text concatenation and functions
*--------------------------------------------------------------------*
        IF ( lv_tchar CA lcv_operators AND lv_tchar EQ lv_substr1 ) OR lv_tchar EQ '('.
          CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
          lv_cnt = lv_cnt + 1.
          lv_offset1 = lv_cnt.
          lv_cnt2 = lv_cnt + 1.
          lv_numchars = 1.
          CONTINUE.       " --> next character in formula can be analyzed
        ENDIF.

        lv_tlen = lv_cnt2 - lv_offset1.
*--------------------------------------------------------------------*
* Exclude mathematical operators and closing parentheses
*--------------------------------------------------------------------*
        IF   lv_tchar CA lcv_operators
          OR lv_tchar CA ':)'.
          IF    lv_cnt2     = lv_flen
            AND lv_numchars = 1.
            CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
            lv_cnt      = lv_cnt + 1.
            lv_offset1  = lv_cnt.
            lv_cnt2     = lv_cnt + 1.
            lv_numchars = 1.
            CONTINUE.       " --> next character in formula can be analyzed
          ELSE.
            lv_tlen = lv_tlen - 1.
          ENDIF.
        ENDIF.
*--------------------------------------------------------------------*
* Capture reference cell address
*--------------------------------------------------------------------*
        TRY.
            MOVE: iv_reference_formula+lv_offset1(lv_tlen) TO lv_ref_cell_addr. "Ref cell address
          CATCH cx_root.
            lv_errormessage = 'Internal error in Class ZCL_EXCEL_COMMON Method SHIFT_FORMULA Spot 1 '.  " Change to messageclass if possible
            zcx_excel=>raise_text( lv_errormessage ).
        ENDTRY.

*--------------------------------------------------------------------*
* Split cell address into characters and numbers
*--------------------------------------------------------------------*
        CLEAR: lv_tlen,
               lv_tcnt,
               lv_tcol1,
               lv_trow1.
        lv_tlen = STRLEN( lv_ref_cell_addr ).
        IF lv_tlen <> 0.
          CLEAR: lv_tcnt.
          DO lv_tlen TIMES.
            CLEAR: lv_tchar2.
            lv_tchar2 = lv_ref_cell_addr+lv_tcnt(1).
            IF lv_tchar2 CA lcv_letters.
              CONCATENATE lv_tcol1 lv_tchar2 INTO lv_tcol1.
            ELSEIF lv_tchar2 CA lcv_digits.
              CONCATENATE lv_trow1 lv_tchar2 INTO lv_trow1.
            ENDIF.
            lv_tcnt = lv_tcnt + 1.
          ENDDO.
        ENDIF.

        " Is valid column & row ?
        IF lv_tcol1 IS NOT INITIAL AND lv_trow1 IS NOT INITIAL.
          " COLUMN + ROW
          CONCATENATE lv_tcol1 lv_trow1 INTO lv_compare_1.
          " Original condensed string
          lv_compare_2 = lv_ref_cell_addr.
          CONDENSE lv_compare_2.
          IF lv_compare_1 <> lv_compare_2.
            CLEAR: lv_trow1, lv_tchar2.
          ENDIF.
        ENDIF.

*--------------------------------------------------------------------*
* Check for invalid cell address
*--------------------------------------------------------------------*
        IF lv_tcol1 IS INITIAL OR lv_trow1 IS INITIAL.
          CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
          lv_cnt = lv_cnt + 1.
          lv_offset1 = lv_cnt.
          lv_cnt2 = lv_cnt + 1.
          lv_numchars = 1.
          CONTINUE.
        ENDIF.
*--------------------------------------------------------------------*
* Check for range names
*--------------------------------------------------------------------*
        CLEAR: lv_tlen.
        lv_tlen = STRLEN( lv_tcol1 ).
        IF lv_tlen GT 3.
          CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
          lv_cnt = lv_cnt + 1.
          lv_offset1 = lv_cnt.
          lv_cnt2 = lv_cnt + 1.
          lv_numchars = 1.
          CONTINUE.
        ENDIF.
*--------------------------------------------------------------------*
* Check for valid row
*--------------------------------------------------------------------*
        IF lv_trow1 GT 1048576.
          CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
          lv_cnt = lv_cnt + 1.
          lv_offset1 = lv_cnt.
          lv_cnt2 = lv_cnt + 1.
          lv_numchars = 1.
          CONTINUE.
        ENDIF.
*--------------------------------------------------------------------*
* Check for absolute column or row reference
*--------------------------------------------------------------------*
        CLEAR: lv_tcol2,
               lv_trow2,
               lv_abscol,
               lv_absrow.
        lv_tlen2 = STRLEN( lv_tcol1 ) - 1.
        IF lv_tcol1 IS NOT INITIAL.
          lv_abscol = lv_tcol1(1).
        ENDIF.
        IF lv_tlen2 GE 0.
          lv_absrow = lv_tcol1+lv_tlen2(1).
        ENDIF.
        IF lv_abscol EQ '$' AND lv_absrow EQ '$'.
          lv_tlen2 = lv_tlen2 - 1.
          IF lv_tlen2 > 0.
            lv_tcol1 = lv_tcol1+1(lv_tlen2).
          ENDIF.
          lv_tlen2 = lv_tlen2 + 1.
        ELSEIF lv_abscol EQ '$'.
          lv_tcol1 = lv_tcol1+1(lv_tlen2).
        ELSEIF lv_absrow EQ '$'.
          lv_tcol1 = lv_tcol1(lv_tlen2).
        ENDIF.
*--------------------------------------------------------------------*
* Check for valid column
*--------------------------------------------------------------------*
        TRY.
            lv_tcoln = zcl_excel_common=>convert_column2int( lv_tcol1 ) + iv_shift_cols.
          CATCH zcx_excel.
            CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            CONTINUE.
        ENDTRY.
*--------------------------------------------------------------------*
* Check whether there is a referencing problem
*--------------------------------------------------------------------*
        lv_trow2 = lv_trow1 + iv_shift_rows.
        " Could contain spaces
        CONDENSE lv_trow2.
        IF   ( lv_tcoln < 1 AND lv_abscol <> '$' )   " Maybe we should add here max-column and max row-tests as well.
          OR ( lv_trow2 < 1 AND lv_absrow <> '$' ).  " Check how EXCEL behaves in this case
*--------------------------------------------------------------------*
* Referencing problem encountered --> set error
*--------------------------------------------------------------------*
          CONCATENATE lv_cur_form lcv_cell_reference_error INTO lv_cur_form.
        ELSE.
*--------------------------------------------------------------------*
* No referencing problems --> adjust row and column
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* Adjust column
*--------------------------------------------------------------------*
          IF lv_abscol EQ '$'.
            CONCATENATE lv_cur_form lv_abscol lv_tcol1 INTO lv_cur_form.
          ELSEIF iv_shift_cols EQ 0.
            CONCATENATE lv_cur_form lv_tcol1 INTO lv_cur_form.
          ELSE.
            TRY.
                lv_tcol2 = zcl_excel_common=>convert_column2alpha( lv_tcoln ).
                CONCATENATE lv_cur_form lv_tcol2 INTO lv_cur_form.
              CATCH zcx_excel.
                CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
                lv_cnt = lv_cnt + 1.
                lv_offset1 = lv_cnt.
                lv_cnt2 = lv_cnt + 1.
                lv_numchars = 1.
                CONTINUE.
            ENDTRY.
          ENDIF.
*--------------------------------------------------------------------*
* Adjust row
*--------------------------------------------------------------------*
          IF lv_absrow EQ '$'.
            CONCATENATE lv_cur_form lv_absrow lv_trow1 INTO lv_cur_form.
          ELSEIF iv_shift_rows = 0.
            CONCATENATE lv_cur_form lv_trow1 INTO lv_cur_form.
*        elseif lv_trow2 < 1.
*          CONCATENATE lv_cur_form lc_cell_reference_error INTO lv_cur_form.
          ELSE.
            CONCATENATE lv_cur_form lv_trow2 INTO lv_cur_form.
          ENDIF.
        ENDIF.

        lv_numchars = 0.
        IF   lv_tchar CA lcv_operators
          OR lv_tchar CA ':)'.
          CONCATENATE lv_cur_form lv_tchar INTO lv_cur_form RESPECTING BLANKS.
        ENDIF.
        lv_offset1 = lv_cnt2.
      ENDIF.
    ENDIF.
    lv_numchars = lv_numchars + 1.
    lv_cnt   = lv_cnt   + 1.
    lv_cnt2  = lv_cnt   + 1.

  ENDDO.



*--------------------------------------------------------------------*
* Return resulting formula
*--------------------------------------------------------------------*
  IF lv_cur_form IS NOT INITIAL.
    MOVE lv_cur_form TO ev_resulting_formula.
  ENDIF.

ENDMETHOD.


method SHL01.

  DATA:
     lv_bit TYPE i,
     lv_curr_pos  TYPE i VALUE 2,
     lv_prev_pos TYPE i VALUE 1.

  DO 15 TIMES.
    GET BIT lv_curr_pos OF i_pwd_hash INTO lv_bit.
    SET BIT lv_prev_pos OF r_pwd_hash TO lv_bit.
    ADD 1 TO lv_curr_pos.
    ADD 1 TO lv_prev_pos.
  ENDDO.
  SET BIT 16 OF r_pwd_hash TO 0.

endmethod.


method SHR14.

  DATA:
    lv_bit TYPE i,
    lv_curr_pos TYPE i,
    lv_next_pos TYPE i.

  r_pwd_hash = i_pwd_hash.

  DO 14 TIMES.
    lv_curr_pos = 15.
    lv_next_pos = 16.

    DO 15 TIMES.
      GET BIT lv_curr_pos OF r_pwd_hash INTO lv_bit.
      SET BIT lv_next_pos OF r_pwd_hash TO lv_bit.
      SUBTRACT 1 FROM lv_curr_pos.
      SUBTRACT 1 FROM lv_next_pos.
    ENDDO.
    SET BIT 1 OF r_pwd_hash TO 0.
  ENDDO.

endmethod.


method SPLIT_FILE.

  DATA: lt_hlp TYPE TABLE OF text255,
        ls_hlp TYPE text255.

  DATA: lf_ext(10)     TYPE c,
        lf_dot_ext(10) TYPE c.
  DATA: lf_anz TYPE i,
        lf_len TYPE i.
** ---------------------------------------------------------------------

  CLEAR: lt_hlp,
         ep_file,
         ep_extension,
         ep_dotextension.

** Split the whole file at '.'
  SPLIT ip_file AT '.' INTO TABLE lt_hlp.

** get the extenstion from the last line of table
  DESCRIBE TABLE lt_hlp LINES lf_anz.
  IF lf_anz <= 1.
    ep_file = ip_file.
    EXIT.
  ENDIF.

  READ TABLE lt_hlp INTO ls_hlp INDEX lf_anz.
  ep_extension = ls_hlp.
  lf_ext =  ls_hlp.
  IF NOT lf_ext IS INITIAL.
    CONCATENATE '.' lf_ext INTO lf_dot_ext.
  ENDIF.
  ep_dotextension = lf_dot_ext.

** get only the filename
  lf_len = strlen( ip_file ) - strlen( lf_dot_ext ).
  IF lf_len > 0.
    ep_file = ip_file(lf_len).
  ENDIF.

endmethod.


method STRUCTURE_CASE.
  DATA: lt_comp_str        TYPE abap_component_tab.

  CASE is_component-type->kind.
    WHEN cl_abap_typedescr=>kind_elem. "E Elementary Type
      INSERT is_component INTO TABLE xt_components.
    WHEN cl_abap_typedescr=>kind_table. "T Table
      INSERT is_component INTO TABLE xt_components.
    WHEN cl_abap_typedescr=>kind_struct. "S Structure
      lt_comp_str = structure_recursive( is_component = is_component ).
      INSERT LINES OF lt_comp_str INTO TABLE xt_components.
    WHEN OTHERS. "cl_abap_typedescr=>kind_ref or  cl_abap_typedescr=>kind_class or  cl_abap_typedescr=>kind_intf.
* We skip it. for now.
  ENDCASE.
endmethod.


method STRUCTURE_RECURSIVE.
  DATA: lo_struct          TYPE REF TO cl_abap_structdescr,
        lt_components      TYPE abap_component_tab,
        ls_components      TYPE abap_componentdescr.

  REFRESH rt_components.

  lo_struct ?= is_component-type.
  lt_components = lo_struct->get_components( ).

  LOOP AT lt_components INTO ls_components.
    structure_case( EXPORTING is_component  = ls_components
                    CHANGING  xt_components = rt_components ) .
  ENDLOOP.

endmethod.


method TIME_TO_EXCEL_STRING.
  DATA: lv_seconds_in_day TYPE i,
        lv_day_fraction TYPE f,
        lc_time_baseline TYPE t VALUE '000000',
        lc_seconds_in_day TYPE i VALUE 86400.

  lv_seconds_in_day = ip_value - lc_time_baseline.
  lv_day_fraction = lv_seconds_in_day / lc_seconds_in_day.
  ep_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_day_fraction ).
endmethod.


method UNESCAPE_STRING.

  CONSTANTS   lcv_regex                       TYPE string VALUE `^'[^']`    & `|` &  " Beginning single ' OR
                                                                `[^']'$`    & `|` &  " Trailing single '  OR
                                                                `[^']'[^']`.         " Single ' somewhere in between


  DATA:       lv_errormessage                 TYPE string.                          " Can't pass '...'(abc) to exception-class

*--------------------------------------------------------------------*
* This method is used to extract the "real" string from an escaped string.
* An escaped string can be identified by a beginning ' which must be
* accompanied by a trailing '
* All '' in between beginning and trailing ' are treated as single '
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* When allowing clike-input parameters we might encounter trailing
* "real" blanks .  These are automatically eliminated when moving
* the input parameter to a string.
*--------------------------------------------------------------------*
  ev_unescaped_string = iv_escaped.           " Pass through if not escaped

  CHECK ev_unescaped_string IS NOT INITIAL.   " Nothing to do if empty
  CHECK ev_unescaped_string(1) = `'`.         " Nothing to do if not escaped

*--------------------------------------------------------------------*
* Remove leading and trailing '
*--------------------------------------------------------------------*
  REPLACE REGEX `^'(.*)'$` IN ev_unescaped_string WITH '$1'.
  IF sy-subrc <> 0.
    lv_errormessage = 'Input not properly escaped - &'(002).
    zcx_excel=>raise_text( lv_errormessage ).
  ENDIF.

*--------------------------------------------------------------------*
* Any remaining single ' should not be here
*--------------------------------------------------------------------*
  FIND REGEX lcv_regex IN ev_unescaped_string.
  IF sy-subrc = 0.
    lv_errormessage = 'Input not properly escaped - &'(002).
    zcx_excel=>raise_text( lv_errormessage ).
  ENDIF.

*--------------------------------------------------------------------*
* Replace '' with '
*--------------------------------------------------------------------*
  REPLACE ALL OCCURRENCES OF `''` IN ev_unescaped_string WITH `'`.


endmethod.
ENDCLASS.
