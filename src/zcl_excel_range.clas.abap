CLASS zcl_excel_range DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC
  INHERITING FROM zcl_excel_base.

*"* public components of class ZCL_EXCEL_RANGE
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS gcv_print_title_name TYPE string VALUE '_xlnm.Print_Titles'. "#EC NOTEXT

    TYPES:
      BEGIN OF ts_sheet_title,
        title  TYPE zexcel_sheet_title,
        offset TYPE i,
        length TYPE i,
      END OF ts_sheet_title.

    DATA name TYPE zexcel_range_name .
    DATA guid TYPE zexcel_range_guid .

    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE zexcel_range_guid .
    METHODS set_value
      IMPORTING
        !ip_sheet_name   TYPE zexcel_sheet_title
        !ip_start_row    TYPE zexcel_cell_row
        !ip_start_column TYPE zexcel_cell_column_alpha
        !ip_stop_row     TYPE zexcel_cell_row
        !ip_stop_column  TYPE zexcel_cell_column_alpha .
    METHODS get_value
      RETURNING
        VALUE(ep_value) TYPE zexcel_range_value .
    METHODS set_range_value
      IMPORTING
        ip_value TYPE zexcel_range_value .
    METHODS clone REDEFINITION.
    METHODS get_sheet_title
      RETURNING
        VALUE(rs_sheet_title) TYPE ts_sheet_title
      RAISING
        zcx_excel.
    METHODS replace_sheet_title
      IMPORTING
        !iv_new_sheet_title TYPE zexcel_sheet_title
      RAISING
        zcx_excel.
*"* protected components of class ZABAP_EXCEL_WORKSHEET
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_RANGE
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA value TYPE zexcel_range_value.
    DATA ms_sheet_title TYPE ts_sheet_title.
ENDCLASS.



CLASS zcl_excel_range IMPLEMENTATION.


  METHOD get_guid.

    ep_guid = me->guid.

  ENDMETHOD.


  METHOD get_value.

    ep_value = me->value.

  ENDMETHOD.


  METHOD set_range_value.
    me->value = ip_value.
  ENDMETHOD.


  METHOD set_value.
    DATA: lv_start_row_c TYPE c LENGTH 7,
          lv_stop_row_c  TYPE c LENGTH 7,
          lv_value       TYPE string,
          ls_sheet_title TYPE ts_sheet_title.
    lv_stop_row_c = ip_stop_row.
    SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
    SHIFT lv_stop_row_c LEFT DELETING LEADING space.
    lv_start_row_c = ip_start_row.
    SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
    SHIFT lv_start_row_c LEFT DELETING LEADING space.

    ls_sheet_title-title  = ip_sheet_name.
    ls_sheet_title-offset = 0.
    ls_sheet_title-length = strlen( ip_sheet_name ).

    me->value = zcl_excel_common=>escape_string( ip_value = lv_value ).

    IF ip_stop_column IS INITIAL AND ip_stop_row IS INITIAL.
      CONCATENATE ls_sheet_title-title '!$' ip_start_column '$' lv_start_row_c INTO me->value.
    ELSE.
      CONCATENATE ls_sheet_title-title '!$' ip_start_column '$' lv_start_row_c ':$' ip_stop_column '$' lv_stop_row_c INTO me->value.
    ENDIF.

    ms_sheet_title = ls_sheet_title.
  ENDMETHOD.


  METHOD clone.
    DATA lo_excel_range TYPE REF TO zcl_excel_range.

    CREATE OBJECT lo_excel_range.

    lo_excel_range->name  = name.
    lo_excel_range->value = value.

    ro_object = lo_excel_range.
  ENDMETHOD.

  METHOD get_sheet_title.
    CONSTANTS lc_title_submatch_index TYPE i VALUE 1.

    DATA ls_sheet_title TYPE ts_sheet_title.

    IF ms_sheet_title IS NOT INITIAL.
      rs_sheet_title = ms_sheet_title.
      RETURN.
    ENDIF.

    DATA(lo_regex) =
      NEW cl_abap_regex( pattern = `^([^\\\[\]\*\?\:]{1,31})(?:\!\$)(?:[a-zA-Z0-9]+|[a-zA-Z0-9]+\:\$[a-zA-Z0-9]+)$` ).

    DATA(lo_matcher) = lo_regex->create_matcher( text = value ).
    DATA(lv_matches) = lo_matcher->match( ).

    IF lv_matches = abap_false.
      zcx_excel=>raise_text( `Failed to match the regular expression.` ).
    ENDIF.

    ls_sheet_title-title   = lo_matcher->get_submatch( lc_title_submatch_index ).
    ls_sheet_title-offset  = lo_matcher->get_offset( lc_title_submatch_index ).
    ls_sheet_title-length  = lo_matcher->get_length( lc_title_submatch_index ).

    ms_sheet_title = ls_sheet_title.
    rs_sheet_title = ls_sheet_title.
  ENDMETHOD.

  METHOD replace_sheet_title.
    DATA(ls_sheet_title) = get_sheet_title( ).
    REPLACE SECTION OFFSET ls_sheet_title-offset LENGTH ls_sheet_title-length OF value WITH iv_new_sheet_title.

    ms_sheet_title-title  = iv_new_sheet_title.
    ms_sheet_title-offset = 0.
    ms_sheet_title-length = strlen( iv_new_sheet_title ).
  ENDMETHOD.

ENDCLASS.
