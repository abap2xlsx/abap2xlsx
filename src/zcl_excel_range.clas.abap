CLASS zcl_excel_range DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_RANGE
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS gcv_print_title_name TYPE string VALUE '_xlnm.Print_Titles'. "#EC NOTEXT
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
        !ip_value TYPE zexcel_range_value .
*"* protected components of class ZABAP_EXCEL_WORKSHEET
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_RANGE
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA value TYPE zexcel_range_value .
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
          lv_value       TYPE string.
    lv_stop_row_c = ip_stop_row.
    SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
    SHIFT lv_stop_row_c LEFT DELETING LEADING space.
    lv_start_row_c = ip_start_row.
    SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
    SHIFT lv_start_row_c LEFT DELETING LEADING space.
    lv_value = ip_sheet_name.
    me->value = zcl_excel_common=>escape_string( ip_value = lv_value ).

    IF ip_stop_column IS INITIAL AND ip_stop_row IS INITIAL.
      CONCATENATE me->value '!$' ip_start_column '$' lv_start_row_c INTO me->value.
    ELSE.
      CONCATENATE me->value '!$' ip_start_column '$' lv_start_row_c ':$' ip_stop_column '$' lv_stop_row_c INTO me->value.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
