class ZCL_EXCEL_RANGE definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_RANGE
*"* do not include other source files here!!!
public section.

  constants GCV_PRINT_TITLE_NAME type STRING value '_xlnm.Print_Titles'. "#EC NOTEXT
  data NAME type ZEXCEL_RANGE_NAME .
  data GUID type ZEXCEL_RANGE_GUID .

  methods GET_GUID
    returning
      value(EP_GUID) type ZEXCEL_RANGE_GUID .
  methods SET_VALUE
    importing
      !IP_SHEET_NAME type ZEXCEL_SHEET_TITLE
      !IP_START_ROW type ZEXCEL_CELL_ROW
      !IP_START_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA
      !IP_STOP_ROW type ZEXCEL_CELL_ROW
      !IP_STOP_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA .
  methods GET_VALUE
    returning
      value(EP_VALUE) type ZEXCEL_RANGE_VALUE .
  methods SET_RANGE_VALUE
    importing
      !IP_VALUE type ZEXCEL_RANGE_VALUE .
*"* protected components of class ZABAP_EXCEL_WORKSHEET
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_RANGE
*"* do not include other source files here!!!
private section.

  data VALUE type ZEXCEL_RANGE_VALUE .
ENDCLASS.



CLASS ZCL_EXCEL_RANGE IMPLEMENTATION.


method GET_GUID.

  ep_guid = me->guid.

  endmethod.


method GET_VALUE.

  ep_value = me->value.

  endmethod.


method SET_RANGE_VALUE.
  me->value = ip_value.
  endmethod.


method SET_VALUE.
  DATA: lv_start_row_c  TYPE char7,
        lv_stop_row_c   TYPE char7,
        lv_value        TYPE string.
  lv_stop_row_c = ip_stop_row.
  SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
  SHIFT lv_stop_row_c LEFT DELETING LEADING space.
  lv_start_row_c = ip_start_row.
  SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
  SHIFT lv_start_row_c LEFT DELETING LEADING space.
  lv_value = ip_sheet_name.
  me->value = zcl_excel_common=>escape_string( ip_value = lv_value ).

  CONCATENATE me->value '!$' ip_start_column '$' lv_start_row_c ':$' ip_stop_column '$' lv_stop_row_c INTO me->value.
  endmethod.
ENDCLASS.
