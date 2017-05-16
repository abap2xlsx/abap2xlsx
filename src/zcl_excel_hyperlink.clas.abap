class ZCL_EXCEL_HYPERLINK definition
  public
  final
  create private .

*"* public components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!
public section.
  type-pools ABAP .

  class-methods CREATE_EXTERNAL_LINK
    importing
      !IV_URL type STRING
    returning
      value(OV_LINK) type ref to ZCL_EXCEL_HYPERLINK .
  class-methods CREATE_INTERNAL_LINK
    importing
      !IV_LOCATION type STRING
    returning
      value(OV_LINK) type ref to ZCL_EXCEL_HYPERLINK .
  methods IS_INTERNAL
    returning
      value(EV_RET) type ABAP_BOOL .
  methods SET_CELL_REFERENCE
    importing
      !IP_COLUMN type SIMPLE
      !IP_ROW type ZEXCEL_CELL_ROW
    raising
      ZCX_EXCEL .
  methods GET_REF
    returning
      value(EV_REF) type STRING .
  methods GET_URL
    returning
      value(EV_URL) type STRING .
*"* protected components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!
private section.

  data LOCATION type STRING .
  data CELL_REFERENCE type STRING .
  data INTERNAL type ABAP_BOOL .
  data COLUMN type ZEXCEL_CELL_COLUMN_ALPHA .
  data ROW type ZEXCEL_CELL_ROW .

  class-methods CREATE
    importing
      !IV_URL type STRING
      !IV_INTERNAL type ABAP_BOOL
    returning
      value(OV_LINK) type ref to ZCL_EXCEL_HYPERLINK .
ENDCLASS.



CLASS ZCL_EXCEL_HYPERLINK IMPLEMENTATION.


method CREATE.
  data: lo_hyperlink type REF TO zcl_excel_hyperlink.

  create OBJECT lo_hyperlink.

  lo_hyperlink->location = iv_url.
  lo_hyperlink->internal = iv_internal.

  ov_link = lo_hyperlink.
  endmethod.


method CREATE_EXTERNAL_LINK.

  ov_link = zcl_excel_hyperlink=>create( iv_url = iv_url
                                         iv_internal = abap_false ).
  endmethod.


method CREATE_INTERNAL_LINK.
  ov_link = zcl_excel_hyperlink=>create( iv_url = iv_location
                                         iv_internal = abap_true ).
  endmethod.


method GET_REF.
  ev_ref = row.
  CONDENSE ev_ref.
  CONCATENATE column ev_ref INTO ev_ref.
  endmethod.


method GET_URL.
  ev_url = me->location.
  endmethod.


method IS_INTERNAL.
  ev_ret = me->internal.
  endmethod.


method SET_CELL_REFERENCE.
  me->column = zcl_excel_common=>convert_column2alpha( ip_column ). " issue #155 - less restrictive typing for ip_column
  me->row = ip_row.
  endmethod.
ENDCLASS.
