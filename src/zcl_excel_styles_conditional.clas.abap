class ZCL_EXCEL_STYLES_CONDITIONAL definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLES_CONDITIONAL
*"* do not include other source files here!!!
public section.

  methods ADD
    importing
      !IP_STYLE_CONDITIONAL type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    returning
      value(EO_STYLE_CONDITIONAL) type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_STYLE_CONDITIONAL type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
  methods SIZE
    returning
      value(EP_SIZE) type I .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_STYLES_CONDITIONAL
*"* do not include other source files here!!!
private section.

  data STYLES_CONDITIONAL type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_STYLES_CONDITIONAL IMPLEMENTATION.


method ADD.
  styles_conditional->add( ip_style_conditional ).
  endmethod.


method CLEAR.
  styles_conditional->clear( ).
  endmethod.


method CONSTRUCTOR.

  CREATE OBJECT styles_conditional.

  endmethod.


method GET.
  DATA lv_index TYPE i.
  lv_index = ip_index.
  eo_style_conditional ?= styles_conditional->if_object_collection~get( lv_index ).
  endmethod.


method GET_ITERATOR.
  eo_iterator ?= styles_conditional->if_object_collection~get_iterator( ).
  endmethod.


method IS_EMPTY.
  is_empty = styles_conditional->if_object_collection~is_empty( ).
  endmethod.


method REMOVE.
  styles_conditional->remove( ip_style_conditional ).
  endmethod.


method SIZE.
  ep_size = styles_conditional->if_object_collection~size( ).
  endmethod.
ENDCLASS.
