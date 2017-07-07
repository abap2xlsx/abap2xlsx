class ZCL_EXCEL_STYLES_COND definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLES_COND
*"* do not include other source files here!!!
public section.

  methods ADD
    importing
      !IP_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    returning
      value(EO_STYLE_COND) type ref to ZCL_EXCEL_STYLE_COND .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods SIZE
    returning
      value(EP_SIZE) type I .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_STYLES_COND
*"* do not include other source files here!!!
private section.

  data STYLES_COND type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_STYLES_COND IMPLEMENTATION.


METHOD ADD.
  styles_cond->add( ip_style_cond ).
ENDMETHOD.


METHOD CLEAR.
  styles_cond->clear( ).
ENDMETHOD.


METHOD constructor.

  CREATE OBJECT styles_cond.

ENDMETHOD.


METHOD get.
  DATA lv_index TYPE i.
  lv_index = ip_index.
  eo_style_cond ?= styles_cond->if_object_collection~get( lv_index ).
ENDMETHOD.


METHOD get_iterator.
  eo_iterator ?= styles_cond->if_object_collection~get_iterator( ).
ENDMETHOD.


METHOD is_empty.
  is_empty = styles_cond->if_object_collection~is_empty( ).
ENDMETHOD.


METHOD remove.
  styles_cond->remove( ip_style_cond ).
ENDMETHOD.


METHOD size.
  ep_size = styles_cond->if_object_collection~size( ).
ENDMETHOD.
ENDCLASS.
