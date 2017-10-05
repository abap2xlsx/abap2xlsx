class ZCL_EXCEL_STYLES definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLES
*"* do not include other source files here!!!
public section.

  methods ADD
    importing
      !IP_STYLE type ref to ZCL_EXCEL_STYLE .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type I
    returning
      value(EO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_STYLE type ref to ZCL_EXCEL_STYLE .
  methods SIZE
    returning
      value(EP_SIZE) type I .
  methods REGISTER_NEW_STYLE
    importing
      !IO_STYLE type ref to ZCL_EXCEL_STYLE
    returning
      value(EP_STYLE_CODE) type I .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
protected section.
private section.

  data STYLES type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_STYLES IMPLEMENTATION.


method ADD.


  styles->add( ip_style ).
  endmethod.


method CLEAR.


  styles->clear( ).
  endmethod.


method CONSTRUCTOR.


  CREATE OBJECT styles.
  endmethod.


method GET.


  eo_style ?= styles->if_object_collection~get( ip_index ).
  endmethod.


method GET_ITERATOR.


  eo_iterator ?= styles->if_object_collection~get_iterator( ).
  endmethod.


method IS_EMPTY.


  is_empty = styles->if_object_collection~is_empty( ).
  endmethod.


method REGISTER_NEW_STYLE.


  me->add( io_style ).
  ep_style_code = me->size( ) - 1. "style count starts from 0
  endmethod.


method REMOVE.


  styles->remove( ip_style ).
  endmethod.


method SIZE.


  ep_size = styles->if_object_collection~size( ).
  endmethod.
ENDCLASS.
