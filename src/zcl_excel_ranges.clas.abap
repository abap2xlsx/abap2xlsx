class ZCL_EXCEL_RANGES definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_RANGES
*"* do not include other source files here!!!
public section.

  methods ADD
    importing
      !IP_RANGE type ref to ZCL_EXCEL_RANGE .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type I
    returning
      value(EO_RANGE) type ref to ZCL_EXCEL_RANGE .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_RANGE type ref to ZCL_EXCEL_RANGE .
  methods SIZE
    returning
      value(EP_SIZE) type I .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
protected section.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
private section.

  data RANGES type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_RANGES IMPLEMENTATION.


method ADD.
  ranges->add( ip_range ).
  endmethod.


method CLEAR.
  ranges->clear( ).
  endmethod.


method CONSTRUCTOR.


  CREATE OBJECT ranges.

  endmethod.


method GET.
  eo_range ?= ranges->if_object_collection~get( ip_index ).
  endmethod.


method GET_ITERATOR.
  eo_iterator ?= ranges->if_object_collection~get_iterator( ).
  endmethod.


method IS_EMPTY.
  is_empty = ranges->if_object_collection~is_empty( ).
  endmethod.


method REMOVE.
  ranges->remove( ip_range ).
  endmethod.


method SIZE.
  ep_size = ranges->if_object_collection~size( ).
  endmethod.
ENDCLASS.
