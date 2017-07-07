class ZCL_EXCEL_COLUMNS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_COLUMNS
*"* do not include other source files here!!!
public section.

  methods ADD
    importing
      !IO_COLUMN type ref to ZCL_EXCEL_COLUMN .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type I
    returning
      value(EO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IO_COLUMN type ref to ZCL_EXCEL_COLUMN .
  methods SIZE
    returning
      value(EP_SIZE) type I .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
protected section.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
private section.

  data COLUMNS type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_COLUMNS IMPLEMENTATION.


METHOD add.
  columns->add( io_column ).
ENDMETHOD.


METHOD clear.
  columns->clear( ).
ENDMETHOD.


METHOD constructor.

  CREATE OBJECT columns.

ENDMETHOD.


METHOD get.
  eo_column ?= columns->if_object_collection~get( ip_index ).
ENDMETHOD.


METHOD get_iterator.
  eo_iterator ?= columns->if_object_collection~get_iterator( ).
ENDMETHOD.


METHOD is_empty.
  is_empty = columns->if_object_collection~is_empty( ).
ENDMETHOD.


METHOD remove.
  columns->remove( io_column ).
ENDMETHOD.


METHOD size.
  ep_size = columns->if_object_collection~size( ).
ENDMETHOD.
ENDCLASS.
