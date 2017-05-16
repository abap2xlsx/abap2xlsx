class ZCL_EXCEL_WORKSHEETS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
public section.

  data ACTIVE_WORKSHEET type ZEXCEL_ACTIVE_WORKSHEET value 1. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
  data NAME type ZEXCEL_WORKSHEETS_NAME value 'Worksheets'. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .

  methods ADD
    importing
      !IP_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    returning
      value(EO_WORKSHEET) type ref to ZCL_EXCEL_WORKSHEET .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
  methods SIZE
    returning
      value(EP_SIZE) type I .
*"* protected components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
private section.

  data WORKSHEETS type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_WORKSHEETS IMPLEMENTATION.


method ADD.

  worksheets->add( ip_worksheet ).

  endmethod.


method CLEAR.

  worksheets->clear( ).

  endmethod.


method CONSTRUCTOR.

  CREATE OBJECT worksheets.

  endmethod.


method GET.

  DATA lv_index TYPE i.
  lv_index = ip_index.
  eo_worksheet ?= worksheets->if_object_collection~get( lv_index ).

  endmethod.


method GET_ITERATOR.

  eo_iterator ?= worksheets->if_object_collection~get_iterator( ).

  endmethod.


method IS_EMPTY.

  is_empty = worksheets->if_object_collection~is_empty( ).

  endmethod.


method REMOVE.

  worksheets->remove( ip_worksheet ).

  endmethod.


method SIZE.

  ep_size = worksheets->if_object_collection~size( ).

  endmethod.
ENDCLASS.
