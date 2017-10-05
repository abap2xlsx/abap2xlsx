class ZCL_EXCEL_DRAWINGS definition
  public
  final
  create public .

public section.

*"* public components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
  data TYPE type ZEXCEL_DRAWING_TYPE read-only value 'IMAGE'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .

  methods ADD
    importing
      !IP_DRAWING type ref to ZCL_EXCEL_DRAWING .
  methods INCLUDE
    importing
      !IP_DRAWING type ref to ZCL_EXCEL_DRAWING .
  methods CLEAR .
  methods CONSTRUCTOR
    importing
      !IP_TYPE type ZEXCEL_DRAWING_TYPE .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    returning
      value(EO_DRAWING) type ref to ZCL_EXCEL_DRAWING .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_DRAWING type ref to ZCL_EXCEL_DRAWING .
  methods SIZE
    returning
      value(EP_SIZE) type I .
  methods GET_TYPE
    returning
      value(RP_TYPE) type ZEXCEL_DRAWING_TYPE .
*"* protected components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
protected section.
private section.

*"* private components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
  data DRAWINGS type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_DRAWINGS IMPLEMENTATION.


method ADD.
  DATA: lv_index TYPE i.

  drawings->add( ip_drawing ).
  lv_index = drawings->if_object_collection~size( ).
  ip_drawing->create_media_name(
    ip_index = lv_index ).
  endmethod.


method CLEAR.

  drawings->clear( ).
  endmethod.


method CONSTRUCTOR.

  CREATE OBJECT drawings.
  type = ip_type.

  endmethod.


method GET.

  DATA lv_index TYPE i.
  lv_index = ip_index.
  eo_drawing ?= drawings->if_object_collection~get( lv_index ).
  endmethod.


method GET_ITERATOR.

  eo_iterator ?= drawings->if_object_collection~get_iterator( ).
  endmethod.


method GET_TYPE.
  rp_type = me->type.
  endmethod.


method INCLUDE.
  drawings->add( ip_drawing ).
  endmethod.


method IS_EMPTY.

  is_empty = drawings->if_object_collection~is_empty( ).
  endmethod.


method REMOVE.

  drawings->remove( ip_drawing ).
  endmethod.


method SIZE.

  ep_size = drawings->if_object_collection~size( ).
  endmethod.
ENDCLASS.
