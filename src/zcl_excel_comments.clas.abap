class ZCL_EXCEL_COMMENTS definition
  public
  final
  create public .

public section.

  methods ADD
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_COMMENT .
  methods INCLUDE
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_COMMENT .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    returning
      value(EO_COMMENT) type ref to ZCL_EXCEL_COMMENT .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_COMMENT .
  methods SIZE
    returning
      value(EP_SIZE) type I .
protected section.
private section.

  data COMMENTS type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_COMMENTS IMPLEMENTATION.


METHOD add.
  DATA: lv_index TYPE i.

  comments->add( ip_comment ).
  lv_index = comments->size( ).

ENDMETHOD.


METHOD clear.
  comments->clear( ).

ENDMETHOD.


METHOD constructor.
  CREATE OBJECT comments.

ENDMETHOD.


method GET.
  DATA lv_index TYPE i.
  lv_index = ip_index.
  eo_comment ?= comments->get( lv_index ).

endmethod.


method GET_ITERATOR.

  eo_iterator ?= comments->get_iterator( ).
  endmethod.


METHOD include.
  comments->add( ip_comment ).
ENDMETHOD.


method IS_EMPTY.

  is_empty = comments->is_empty( ).
  endmethod.


method REMOVE.

  comments->remove( ip_comment ).
  endmethod.


method SIZE.

  ep_size = comments->size( ).
  endmethod.
ENDCLASS.
