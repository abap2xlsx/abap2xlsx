class ZCL_EXCEL_COMMENTS definition
  public
  final
  create public .

public section.

  methods CONSTRUCTOR .
  methods ADD
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_COMMENT .
  methods CLEAR .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    exporting
      !EO_COMMENT type ref to ZCL_EXCEL_COMMENT .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods INCLUDE
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_COMMENT .
  methods IS_EMPTY
    exporting
      !IS_EMPTY type FLAG .
  methods REMOVE
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_COMMENT .
  methods SIZE
    exporting
      !EP_SIZE type I .
protected section.
private section.

  data COMMENTS type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_COMMENTS IMPLEMENTATION.


  method ADD.
      DATA: lv_index TYPE i.
   comments->add( ip_comment ).
  lv_index = comments->if_object_collection~size( ).
  endmethod.


  method CLEAR.
    comments->clear( ).
  endmethod.


  method CONSTRUCTOR.
     CREATE OBJECT comments.
  endmethod.


  method GET.

      DATA lv_index TYPE i.
  lv_index = ip_index .
  eo_comment ?= comments->if_object_collection~get( lv_index ).
  endmethod.


  method GET_ITERATOR.
     eo_iterator ?= comments->if_object_collection~get_iterator( ).
  endmethod.


  method INCLUDE.
    comments->add( ip_comment ).
  endmethod.


  method IS_EMPTY.
    is_empty = comments->if_object_collection~is_empty( ).

  endmethod.


  method REMOVE.
    comments->remove( ip_comment ).
  endmethod.


  method SIZE.
    ep_size = comments->if_object_collection~size( ).
  endmethod.
ENDCLASS.
