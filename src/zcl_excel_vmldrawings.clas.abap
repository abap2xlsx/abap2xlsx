class ZCL_EXCEL_VMLDRAWINGS definition
  public
  final
  create public .

public section.

  methods CONSTRUCTOR .
  methods ADD
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_VMLDRAWING .
  methods CLEAR .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    exporting
      !EO_COMMENT type ref to ZCL_EXCEL_VMLDRAWING .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods INCLUDE
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_VMLDRAWING .
  methods IS_EMPTY
    exporting
      !IS_EMPTY type FLAG .
  methods REMOVE
    importing
      !IP_COMMENT type ref to ZCL_EXCEL_VMLDRAWING .
  methods SIZE
    returning
      value(EP_SIZE) type I .
protected section.
private section.

  data vmldrawings type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_VMLDRAWINGS IMPLEMENTATION.


  method ADD.
      DATA: lv_index TYPE i.
   vmldrawings->add( ip_comment ).
  lv_index = vmldrawings->if_object_collection~size( ).
  endmethod.


  method CLEAR.
    vmldrawings->clear( ).
  endmethod.


  method CONSTRUCTOR.
     CREATE OBJECT vmldrawings.
  endmethod.


  method GET.

      DATA lv_index TYPE i.
  lv_index = ip_index .
  eo_comment ?= vmldrawings->if_object_collection~get( lv_index ).
  endmethod.


  method GET_ITERATOR.
     eo_iterator ?= vmldrawings->if_object_collection~get_iterator( ).
  endmethod.


  method INCLUDE.
    vmldrawings->add( ip_comment ).
  endmethod.


  method IS_EMPTY.
    is_empty = vmldrawings->if_object_collection~is_empty( ).

  endmethod.


  method REMOVE.
    vmldrawings->remove( ip_comment ).
  endmethod.


  method SIZE.
    ep_size = vmldrawings->if_object_collection~size( ).
  endmethod.
ENDCLASS.
