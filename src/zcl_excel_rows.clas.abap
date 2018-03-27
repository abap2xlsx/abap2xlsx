*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_ROWS DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class ZCL_EXCEL_ROWS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_ROWS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
public section.

  methods ADD
    importing
      !IO_ROW type ref to ZCL_EXCEL_ROW .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type I
    returning
      value(EO_ROW) type ref to ZCL_EXCEL_ROW .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IO_ROW type ref to ZCL_EXCEL_ROW .
  methods SIZE
    returning
      value(EP_SIZE) type I .
  PROTECTED SECTION.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA rows TYPE REF TO cl_object_collection .
ENDCLASS.



CLASS ZCL_EXCEL_ROWS IMPLEMENTATION.


  METHOD add.
    rows->add( io_row ).
  ENDMETHOD.                    "ADD


  METHOD clear.
    rows->clear( ).
  ENDMETHOD.                    "CLEAR


  METHOD constructor.

    CREATE OBJECT rows.

  ENDMETHOD.                    "CONSTRUCTOR


  METHOD get.
    eo_row ?= rows->if_object_collection~get( ip_index ).
  ENDMETHOD.                    "GET


  METHOD get_iterator.
    eo_iterator ?= rows->if_object_collection~get_iterator( ).
  ENDMETHOD.                    "GET_ITERATOR


  METHOD is_empty.
    is_empty = rows->if_object_collection~is_empty( ).
  ENDMETHOD.                    "IS_EMPTY


  METHOD remove.
    rows->remove( io_row ).
  ENDMETHOD.                    "REMOVE


  METHOD size.
    ep_size = rows->if_object_collection~size( ).
  ENDMETHOD.                    "SIZE
ENDCLASS.
