CLASS zcl_excel_ranges DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_RANGES
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_range TYPE REF TO zcl_excel_range .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index       TYPE i
      RETURNING
        VALUE(eo_range) TYPE REF TO zcl_excel_range .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_range TYPE REF TO zcl_excel_range .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA ranges TYPE REF TO zcl_excel_collection .
ENDCLASS.



CLASS zcl_excel_ranges IMPLEMENTATION.


  METHOD add.
    ranges->add( ip_range ).
  ENDMETHOD.


  METHOD clear.
    ranges->clear( ).
  ENDMETHOD.


  METHOD constructor.


    CREATE OBJECT ranges.

  ENDMETHOD.


  METHOD get.
    eo_range ?= ranges->get( ip_index ).
  ENDMETHOD.


  METHOD get_iterator.
    eo_iterator ?= ranges->get_iterator( ).
  ENDMETHOD.


  METHOD is_empty.
    is_empty = ranges->is_empty( ).
  ENDMETHOD.


  METHOD remove.
    ranges->remove( ip_range ).
  ENDMETHOD.


  METHOD size.
    ep_size = ranges->size( ).
  ENDMETHOD.
ENDCLASS.
