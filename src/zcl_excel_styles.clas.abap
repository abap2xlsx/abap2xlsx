CLASS zcl_excel_styles DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLES
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_style TYPE REF TO zcl_excel_style .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index       TYPE i
      RETURNING
        VALUE(eo_style) TYPE REF TO zcl_excel_style .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_style TYPE REF TO zcl_excel_style .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS register_new_style
      IMPORTING
        !io_style            TYPE REF TO zcl_excel_style
      RETURNING
        VALUE(ep_style_code) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA styles TYPE REF TO zcl_excel_collection .
ENDCLASS.



CLASS zcl_excel_styles IMPLEMENTATION.


  METHOD add.


    styles->add( ip_style ).
  ENDMETHOD.


  METHOD clear.


    styles->clear( ).
  ENDMETHOD.


  METHOD constructor.


    CREATE OBJECT styles.
  ENDMETHOD.


  METHOD get.


    eo_style ?= styles->get( ip_index ).
  ENDMETHOD.


  METHOD get_iterator.


    eo_iterator ?= styles->get_iterator( ).
  ENDMETHOD.


  METHOD is_empty.


    is_empty = styles->is_empty( ).
  ENDMETHOD.


  METHOD register_new_style.


    me->add( io_style ).
    ep_style_code = me->size( ) - 1. "style count starts from 0
  ENDMETHOD.


  METHOD remove.


    styles->remove( ip_style ).
  ENDMETHOD.


  METHOD size.


    ep_size = styles->size( ).
  ENDMETHOD.
ENDCLASS.
