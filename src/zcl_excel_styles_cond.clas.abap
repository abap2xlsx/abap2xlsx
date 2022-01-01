CLASS zcl_excel_styles_cond DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLES_COND
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index            TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLES_COND
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA styles_cond TYPE REF TO zcl_excel_collection .
ENDCLASS.



CLASS zcl_excel_styles_cond IMPLEMENTATION.


  METHOD add.
    styles_cond->add( ip_style_cond ).
  ENDMETHOD.


  METHOD clear.
    styles_cond->clear( ).
  ENDMETHOD.


  METHOD constructor.

    CREATE OBJECT styles_cond.

  ENDMETHOD.


  METHOD get.
    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_style_cond ?= styles_cond->get( lv_index ).
  ENDMETHOD.


  METHOD get_iterator.
    eo_iterator ?= styles_cond->get_iterator( ).
  ENDMETHOD.


  METHOD is_empty.
    is_empty = styles_cond->is_empty( ).
  ENDMETHOD.


  METHOD remove.
    styles_cond->remove( ip_style_cond ).
  ENDMETHOD.


  METHOD size.
    ep_size = styles_cond->size( ).
  ENDMETHOD.
ENDCLASS.
