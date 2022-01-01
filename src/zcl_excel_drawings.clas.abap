CLASS zcl_excel_drawings DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
    DATA type TYPE zexcel_drawing_type READ-ONLY VALUE 'IMAGE'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .

    METHODS add
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS include
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS clear .
    METHODS constructor
      IMPORTING
        !ip_type TYPE zexcel_drawing_type .
    METHODS get
      IMPORTING
        !ip_index         TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_drawing) TYPE REF TO zcl_excel_drawing .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_type
      RETURNING
        VALUE(rp_type) TYPE zexcel_drawing_type .
*"* protected components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
    DATA drawings TYPE REF TO zcl_excel_collection .
ENDCLASS.



CLASS zcl_excel_drawings IMPLEMENTATION.


  METHOD add.
    DATA: lv_index TYPE i.

    drawings->add( ip_drawing ).
    lv_index = drawings->size( ).
    ip_drawing->create_media_name(
      ip_index = lv_index ).
  ENDMETHOD.


  METHOD clear.

    drawings->clear( ).
  ENDMETHOD.


  METHOD constructor.

    CREATE OBJECT drawings.
    type = ip_type.

  ENDMETHOD.


  METHOD get.

    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_drawing ?= drawings->get( lv_index ).
  ENDMETHOD.


  METHOD get_iterator.

    eo_iterator ?= drawings->get_iterator( ).
  ENDMETHOD.


  METHOD get_type.
    rp_type = me->type.
  ENDMETHOD.


  METHOD include.
    drawings->add( ip_drawing ).
  ENDMETHOD.


  METHOD is_empty.

    is_empty = drawings->is_empty( ).
  ENDMETHOD.


  METHOD remove.

    drawings->remove( ip_drawing ).
  ENDMETHOD.


  METHOD size.

    ep_size = drawings->size( ).
  ENDMETHOD.
ENDCLASS.
