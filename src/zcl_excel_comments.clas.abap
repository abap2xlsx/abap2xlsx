CLASS zcl_excel_comments DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS include
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index         TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_comment) TYPE REF TO zcl_excel_comment .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA comments TYPE REF TO zcl_excel_collection .
ENDCLASS.



CLASS zcl_excel_comments IMPLEMENTATION.


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


  METHOD get.
    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_comment ?= comments->get( lv_index ).

  ENDMETHOD.


  METHOD get_iterator.

    eo_iterator ?= comments->get_iterator( ).
  ENDMETHOD.


  METHOD include.
    comments->add( ip_comment ).
  ENDMETHOD.


  METHOD is_empty.

    is_empty = comments->is_empty( ).
  ENDMETHOD.


  METHOD remove.

    comments->remove( ip_comment ).
  ENDMETHOD.


  METHOD size.

    ep_size = comments->size( ).
  ENDMETHOD.
ENDCLASS.
