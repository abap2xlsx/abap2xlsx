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
    METHODS constructor
      IMPORTING
        !io_from TYPE REF TO zcl_excel_comments OPTIONAL .
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

    DATA: lo_comment  TYPE REF TO zcl_excel_comment,
          lo_iterator TYPE REF TO zcl_excel_collection_iterator.

    CREATE OBJECT comments.

    IF io_from IS BOUND.
* Copy constructor: create new instance with copy of attributes from io_from
* Copy all attributes of io_from to the new instance

* The receiver may change the collection without affecting the original
      lo_iterator = io_from->comments->get_iterator( ).
      WHILE lo_iterator->has_next( ) = abap_true.
        lo_comment ?= lo_iterator->get_next( ).
        include( lo_comment ).
      ENDWHILE.
    ENDIF.

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
