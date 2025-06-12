CLASS zcl_excel_comments DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      ty_boxes TYPE STANDARD TABLE OF zcl_excel_comment=>ty_box
          WITH NON-UNIQUE DEFAULT KEY .

    DATA gv_full_vml TYPE string READ-ONLY .

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
    METHODS set_boxes
      IMPORTING
        !it_boxes    TYPE ty_boxes OPTIONAL
        !iv_full_vml TYPE string OPTIONAL .
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA comments TYPE REF TO zcl_excel_collection .
    DATA gt_boxes TYPE ty_boxes .
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
* Copy constructor: copy attributes from original
* So the receiver may change the collection without affecting the original
      lo_iterator = io_from->comments->get_iterator( ).
      WHILE lo_iterator->has_next( ) = abap_true.
        lo_comment ?= lo_iterator->get_next( ).
        include( lo_comment ).
      ENDWHILE.
      gt_boxes    = io_from->gt_boxes.
      gv_full_vml = io_from->gv_full_vml.
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


  METHOD set_boxes.

    DATA:
      lo_comments TYPE REF TO zcl_excel_collection_iterator,
      lo_comment  TYPE REF TO zcl_excel_comment.

    FIELD-SYMBOLS:
      <ls_box> TYPE zcl_excel_comment=>ty_box.

    IF it_boxes IS NOT INITIAL.
      gt_boxes = it_boxes.
    ENDIF.

    IF iv_full_vml IS NOT INITIAL.
      gv_full_vml = iv_full_vml.
    ENDIF.

    IF gt_boxes IS NOT INITIAL.

      lo_comments = comments->get_iterator( ).
      WHILE lo_comments->has_next( ) EQ abap_true.
        READ TABLE gt_boxes INDEX 1 ASSIGNING <ls_box>.
        CHECK sy-subrc EQ 0.
        lo_comment ?= lo_comments->get_next( ).
        lo_comment->set_box( <ls_box> ).
        DELETE gt_boxes INDEX 1.
      ENDWHILE.

    ENDIF.

  ENDMETHOD.


ENDCLASS.
