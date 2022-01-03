*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_ROWS DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_excel_rows DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_ROWS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PUBLIC SECTION.
    METHODS add
      IMPORTING
        !io_row TYPE REF TO zcl_excel_row .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index     TYPE i
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !io_row TYPE REF TO zcl_excel_row .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_min_index
      RETURNING
        VALUE(ep_index) TYPE i .
    METHODS get_max_index
      RETURNING
        VALUE(ep_index) TYPE i .
  PROTECTED SECTION.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  PRIVATE SECTION.
    TYPES:
      BEGIN OF mty_s_hashed_row,
        row_index TYPE int4,
        row       TYPE REF TO zcl_excel_row,
      END OF mty_s_hashed_row ,
      mty_ts_hashed_row TYPE HASHED TABLE OF mty_s_hashed_row WITH UNIQUE KEY row_index.

    DATA rows TYPE REF TO zcl_excel_collection .
    DATA rows_hashed TYPE mty_ts_hashed_row .

ENDCLASS.



CLASS zcl_excel_rows IMPLEMENTATION.


  METHOD add.
    DATA: ls_hashed_row TYPE mty_s_hashed_row.

    ls_hashed_row-row_index = io_row->get_row_index( ).
    ls_hashed_row-row = io_row.

    INSERT ls_hashed_row INTO TABLE rows_hashed.

    rows->add( io_row ).
  ENDMETHOD.                    "ADD


  METHOD clear.
    CLEAR rows_hashed.
    rows->clear( ).
  ENDMETHOD.                    "CLEAR


  METHOD constructor.

    CREATE OBJECT rows.

  ENDMETHOD.                    "CONSTRUCTOR


  METHOD get.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    READ TABLE rows_hashed WITH KEY row_index = ip_index ASSIGNING <ls_hashed_row>.
    IF sy-subrc = 0.
      eo_row = <ls_hashed_row>-row.
    ENDIF.
  ENDMETHOD.                    "GET


  METHOD get_iterator.
    eo_iterator ?= rows->get_iterator( ).
  ENDMETHOD.                    "GET_ITERATOR


  METHOD get_max_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hashed ASSIGNING <ls_hashed_row>.
      IF <ls_hashed_row>-row_index > ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_min_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hashed ASSIGNING <ls_hashed_row>.
      IF ep_index = 0 OR <ls_hashed_row>-row_index < ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD is_empty.
    is_empty = rows->is_empty( ).
  ENDMETHOD.                    "IS_EMPTY


  METHOD remove.
    DELETE TABLE rows_hashed WITH TABLE KEY row_index = io_row->get_row_index( ) .
    rows->remove( io_row ).
  ENDMETHOD.                    "REMOVE


  METHOD size.
    ep_size = rows->size( ).
  ENDMETHOD.                    "SIZE
ENDCLASS.
