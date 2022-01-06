CLASS zcl_excel_columns DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_COLUMNS
*"* do not include other source files here!!!
  PUBLIC SECTION.
    METHODS add
      IMPORTING
        !io_column TYPE REF TO zcl_excel_column .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index        TYPE i
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !io_column TYPE REF TO zcl_excel_column .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  PRIVATE SECTION.
    TYPES:
      BEGIN OF mty_s_hashed_column,
        column_index TYPE int4,
        column       TYPE REF TO zcl_excel_column,
      END OF mty_s_hashed_column ,
      mty_ts_hashed_column TYPE HASHED TABLE OF mty_s_hashed_column WITH UNIQUE KEY column_index.

    DATA columns TYPE REF TO zcl_excel_collection .
    DATA columns_hashed TYPE mty_ts_hashed_column .
ENDCLASS.



CLASS zcl_excel_columns IMPLEMENTATION.


  METHOD add.
    DATA: ls_hashed_column TYPE mty_s_hashed_column.

    ls_hashed_column-column_index = io_column->get_column_index( ).
    ls_hashed_column-column = io_column.

    INSERT ls_hashed_column INTO TABLE columns_hashed .

    columns->add( io_column ).
  ENDMETHOD.


  METHOD clear.
    CLEAR columns_hashed.
    columns->clear( ).
  ENDMETHOD.


  METHOD constructor.

    CREATE OBJECT columns.

  ENDMETHOD.


  METHOD get.
    FIELD-SYMBOLS: <ls_hashed_column> TYPE mty_s_hashed_column.

    READ TABLE columns_hashed WITH KEY column_index = ip_index ASSIGNING <ls_hashed_column>.
    IF sy-subrc = 0.
      eo_column = <ls_hashed_column>-column.
    ENDIF.
  ENDMETHOD.


  METHOD get_iterator.
    eo_iterator ?= columns->get_iterator( ).
  ENDMETHOD.


  METHOD is_empty.
    is_empty = columns->is_empty( ).
  ENDMETHOD.


  METHOD remove.
    DELETE TABLE columns_hashed WITH TABLE KEY column_index = io_column->get_column_index( ) .
    columns->remove( io_column ).
  ENDMETHOD.


  METHOD size.
    ep_size = columns->size( ).
  ENDMETHOD.
ENDCLASS.
