CLASS zcl_excel_worksheets DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC
  INHERITING FROM zcl_excel_base.

*"* public components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PUBLIC SECTION.

    DATA active_worksheet TYPE zexcel_active_worksheet VALUE 1. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
    DATA name TYPE zexcel_worksheets_name VALUE 'Worksheets'. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .

    METHODS add
      IMPORTING
        !ip_worksheet TYPE REF TO zcl_excel_worksheet .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index           TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_worksheet TYPE REF TO zcl_excel_worksheet .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS clone REDEFINITION.
*"* protected components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA worksheets TYPE REF TO zcl_excel_collection .
ENDCLASS.



CLASS zcl_excel_worksheets IMPLEMENTATION.


  METHOD add.

    worksheets->add( ip_worksheet ).

  ENDMETHOD.


  METHOD clear.

    worksheets->clear( ).

  ENDMETHOD.


  METHOD constructor.
    super->constructor( ).
    CREATE OBJECT worksheets.

  ENDMETHOD.


  METHOD get.

    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_worksheet ?= worksheets->get( lv_index ).

  ENDMETHOD.


  METHOD get_iterator.

    eo_iterator ?= worksheets->get_iterator( ).

  ENDMETHOD.


  METHOD is_empty.

    is_empty = worksheets->is_empty( ).

  ENDMETHOD.


  METHOD remove.

    worksheets->remove( ip_worksheet ).

  ENDMETHOD.


  METHOD size.

    ep_size = worksheets->size( ).

  ENDMETHOD.


  METHOD clone.
    DATA lo_excel_worksheets TYPE REF TO zcl_excel_worksheets.

    CREATE OBJECT lo_excel_worksheets.

    lo_excel_worksheets->worksheets ?= worksheets->clone( ).

    lo_excel_worksheets->active_worksheet = active_worksheet.
    lo_excel_worksheets->name             = name.

    ro_object = lo_excel_worksheets.
  ENDMETHOD.
ENDCLASS.
