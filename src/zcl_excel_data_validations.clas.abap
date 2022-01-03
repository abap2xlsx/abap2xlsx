CLASS zcl_excel_data_validations DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_data_validation TYPE REF TO zcl_excel_data_validation .
    METHODS clear .
    METHODS constructor .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_data_validation TYPE REF TO zcl_excel_data_validation .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA data_validations TYPE REF TO zcl_excel_collection .
ENDCLASS.



CLASS zcl_excel_data_validations IMPLEMENTATION.


  METHOD add.
    data_validations->add( ip_data_validation ).
  ENDMETHOD.


  METHOD clear.
    data_validations->clear( ).
  ENDMETHOD.


  METHOD constructor.

    CREATE OBJECT data_validations.

  ENDMETHOD.


  METHOD get_iterator.
    eo_iterator ?= data_validations->get_iterator( ).
  ENDMETHOD.


  METHOD is_empty.
    is_empty = data_validations->is_empty( ).
  ENDMETHOD.


  METHOD remove.
    data_validations->remove( ip_data_validation ).
  ENDMETHOD.


  METHOD size.
    ep_size = data_validations->size( ).
  ENDMETHOD.
ENDCLASS.
