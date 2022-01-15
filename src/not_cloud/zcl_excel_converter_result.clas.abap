CLASS zcl_excel_converter_result DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_converter_alv
  ABSTRACT
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
  PUBLIC SECTION.
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
  PROTECTED SECTION.

    METHODS get_table
      IMPORTING
        !io_object     TYPE REF TO object
      RETURNING
        VALUE(ro_data) TYPE REF TO data .
*"* private components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
*"* private components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_converter_result IMPLEMENTATION.


  METHOD get_table.
    DATA: lo_object   TYPE REF TO object,
          ls_seoclass TYPE seoclass,
          l_method    TYPE string.

    SELECT SINGLE * INTO ls_seoclass
      FROM seoclass
      WHERE clsname = 'IF_SALV_BS_DATA_SOURCE'.

    IF sy-subrc = 0.
      l_method = 'GET_TABLE_REF'.
      lo_object ?= io_object.
      CALL METHOD lo_object->(l_method)
        RECEIVING
          value = ro_data.
    ELSE.
      l_method = 'GET_REF_TO_TABLE'.
      lo_object ?= io_object.
      CALL METHOD lo_object->(l_method)
        RECEIVING
          value = ro_data.
    ENDIF.

  ENDMETHOD.
ENDCLASS.
