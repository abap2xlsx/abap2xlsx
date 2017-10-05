class ZCL_EXCEL_CONVERTER_RESULT definition
  public
  inheriting from ZCL_EXCEL_CONVERTER_ALV
  abstract
  create public .

*"* public components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
public section.
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
protected section.

  methods GET_TABLE
    importing
      !IO_OBJECT type ref to OBJECT
    returning
      value(RO_DATA) type ref to DATA .
*"* private components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
*"* private components of class ZCL_EXCEL_CONVERTER_RESULT
*"* do not include other source files here!!!
private section.
ENDCLASS.



CLASS ZCL_EXCEL_CONVERTER_RESULT IMPLEMENTATION.


method GET_TABLE.
  DATA: lo_object    TYPE REF TO object,
        ls_seoclass  TYPE seoclass,
        l_method     TYPE string.

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

  endmethod.
ENDCLASS.
