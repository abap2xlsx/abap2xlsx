CLASS zcl_excel_style DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLE
*"* do not include other source files here!!!
  PUBLIC SECTION.

    DATA font TYPE REF TO zcl_excel_style_font .
    DATA fill TYPE REF TO zcl_excel_style_fill .
    DATA borders TYPE REF TO zcl_excel_style_borders .
    DATA alignment TYPE REF TO zcl_excel_style_alignment .
    DATA number_format TYPE REF TO zcl_excel_style_number_format .
    DATA protection TYPE REF TO zcl_excel_style_protection .

    METHODS constructor
      IMPORTING
        !ip_guid TYPE zexcel_cell_style OPTIONAL .
    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE zexcel_cell_style .
*"* protected components of class ZABAP_EXCEL_STYLE
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA guid TYPE zexcel_cell_style .
ENDCLASS.



CLASS zcl_excel_style IMPLEMENTATION.


  METHOD constructor.


    CREATE OBJECT font.
    CREATE OBJECT fill.
    CREATE OBJECT borders.
    CREATE OBJECT alignment.
    CREATE OBJECT number_format.
    CREATE OBJECT protection.

    IF ip_guid IS NOT INITIAL.
      me->guid = ip_guid.
    ELSE.
      me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).
    ENDIF.

  ENDMETHOD.


  METHOD get_guid.


    ep_guid = me->guid.
  ENDMETHOD.
ENDCLASS.
