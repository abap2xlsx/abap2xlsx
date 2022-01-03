CLASS zcl_excel_hyperlink DEFINITION
  PUBLIC
  FINAL
  CREATE PRIVATE .

*"* public components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CLASS-METHODS create_external_link
      IMPORTING
        !iv_url        TYPE string
      RETURNING
        VALUE(ov_link) TYPE REF TO zcl_excel_hyperlink .
    CLASS-METHODS create_internal_link
      IMPORTING
        !iv_location   TYPE string
      RETURNING
        VALUE(ov_link) TYPE REF TO zcl_excel_hyperlink .
    METHODS is_internal
      RETURNING
        VALUE(ev_ret) TYPE abap_bool .
    METHODS set_cell_reference
      IMPORTING
        !ip_column TYPE simple
        !ip_row    TYPE zexcel_cell_row
      RAISING
        zcx_excel .
    METHODS get_ref
      RETURNING
        VALUE(ev_ref) TYPE string .
    METHODS get_url
      RETURNING
        VALUE(ev_url) TYPE string .
*"* protected components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA location TYPE string .
    DATA cell_reference TYPE string .
    DATA internal TYPE abap_bool .
    DATA column TYPE zexcel_cell_column_alpha .
    DATA row TYPE zexcel_cell_row .

    CLASS-METHODS create
      IMPORTING
        !iv_url        TYPE string
        !iv_internal   TYPE abap_bool
      RETURNING
        VALUE(ov_link) TYPE REF TO zcl_excel_hyperlink .
ENDCLASS.



CLASS zcl_excel_hyperlink IMPLEMENTATION.


  METHOD create.
    DATA: lo_hyperlink TYPE REF TO zcl_excel_hyperlink.

    CREATE OBJECT lo_hyperlink.

    lo_hyperlink->location = iv_url.
    lo_hyperlink->internal = iv_internal.

    ov_link = lo_hyperlink.
  ENDMETHOD.


  METHOD create_external_link.

    ov_link = zcl_excel_hyperlink=>create( iv_url = iv_url
                                           iv_internal = abap_false ).
  ENDMETHOD.


  METHOD create_internal_link.
    ov_link = zcl_excel_hyperlink=>create( iv_url = iv_location
                                           iv_internal = abap_true ).
  ENDMETHOD.


  METHOD get_ref.
    ev_ref = row.
    CONDENSE ev_ref.
    CONCATENATE column ev_ref INTO ev_ref.
  ENDMETHOD.


  METHOD get_url.
    ev_url = me->location.
  ENDMETHOD.


  METHOD is_internal.
    ev_ret = me->internal.
  ENDMETHOD.


  METHOD set_cell_reference.
    me->column = zcl_excel_common=>convert_column2alpha( ip_column ). " issue #155 - less restrictive typing for ip_column
    me->row = ip_row.
  ENDMETHOD.
ENDCLASS.
