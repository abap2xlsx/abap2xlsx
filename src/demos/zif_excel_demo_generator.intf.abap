INTERFACE zif_excel_demo_generator
  PUBLIC .
  TYPES: BEGIN OF ty_information,
           objid    TYPE wwwdata-objid,
           text     TYPE wwwdata-text,
           filename TYPE string,
         END OF ty_information.
  METHODS generate_excel
    RETURNING
      VALUE(result) TYPE REF TO zcl_excel
    RAISING
      zcx_excel.
  METHODS get_information
    RETURNING
      VALUE(result) TYPE ty_information.
  METHODS get_next_generator
    RETURNING
      VALUE(result) TYPE REF TO zif_excel_demo_generator.
  METHODS checker_initialization.
  METHODS cleanup_for_diff
    IMPORTING
      xstring       TYPE xstring
    RETURNING
      VALUE(result) TYPE REF TO cl_abap_zip.
ENDINTERFACE.
