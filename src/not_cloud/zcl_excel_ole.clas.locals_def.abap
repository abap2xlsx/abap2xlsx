*"* use this source file for any type of declarations (class
*"* definitions, interfaces or type declarations) you need for
*"* components in the private section

*--------------------------------------------------------------------*
* CLASS c_oi_proxy_error
*--------------------------------------------------------------------*
* use for method bind_ALV
*--------------------------------------------------------------------*
CLASS c_oi_proxy_error DEFINITION.
  PUBLIC SECTION.
    INTERFACES: i_oi_error.
    DATA: error_nr TYPE i.
    DATA: error_string TYPE sy-msgv1.

    METHODS: constructor IMPORTING object_name TYPE c
                                   method_name TYPE c.
  PRIVATE SECTION.
    CONSTANTS:
          ret_call_not_flushed        TYPE i VALUE -999999.

    DATA: message_id TYPE sy-msgid,
          message_nr TYPE sy-msgno,
          param1 TYPE sy-msgv1,
          param2 TYPE sy-msgv2,
          param3 TYPE sy-msgv3,
          param4 TYPE sy-msgv4.
ENDCLASS.                    "c_oi_proxy_error DEFINITION

*--------------------------------------------------------------------*
* CLASS lcl_gui_alv_grid
*--------------------------------------------------------------------*
* to get protected attribute and method of cl_gui_alv_grid
* use for method bind_ALV
*--------------------------------------------------------------------*
CLASS lcl_gui_alv_grid DEFINITION INHERITING FROM cl_gui_alv_grid.

  PUBLIC SECTION.
* get ALV grid data
    METHODS: get_alv_attributes
      IMPORTING
        io_grid   TYPE REF TO cl_gui_alv_grid " ALV grid
      EXPORTING
        et_table  TYPE REF TO data.           " dta table

ENDCLASS.                    "lcl_gui_alv_grid DEFINITION
