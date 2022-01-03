*"* use this source file for any type of declarations (class
*"* definitions, interfaces or type declarations) you need for
*"* components in the private section

* Signal for "Not found"
CLASS lcx_not_found DEFINITION INHERITING FROM cx_static_check.
  PUBLIC SECTION.
    DATA error TYPE string.
    METHODS constructor
      IMPORTING error    TYPE string
                textid   TYPE sotr_conc OPTIONAL
                previous TYPE REF TO cx_root OPTIONAL.
    METHODS if_message~get_text REDEFINITION.
ENDCLASS.
