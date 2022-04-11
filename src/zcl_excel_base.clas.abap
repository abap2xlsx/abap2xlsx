"! <p class="shorttext synchronized" lang="en">Base class for Excel components and groups of components</p>
CLASS zcl_excel_base DEFINITION
  PUBLIC
  ABSTRACT
  CREATE PROTECTED.

  PUBLIC SECTION.
    METHODS clone
          ABSTRACT
      RETURNING
        VALUE(ro_object) TYPE REF TO zcl_excel_base
      RAISING
        zcx_excel.

  PROTECTED SECTION.
  PRIVATE SECTION.

ENDCLASS.



CLASS zcl_excel_base IMPLEMENTATION.
ENDCLASS.
