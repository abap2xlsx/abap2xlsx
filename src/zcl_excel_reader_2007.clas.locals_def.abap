*"* use this source file for any type declarations (class
*"* definitions, interfaces or data types) you need for method
*"* implementation or private method's signature

*
CLASS lcl_zip_archive DEFINITION ABSTRACT.
  PUBLIC SECTION.
    METHODS read ABSTRACT
      IMPORTING i_filename       TYPE csequence
      RETURNING VALUE(r_content) TYPE xstring " Remember copy-on-write!
      RAISING   zcx_excel.
ENDCLASS.
