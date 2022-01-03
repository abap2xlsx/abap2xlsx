*"* local class implementation for public class
*"* use this source file for the implementation part of
*"* local helper classes
  TYPES: BEGIN OF t_relationship,
    id     TYPE string,
    type   TYPE string,
    target TYPE string,
    END OF t_relationship.

*
  CLASS lcl_abap_zip_archive DEFINITION
      INHERITING FROM lcl_zip_archive
      CREATE PRIVATE.
    PUBLIC SECTION.
      CLASS-METHODS create
        IMPORTING i_data TYPE xstring
        RETURNING VALUE(r_zip) TYPE REF TO lcl_zip_archive
        RAISING zcx_excel.
      METHODS read REDEFINITION.
    PRIVATE SECTION.
      DATA: abap_zip TYPE REF TO cl_abap_zip.
      METHODS constructor IMPORTING i_data TYPE xstring
                          RAISING zcx_excel.
  ENDCLASS.                    "lcl_abap_zip_archive DEFINITION

*
  CLASS lcl_alternate_zip_archive DEFINITION
    INHERITING FROM lcl_zip_archive
    CREATE PRIVATE.
    PUBLIC SECTION.
      CLASS-METHODS create
        IMPORTING i_data TYPE xstring
                  i_alternate_zip_class TYPE seoclsname
        RETURNING VALUE(r_zip) TYPE REF TO lcl_zip_archive
        RAISING zcx_excel.
      METHODS read REDEFINITION.
    PRIVATE SECTION.
      DATA: alternate_zip TYPE REF TO object.
      METHODS constructor
        IMPORTING i_data TYPE xstring
                  i_alternate_zip_class TYPE seoclsname
        RAISING zcx_excel.
  ENDCLASS.                    "lcl_alternate_zip_archive DEFINITION

*
  CLASS lcl_abap_zip_archive IMPLEMENTATION.
    METHOD create.
      CREATE OBJECT r_zip TYPE lcl_abap_zip_archive
        EXPORTING
          i_data = i_data.
    ENDMETHOD.                    "create
    METHOD constructor.
      DATA: lv_errormessage TYPE string.
      super->constructor( ).
      CREATE OBJECT abap_zip.
      abap_zip->load(
                 EXPORTING
                   zip = i_data
                 EXCEPTIONS
                   zip_parse_error = 1
                   OTHERS          = 2 ).
      IF sy-subrc <> 0.
        lv_errormessage = 'ZIP parse error'(002).
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = lv_errormessage.
      ENDIF.

    ENDMETHOD.                    "constructor
    METHOD read.
      DATA: lv_errormessage TYPE string.
      CALL METHOD abap_zip->get
        EXPORTING
          name                    = i_filename
        IMPORTING
          content                 = r_content
        EXCEPTIONS
          zip_index_error         = 1
          zip_decompression_error = 2
          OTHERS                  = 3.
      IF sy-subrc <> 0.
        lv_errormessage = 'File not found in zip-archive'(003).
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = lv_errormessage.
      ENDIF.

    ENDMETHOD.                    "read
  ENDCLASS.                    "lcl_abap_zip_archive IMPLEMENTATION

*
  CLASS lcl_alternate_zip_archive IMPLEMENTATION.
    METHOD create.
      CREATE OBJECT r_zip TYPE lcl_alternate_zip_archive
        EXPORTING
          i_alternate_zip_class = i_alternate_zip_class
          i_data                = i_data.
    ENDMETHOD.                    "create
    METHOD constructor.
      DATA: lv_errormessage TYPE string.
      super->constructor( ).
      CREATE OBJECT alternate_zip TYPE (i_alternate_zip_class).
      TRY.
          CALL METHOD alternate_zip->('LOAD')
            EXPORTING
              zip             = i_data
            EXCEPTIONS
              zip_parse_error = 1
              OTHERS          = 2.
        CATCH cx_sy_dyn_call_illegal_method.
          lv_errormessage = 'Method LOAD missing in alternative zipclass'. "#EC NOTEXT   This is a workaround until class CL_ABAP_ZIP is fixed
          RAISE EXCEPTION TYPE zcx_excel
            EXPORTING
              error = lv_errormessage.
      ENDTRY.

      IF sy-subrc <> 0.
        lv_errormessage = 'ZIP parse error'(002).
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = lv_errormessage.
      ENDIF.

    ENDMETHOD.                    "constructor
    METHOD read.
      DATA: lv_errormessage TYPE string.
      TRY.
          CALL METHOD alternate_zip->('GET')
            EXPORTING
              name                    = i_filename
            IMPORTING
              content                 = r_content    " Contents
            EXCEPTIONS
              zip_index_error         = 1
              zip_decompression_error = 2
              OTHERS                  = 3.
        CATCH cx_sy_dyn_call_illegal_method.
          lv_errormessage = 'Method GET missing in alternative zipclass'. "#EC NOTEXT   This is a workaround until class CL_ABAP_ZIP is fixed
          RAISE EXCEPTION TYPE zcx_excel
            EXPORTING
              error = lv_errormessage.
      ENDTRY.
      IF sy-subrc <> 0.
        lv_errormessage = 'File not found in zip-archive'(003).
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = lv_errormessage.
      ENDIF.

    ENDMETHOD.                    "read
  ENDCLASS.                    "lcl_alternate_zip_archive IMPLEMENTATION
