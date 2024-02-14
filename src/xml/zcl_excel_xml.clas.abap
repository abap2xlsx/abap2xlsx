CLASS zcl_excel_xml DEFINITION
  PUBLIC
  FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml.

    CLASS-METHODS create
*      IMPORTING
*        !type TYPE i DEFAULT 0
      RETURNING
        VALUE(rval) TYPE REF TO zif_excel_xml.

    CLASS-METHODS rewrite_xml_via_sxml
      IMPORTING
        iv_xml_string TYPE string
      RETURNING
        VALUE(rv_string) TYPE string.

  PROTECTED SECTION.
  PRIVATE SECTION.
    CLASS-DATA singleton TYPE REF TO zif_excel_xml.
ENDCLASS.


CLASS zcl_excel_xml IMPLEMENTATION.
  METHOD create.
    IF singleton IS NOT BOUND.
      singleton = lcl_isxml=>get_singleton( ).
    ENDIF.
    rval = singleton.
  ENDMETHOD.

  METHOD rewrite_xml_via_sxml.
    rv_string = lcl_rewrite_xml_via_sxml=>execute( iv_xml_string ).
  ENDMETHOD.

  METHOD zif_excel_xml~create_document.
    rval = singleton->create_document( ).
  ENDMETHOD.

  METHOD zif_excel_xml~create_encoding.
    rval = singleton->create_encoding( byte_order    = byte_order
                                       character_set = character_set ).
  ENDMETHOD.

  METHOD zif_excel_xml~create_parser.
    rval = singleton->create_parser( document       = document
                                     istream        = istream
                                     stream_factory = stream_factory ).
  ENDMETHOD.

  METHOD zif_excel_xml~create_renderer.
    rval = singleton->create_renderer( document = document
                                       ostream  = ostream ).
  ENDMETHOD.

  METHOD zif_excel_xml~create_stream_factory.
    rval = singleton->create_stream_factory( ).
  ENDMETHOD.
ENDCLASS.
