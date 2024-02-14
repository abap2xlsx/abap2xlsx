INTERFACE zif_excel_xml
  PUBLIC.

  METHODS create_document
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_document.

  METHODS create_encoding
    IMPORTING
      !byte_order TYPE i
      !character_set TYPE string
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_encoding.

  METHODS create_parser
    IMPORTING
      !document TYPE REF TO zif_excel_xml_document
      !istream TYPE REF TO zif_excel_xml_istream
      !stream_factory TYPE REF TO zif_excel_xml_stream_factory
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_parser.

  METHODS create_renderer
    IMPORTING
      !document TYPE REF TO zif_excel_xml_document
      !ostream TYPE REF TO zif_excel_xml_ostream
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_renderer.

  METHODS create_stream_factory
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_stream_factory.

ENDINTERFACE.
