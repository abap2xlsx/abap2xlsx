INTERFACE zif_excel_xml_stream_factory
  PUBLIC.

  METHODS create_istream_string
    IMPORTING
      !string TYPE string
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_istream.

  METHODS create_istream_xstring
    IMPORTING
      !string TYPE xstring
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_istream.

  METHODS create_ostream_cstring
    IMPORTING
      !string TYPE REF TO string
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_ostream.

  METHODS create_ostream_xstring
    IMPORTING
      !string TYPE REF TO xstring
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_ostream.

ENDINTERFACE.
