INTERFACE zif_excel_xml_parser
  PUBLIC.

  "! Mode for SET_VALIDATING
  CONSTANTS co_no_validation TYPE i VALUE 0 ##NO_TEXT.

  METHODS add_strip_space_element.
  " importing
  "   !NAME type STRING default '*'
  "   !URI type STRING default ''
  " returning
  "   value(RVAL) type BOOLEAN .

  METHODS parse.
*    RETURNING
*      VALUE(rval) TYPE i.

  METHODS set_normalizing
    IMPORTING
      !is_normalizing TYPE boolean DEFAULT 'X'.
*    RETURNING
*      VALUE(rval) TYPE boolean.

  METHODS set_validating
    IMPORTING
      !mode TYPE i DEFAULT '1'.
*    RETURNING
*      VALUE(rval) TYPE boolean.

ENDINTERFACE.
