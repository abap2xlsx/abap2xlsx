INTERFACE zif_excel_xml_unknown
  PUBLIC.

  METHODS query_interface
    IMPORTING
      !iid TYPE i
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_unknown.

ENDINTERFACE.
