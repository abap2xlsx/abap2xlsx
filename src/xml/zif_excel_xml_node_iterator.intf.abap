INTERFACE zif_excel_xml_node_iterator
  PUBLIC.

  METHODS get_next
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node.

ENDINTERFACE.
