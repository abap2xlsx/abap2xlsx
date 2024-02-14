INTERFACE zif_excel_xml_named_node_map
  PUBLIC.

  METHODS create_iterator
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node_iterator.

ENDINTERFACE.
