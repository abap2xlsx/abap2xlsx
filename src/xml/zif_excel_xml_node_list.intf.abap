INTERFACE zif_excel_xml_node_list
  PUBLIC.

  METHODS create_iterator
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node_iterator.
ENDINTERFACE.
