INTERFACE zif_excel_xml_node
  PUBLIC.

  INTERFACES zif_excel_xml_unknown.

  ALIASES query_interface
    FOR zif_excel_xml_unknown~query_interface.

  CONSTANTS co_node_attribute         TYPE i VALUE 8 ##NO_TEXT.
  CONSTANTS co_node_attribute_decl    TYPE i VALUE 2097152 ##NO_TEXT.
  CONSTANTS co_node_att_list_decl     TYPE i VALUE 1048576 ##NO_TEXT.
  CONSTANTS co_node_cdata_section     TYPE i VALUE 32 ##NO_TEXT.
  CONSTANTS co_node_comment           TYPE i VALUE 512 ##NO_TEXT.
  CONSTANTS co_node_cond_dtd_section  TYPE i VALUE 131072 ##NO_TEXT.
  CONSTANTS co_node_content_particle  TYPE i VALUE 524288 ##NO_TEXT.
  CONSTANTS co_node_document          TYPE i VALUE 1 ##NO_TEXT.
  CONSTANTS co_node_document_fragment TYPE i VALUE 2 ##NO_TEXT.
  CONSTANTS co_node_document_type     TYPE i VALUE 65536 ##NO_TEXT.
  CONSTANTS co_node_element           TYPE i VALUE 4 ##NO_TEXT.
  CONSTANTS co_node_element_decl      TYPE i VALUE 262144 ##NO_TEXT.
  CONSTANTS co_node_entity_decl       TYPE i VALUE 4194304 ##NO_TEXT.
  CONSTANTS co_node_entity_ref        TYPE i VALUE 64 ##NO_TEXT.
  CONSTANTS co_node_namespace_decl    TYPE i VALUE 16777216 ##NO_TEXT.
  CONSTANTS co_node_notations_decl    TYPE i VALUE 8388608 ##NO_TEXT.
  CONSTANTS co_node_pi_parsed         TYPE i VALUE 256 ##NO_TEXT.
  CONSTANTS co_node_pi_unparsed       TYPE i VALUE 128 ##NO_TEXT.
  CONSTANTS co_node_text              TYPE i VALUE 16 ##NO_TEXT.
  CONSTANTS co_node_xxx               TYPE i VALUE 0 ##NO_TEXT.

  METHODS append_child
    IMPORTING
      !new_child TYPE REF TO zif_excel_xml_node.
*    RETURNING
*      VALUE(rval) TYPE i.

  METHODS clone
*    IMPORTING
*      !depth TYPE i DEFAULT -1
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node.

  METHODS create_iterator
*    IMPORTING
*      !depth TYPE i DEFAULT 0
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node_iterator.

  METHODS get_attributes
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_named_node_map.

  METHODS get_children
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node_list.

  METHODS get_first_child
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node.

  METHODS get_name
    RETURNING
      VALUE(rval) TYPE string.

  METHODS get_namespace_prefix
    RETURNING
      VALUE(rval) TYPE string.

  METHODS get_namespace_uri
    RETURNING
      VALUE(rval) TYPE string.

  METHODS get_next
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node.

*  METHODS get_type
*    RETURNING
*      VALUE(rval) TYPE i.

  METHODS get_value
    RETURNING
      VALUE(rval) TYPE string.

  METHODS set_value
    IMPORTING
      !value TYPE string.
*    RETURNING
*      VALUE(rval) TYPE i.
ENDINTERFACE.
