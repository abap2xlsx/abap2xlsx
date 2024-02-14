INTERFACE zif_excel_xml_document
  PUBLIC.

  INTERFACES zif_excel_xml_node.

  ALIASES append_child
    FOR zif_excel_xml_node~append_child.
  ALIASES get_first_child
    FOR zif_excel_xml_node~get_first_child.

  METHODS create_element
    IMPORTING
      !name TYPE string
*      !NAMESPACE type STRING default ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS create_simple_element
    IMPORTING
      !name TYPE string
*      !NAMESPACE type STRING default ''
      !parent TYPE REF TO zif_excel_xml_node
*      !VALUE type STRING default ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS create_simple_element_ns
    IMPORTING
      !name TYPE string
      !parent TYPE REF TO zif_excel_xml_node
      !prefix TYPE string DEFAULT ''
*      !URI type STRING default ''
*      !VALUE type STRING default ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS find_from_name
    IMPORTING
*      !DEPTH type I default 0
      !name TYPE string
*      !NAMESPACE type STRING default ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS find_from_name_ns
    IMPORTING
*      !DEPTH type I default 0
      !name TYPE string
      !uri TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS get_elements_by_tag_name
    IMPORTING
*      !depth TYPE i DEFAULT 0
      !name TYPE string
*      !namespace TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node_collection.

  METHODS get_elements_by_tag_name_ns
    IMPORTING
*      !depth TYPE i DEFAULT 0
      !name TYPE string
      !uri TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_node_collection.

  METHODS get_root_element
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS set_declaration
    IMPORTING
      !declaration TYPE boolean.

  METHODS set_encoding
    IMPORTING
      !encoding TYPE REF TO zif_excel_xml_encoding.

  METHODS set_standalone
    IMPORTING
      !standalone TYPE abap_bool.

ENDINTERFACE.
