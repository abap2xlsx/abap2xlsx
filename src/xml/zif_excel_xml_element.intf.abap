INTERFACE zif_excel_xml_element
  PUBLIC.

  INTERFACES zif_excel_xml_node.

  ALIASES append_child
    FOR zif_excel_xml_node~append_child.
  ALIASES clone
    FOR zif_excel_xml_node~clone.
  ALIASES create_iterator
    FOR zif_excel_xml_node~create_iterator.
  ALIASES get_attributes
    FOR zif_excel_xml_node~get_attributes.
  ALIASES get_children
    FOR zif_excel_xml_node~get_children.
  ALIASES get_first_child
    FOR zif_excel_xml_node~get_first_child.
  ALIASES get_name
    FOR zif_excel_xml_node~get_name.
  ALIASES get_next
    FOR zif_excel_xml_node~get_next.
  ALIASES get_value
    FOR zif_excel_xml_node~get_value.
  ALIASES set_value
    FOR zif_excel_xml_node~set_value.

  METHODS find_from_name
    IMPORTING
*      !depth TYPE i DEFAULT 0
      !name TYPE string
*      !namespace TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS find_from_name_ns
    IMPORTING
      !depth TYPE i DEFAULT 0
      !name TYPE string
      !uri TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_element.

  METHODS get_attribute
    IMPORTING
      !name TYPE string
*      !namespace TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE string.

  METHODS get_attribute_node_ns
    IMPORTING
      !name TYPE string
      !uri TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE REF TO zif_excel_xml_attribute.

  METHODS get_attribute_ns
    IMPORTING
      !name TYPE string
      !uri TYPE string DEFAULT ''
    RETURNING
      VALUE(rval) TYPE string.

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

  METHODS remove_attribute_ns
    IMPORTING
      !name TYPE string.
*      !uri TYPE string DEFAULT ''
*    RETURNING
*      VALUE(rval) TYPE i.

  METHODS set_attribute
    IMPORTING
      !name TYPE string
      !namespace TYPE string DEFAULT ''
      !value TYPE string DEFAULT ''.
*    RETURNING
*      VALUE(rval) TYPE i.

  METHODS set_attribute_ns
    IMPORTING
      !name TYPE string
      !prefix TYPE string DEFAULT ''
*      !uri TYPE string DEFAULT ''
      !value TYPE string DEFAULT ''.
*    RETURNING
*      VALUE(rval) TYPE i.

ENDINTERFACE.
