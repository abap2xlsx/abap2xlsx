*"* use this source file for any type of declarations (class
*"* definitions, interfaces or type declarations) you need for
*"* components in the private section

CLASS lcl_isxml_attribute DEFINITION DEFERRED.
CLASS lcl_isxml_character_data DEFINITION DEFERRED.
CLASS lcl_isxml_document DEFINITION DEFERRED.
CLASS lcl_isxml_element DEFINITION DEFERRED.
CLASS lcl_isxml_encoding DEFINITION DEFERRED.
CLASS lcl_isxml DEFINITION DEFERRED.
CLASS lcl_isxml_istream_string DEFINITION DEFERRED.
CLASS lcl_isxml_istream_xstring DEFINITION DEFERRED.
CLASS lcl_isxml_named_node_map DEFINITION DEFERRED.
CLASS lcl_isxml_node DEFINITION DEFERRED.
CLASS lcl_isxml_node_collection DEFINITION DEFERRED.
CLASS lcl_isxml_node_iterator DEFINITION DEFERRED.
CLASS lcl_isxml_node_list DEFINITION DEFERRED.
CLASS lcl_isxml_ostream_string DEFINITION DEFERRED.
CLASS lcl_isxml_ostream_xstring DEFINITION DEFERRED.
CLASS lcl_isxml_parser DEFINITION DEFERRED.
CLASS lcl_isxml_renderer DEFINITION DEFERRED.
CLASS lcl_isxml_stream DEFINITION DEFERRED.
CLASS lcl_isxml_stream_factory DEFINITION DEFERRED.
CLASS lcl_isxml_text DEFINITION DEFERRED.
CLASS lcl_isxml_unknown DEFINITION DEFERRED.

INTERFACE lif_isxml_all_friends.
ENDINTERFACE.


CLASS lcl_isxml_root_all DEFINITION.
  PUBLIC SECTION.
    INTERFACES lif_isxml_all_friends.
ENDCLASS.


CLASS lcl_isxml DEFINITION
    INHERITING FROM lcl_isxml_root_all
    CREATE PROTECTED
    FRIENDS zcl_excel_xml
            lif_isxml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml.

    TYPES tv_element_name_id TYPE i.
    TYPES tv_node_id         TYPE i.
    TYPES tv_node_type       TYPE i.
    TYPES tv_namespace_id    TYPE i.
    TYPES tt_node_id         TYPE STANDARD TABLE OF tv_node_id WITH DEFAULT KEY.
    TYPES:
      BEGIN OF ts_node,
        id        TYPE tv_node_id,
        type      TYPE tv_node_type,
        parent_id TYPE tv_node_id,
        text      TYPE string,
        object    TYPE REF TO object,
      END OF ts_node.
    TYPES tt_node TYPE SORTED TABLE OF ts_node WITH UNIQUE KEY id
                    WITH NON-UNIQUE SORTED KEY by_parent COMPONENTS parent_id.
    TYPES:
      BEGIN OF ts_namespace,
        id  TYPE tv_namespace_id,
        uri TYPE string,
      END OF ts_namespace.
    TYPES tt_namespace TYPE HASHED TABLE OF ts_namespace WITH UNIQUE KEY id
                    WITH UNIQUE HASHED KEY by_uri COMPONENTS uri.
    TYPES:
      BEGIN OF ts_element_name,
        id           TYPE tv_element_name_id,
        name         TYPE string,
        namespace_id TYPE tv_namespace_id,
      END OF ts_element_name.
    TYPES tt_element_name TYPE HASHED TABLE OF ts_element_name WITH UNIQUE KEY id
                    WITH NON-UNIQUE SORTED KEY by_name COMPONENTS name namespace_id.
    TYPES:
      BEGIN OF ts_attribute,
        name  TYPE string,
        value TYPE string,
      END OF ts_attribute.
    TYPES tt_attribute TYPE STANDARD TABLE OF ts_attribute WITH DEFAULT KEY.
    TYPES:
      BEGIN OF ts_element,
        id         TYPE tv_node_id,
        name_id    TYPE tv_element_name_id,
        attributes TYPE tt_attribute,
        object     TYPE REF TO lcl_isxml_element,
      END OF ts_element.
    TYPES tt_element TYPE HASHED TABLE OF ts_element WITH UNIQUE KEY id
                    WITH NON-UNIQUE SORTED KEY by_name COMPONENTS name_id.
    TYPES:
      BEGIN OF ts_parse_element_level,
        level   TYPE i,
        node_id TYPE tv_node_id,
      END OF ts_parse_element_level.
    TYPES tt_parse_element_level TYPE HASHED TABLE OF ts_parse_element_level WITH UNIQUE KEY level.

  PRIVATE SECTION.

    CLASS-METHODS get_singleton
      RETURNING
        VALUE(rval) TYPE REF TO zif_excel_xml.

    CLASS-DATA singleton TYPE REF TO zif_excel_xml."lcl_isxml.
    CLASS-DATA no_node   TYPE REF TO lcl_isxml_node.

ENDCLASS.
