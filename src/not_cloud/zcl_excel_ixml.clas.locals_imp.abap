*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations

CLASS lcx_unexpected DEFINITION INHERITING FROM cx_no_check.
ENDCLASS.


INTERFACE lif_wrap_ixml_all_friends.
ENDINTERFACE.


INTERFACE lif_wrap_ixml_istream.
  INTERFACES zif_excel_xml_istream.
  DATA ixml_istream TYPE REF TO if_ixml_istream.
ENDINTERFACE.


INTERFACE lif_wrap_ixml_ostream.
  INTERFACES zif_excel_xml_ostream.
  DATA ixml_ostream TYPE REF TO if_ixml_ostream.
ENDINTERFACE.


CLASS lcl_wrap_ixml_unknown DEFINITION
    CREATE PROTECTED.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_unknown.
    INTERFACES lif_wrap_ixml_all_friends.

ENDCLASS.


CLASS lcl_wrap_ixml_node DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PROTECTED
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_node.

  PRIVATE SECTION.

    DATA ixml_node TYPE REF TO if_ixml_node.

ENDCLASS.


CLASS lcl_wrap_ixml DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml.

    TYPES:
      BEGIN OF ts_wrapped_ixml_object,
        ixml_object         TYPE REF TO object,
        ixml_object_wrapper TYPE REF TO object,
      END OF ts_wrapped_ixml_object.
    TYPES tt_wrapped_ixml_object TYPE HASHED TABLE OF ts_wrapped_ixml_object WITH UNIQUE KEY ixml_object
                                WITH UNIQUE HASHED KEY by_wrapper COMPONENTS ixml_object_wrapper.

    CLASS-DATA wrapped_ixml_objects TYPE tt_wrapped_ixml_object.

    CLASS-METHODS get_singleton
      RETURNING
        VALUE(ro_result) TYPE REF TO zif_excel_xml.

    CLASS-METHODS unwrap_ixml
      IMPORTING
        io_wrap_ixml_unknown TYPE REF TO lcl_wrap_ixml_unknown
      RETURNING
        VALUE(ro_result)     TYPE REF TO object.

    CLASS-METHODS wrap_ixml
      IMPORTING
        io_ixml_unknown  TYPE REF TO if_ixml_unknown
      RETURNING
        VALUE(ro_result) TYPE REF TO object.

  PRIVATE SECTION.

    CLASS-DATA singleton TYPE REF TO lcl_wrap_ixml.
    DATA ixml TYPE REF TO if_ixml.

ENDCLASS.


CLASS lcl_wrap_ixml_attribute DEFINITION
    INHERITING FROM lcl_wrap_ixml_node
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_attribute.

  PRIVATE SECTION.

    DATA ixml_attribute TYPE REF TO if_ixml_attribute.

ENDCLASS.


CLASS lcl_wrap_ixml_character_data DEFINITION
    INHERITING FROM lcl_wrap_ixml_node
    CREATE PROTECTED
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_character_data.

ENDCLASS.


CLASS lcl_wrap_ixml_document DEFINITION
    INHERITING FROM lcl_wrap_ixml_node
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_document.

  PRIVATE SECTION.

    DATA ixml_document TYPE REF TO if_ixml_document.

ENDCLASS.


CLASS lcl_wrap_ixml_element DEFINITION
    INHERITING FROM lcl_wrap_ixml_node
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_element.

  PRIVATE SECTION.

    DATA ixml_element TYPE REF TO if_ixml_element.

ENDCLASS.


CLASS lcl_wrap_ixml_encoding DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_encoding.

  PRIVATE SECTION.

    DATA ixml_encoding TYPE REF TO if_ixml_encoding.

ENDCLASS.


CLASS lcl_wrap_ixml_istream_string DEFINITION
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES lif_wrap_ixml_istream.

ENDCLASS.


CLASS lcl_wrap_ixml_istream_xstring DEFINITION
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES lif_wrap_ixml_istream.

ENDCLASS.


CLASS lcl_wrap_ixml_named_node_map DEFINITION
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_named_node_map.

  PRIVATE SECTION.

    DATA ixml_named_node_map TYPE REF TO if_ixml_named_node_map.

ENDCLASS.


CLASS lcl_wrap_ixml_node_collection DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_node_collection.

  PRIVATE SECTION.

    DATA ixml_node_collection TYPE REF TO if_ixml_node_collection.

ENDCLASS.


CLASS lcl_wrap_ixml_node_iterator DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_node_iterator.

  PRIVATE SECTION.

    DATA ixml_node_iterator TYPE REF TO if_ixml_node_iterator.

ENDCLASS.


CLASS lcl_wrap_ixml_node_list DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_node_list.

  PRIVATE SECTION.

    DATA ixml_node_list TYPE REF TO if_ixml_node_list.

ENDCLASS.


CLASS lcl_wrap_ixml_ostream_string DEFINITION
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES lif_wrap_ixml_ostream.

ENDCLASS.


CLASS lcl_wrap_ixml_ostream_xstring DEFINITION
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES lif_wrap_ixml_ostream.

ENDCLASS.


CLASS lcl_wrap_ixml_parser DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_parser.

  PRIVATE SECTION.

    DATA ixml_parser TYPE REF TO if_ixml_parser.

ENDCLASS.


CLASS lcl_wrap_ixml_renderer DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_renderer.

  PRIVATE SECTION.

    DATA ixml_renderer TYPE REF TO if_ixml_renderer.

ENDCLASS.


CLASS lcl_wrap_ixml_stream DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_stream.

ENDCLASS.


CLASS lcl_wrap_ixml_stream_factory DEFINITION
    INHERITING FROM lcl_wrap_ixml_unknown
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_stream_factory.

  PRIVATE SECTION.

    DATA ixml_stream_factory TYPE REF TO if_ixml_stream_factory.

ENDCLASS.


CLASS lcl_wrap_ixml_text DEFINITION
    INHERITING FROM lcl_wrap_ixml_character_data
    CREATE PRIVATE
    FRIENDS lif_wrap_ixml_all_friends.

  PUBLIC SECTION.

    INTERFACES zif_excel_xml_text.

  PRIVATE SECTION.

    DATA ixml_text TYPE REF TO if_ixml_text.

ENDCLASS.


CLASS lcl_wrap_ixml IMPLEMENTATION.
  METHOD get_singleton.
    IF singleton IS NOT BOUND.
      CREATE OBJECT singleton.
      singleton->ixml = cl_ixml=>create( ).
    ENDIF.
    ro_result = singleton.
  ENDMETHOD.

  METHOD unwrap_ixml.
    DATA lr_wrapped_ixml_object TYPE REF TO ts_wrapped_ixml_object.

    CHECK io_wrap_ixml_unknown IS BOUND.

    READ TABLE wrapped_ixml_objects
         WITH TABLE KEY by_wrapper
         COMPONENTS ixml_object_wrapper = io_wrap_ixml_unknown
         REFERENCE INTO lr_wrapped_ixml_object.
    IF sy-subrc = 0.
      ro_result = lr_wrapped_ixml_object->ixml_object.
    ENDIF.
  ENDMETHOD.

  METHOD wrap_ixml.
    DATA lr_wrapped_ixml_object       TYPE REF TO ts_wrapped_ixml_object.
    DATA ls_wrapped_ixml_object       TYPE ts_wrapped_ixml_object.
    DATA lv_class_name                TYPE string.
    DATA lo_wrap_ixml_attribute       TYPE REF TO lcl_wrap_ixml_attribute.
    DATA lo_wrap_ixml_node            TYPE REF TO lcl_wrap_ixml_node.
    DATA lo_wrap_ixml_document        TYPE REF TO lcl_wrap_ixml_document.
    DATA lo_wrap_ixml_element         TYPE REF TO lcl_wrap_ixml_element.
    DATA lo_wrap_ixml_named_node_map  TYPE REF TO lcl_wrap_ixml_named_node_map.
    DATA lo_wrap_ixml_node_collection TYPE REF TO lcl_wrap_ixml_node_collection.
    DATA lo_wrap_ixml_node_iterator   TYPE REF TO lcl_wrap_ixml_node_iterator.
    DATA lo_wrap_ixml_node_list       TYPE REF TO lcl_wrap_ixml_node_list.
    DATA lo_wrap_ixml_text            TYPE REF TO lcl_wrap_ixml_text.

    IF io_ixml_unknown IS NOT BOUND.
      RETURN.
    ENDIF.

    READ TABLE wrapped_ixml_objects WITH TABLE KEY ixml_object = io_ixml_unknown
         REFERENCE INTO lr_wrapped_ixml_object.
    IF sy-subrc <> 0.
      CLEAR ls_wrapped_ixml_object.
      ls_wrapped_ixml_object-ixml_object = io_ixml_unknown.
      lv_class_name = cl_abap_typedescr=>describe_by_object_ref( io_ixml_unknown )->get_relative_name( ).
      CASE lv_class_name.
        WHEN 'CL_IXML_ATTRIBUTE'.
          CREATE OBJECT lo_wrap_ixml_attribute.
          lo_wrap_ixml_attribute->ixml_attribute ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_attribute.
          lo_wrap_ixml_node ?= lo_wrap_ixml_attribute.
          lo_wrap_ixml_node->ixml_node ?= io_ixml_unknown.
        WHEN 'CL_IXML_DOCUMENT'.
          CREATE OBJECT lo_wrap_ixml_document.
          lo_wrap_ixml_document->ixml_document ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_document.
          lo_wrap_ixml_node ?= lo_wrap_ixml_document.
          lo_wrap_ixml_node->ixml_node ?= io_ixml_unknown.
        WHEN 'CL_IXML_ELEMENT'.
          CREATE OBJECT lo_wrap_ixml_element.
          lo_wrap_ixml_element->ixml_element ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_element.
          lo_wrap_ixml_node ?= lo_wrap_ixml_element.
          lo_wrap_ixml_node->ixml_node ?= io_ixml_unknown.
        WHEN 'CL_IXML_NAMED_NODE_MAP'.
          CREATE OBJECT lo_wrap_ixml_named_node_map.
          lo_wrap_ixml_named_node_map->ixml_named_node_map ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_named_node_map.
        WHEN 'CL_IXML_NODE_COLLECTION'.
          CREATE OBJECT lo_wrap_ixml_node_collection.
          lo_wrap_ixml_node_collection->ixml_node_collection ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_node_collection.
        WHEN 'CL_IXML_NODE_ITERATOR'.
          CREATE OBJECT lo_wrap_ixml_node_iterator.
          lo_wrap_ixml_node_iterator->ixml_node_iterator ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_node_iterator.
        WHEN 'CL_IXML_NODE_LIST'.
          CREATE OBJECT lo_wrap_ixml_node_list.
          lo_wrap_ixml_node_list->ixml_node_list ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_node_list.
        WHEN 'CL_IXML_TEXT'.
          CREATE OBJECT lo_wrap_ixml_text.
          lo_wrap_ixml_text->ixml_text ?= io_ixml_unknown.
          ls_wrapped_ixml_object-ixml_object_wrapper = lo_wrap_ixml_text.
          lo_wrap_ixml_node ?= lo_wrap_ixml_text.
          lo_wrap_ixml_node->ixml_node ?= io_ixml_unknown.
      ENDCASE.
      INSERT ls_wrapped_ixml_object INTO TABLE wrapped_ixml_objects REFERENCE INTO lr_wrapped_ixml_object.
    ENDIF.

    ro_result = lr_wrapped_ixml_object->ixml_object_wrapper.
  ENDMETHOD.

  METHOD zif_excel_xml~create_document.
    DATA lo_ixml_document TYPE REF TO if_ixml_document.

    lo_ixml_document = ixml->create_document( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_document ).
  ENDMETHOD.

  METHOD zif_excel_xml~create_encoding.
    DATA lo_encoding TYPE REF TO lcl_wrap_ixml_encoding.

    CREATE OBJECT lo_encoding.
    lo_encoding->ixml_encoding = ixml->create_encoding( byte_order    = byte_order
                                                        character_set = character_set ).
    rval = lo_encoding.
  ENDMETHOD.

  METHOD zif_excel_xml~create_parser.
    DATA parser                   TYPE REF TO lcl_wrap_ixml_parser.
    DATA wrap_ixml_document       TYPE REF TO lcl_wrap_ixml_document.
    DATA wrap_ixml_istream        TYPE REF TO lif_wrap_ixml_istream.
    DATA wrap_ixml_stream_factory TYPE REF TO lcl_wrap_ixml_stream_factory.

    CREATE OBJECT parser.
    wrap_ixml_document ?= document.
    wrap_ixml_istream ?= istream.
    wrap_ixml_stream_factory ?= stream_factory.
    parser->ixml_parser = ixml->create_parser( document       = wrap_ixml_document->ixml_document
                                               istream        = wrap_ixml_istream->ixml_istream
                                               stream_factory = wrap_ixml_stream_factory->ixml_stream_factory ).
    rval = parser.
  ENDMETHOD.

  METHOD zif_excel_xml~create_renderer.
    DATA lo_renderer        TYPE REF TO lcl_wrap_ixml_renderer.
    DATA wrap_ixml_document TYPE REF TO lcl_wrap_ixml_document.
    DATA wrap_ixml_ostream  TYPE REF TO lif_wrap_ixml_ostream.

    CREATE OBJECT lo_renderer.
    wrap_ixml_document ?= document.
    wrap_ixml_ostream ?= ostream.
    lo_renderer->ixml_renderer ?= ixml->create_renderer( document = wrap_ixml_document->ixml_document
                                                         ostream  = wrap_ixml_ostream->ixml_ostream ).
    rval = lo_renderer.
  ENDMETHOD.

  METHOD zif_excel_xml~create_stream_factory.
    DATA lo_stream_factory TYPE REF TO lcl_wrap_ixml_stream_factory.

    CREATE OBJECT lo_stream_factory.
    lo_stream_factory->ixml_stream_factory = ixml->create_stream_factory( ).
    rval = lo_stream_factory.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_attribute IMPLEMENTATION.

ENDCLASS.


CLASS lcl_wrap_ixml_character_data IMPLEMENTATION.

ENDCLASS.


CLASS lcl_wrap_ixml_document IMPLEMENTATION.
  METHOD zif_excel_xml_document~create_element.
    DATA lo_ixml_element TYPE REF TO if_ixml_element.

    lo_ixml_element = ixml_document->create_element( name = name ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~create_simple_element.
    DATA lo_wrap_ixml_parent TYPE REF TO lcl_wrap_ixml_node.
    DATA lo_ixml_element     TYPE REF TO if_ixml_element.

    lo_wrap_ixml_parent ?= parent.
    lo_ixml_element = ixml_document->create_simple_element( name   = name
                                                            parent = lo_wrap_ixml_parent->ixml_node ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~create_simple_element_ns.
    DATA lo_wrap_ixml_parent TYPE REF TO lcl_wrap_ixml_node.
    DATA lo_ixml_element     TYPE REF TO if_ixml_element.

    lo_wrap_ixml_parent ?= parent.
    lo_ixml_element = ixml_document->create_simple_element_ns( name   = name
                                                               parent = lo_wrap_ixml_parent->ixml_node
                                                               prefix = prefix ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~find_from_name.
    DATA lo_ixml_element TYPE REF TO if_ixml_element.

    lo_ixml_element = ixml_document->find_from_name( name = name ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~find_from_name_ns.
    DATA lo_ixml_element TYPE REF TO if_ixml_element.

    lo_ixml_element = ixml_document->find_from_name_ns( name = name
                                                        uri  = uri ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~get_elements_by_tag_name.
    DATA lo_ixml_node_collection TYPE REF TO if_ixml_node_collection.

    lo_ixml_node_collection = ixml_document->get_elements_by_tag_name( name = name ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_collection ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~get_elements_by_tag_name_ns.
    DATA lo_ixml_node_collection TYPE REF TO if_ixml_node_collection.

    lo_ixml_node_collection = ixml_document->get_elements_by_tag_name_ns( name = name
                                                                          uri  = uri ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_collection ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~get_root_element.
    DATA lo_ixml_element TYPE REF TO if_ixml_element.

    lo_ixml_element = ixml_document->get_root_element( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~set_declaration.
    ixml_document->set_declaration( declaration = declaration ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~set_encoding.
    DATA wrap_ixml_encoding TYPE REF TO lcl_wrap_ixml_encoding.

    wrap_ixml_encoding ?= encoding.
    ixml_document->set_encoding( encoding = wrap_ixml_encoding->ixml_encoding ).
  ENDMETHOD.

  METHOD zif_excel_xml_document~set_standalone.
    ixml_document->set_standalone( standalone ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_element IMPLEMENTATION.
  METHOD zif_excel_xml_element~find_from_name.
    DATA lo_ixml_element TYPE REF TO if_ixml_element.

    lo_ixml_element = ixml_element->find_from_name( name = name ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~find_from_name_ns.
    DATA lo_ixml_element TYPE REF TO if_ixml_element.

    lo_ixml_element = ixml_element->find_from_name_ns( depth = depth
                                                       name  = name
                                                       uri   = uri ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_element ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~get_attribute.
    rval = ixml_element->get_attribute( name = name ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~get_attribute_node_ns.
    DATA lo_ixml_attribute TYPE REF TO if_ixml_attribute.

    lo_ixml_attribute = ixml_element->get_attribute_node_ns( name = name
                                                             uri  = uri ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_attribute ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~get_attribute_ns.
    rval = ixml_element->get_attribute_ns( name = name
                                           uri  = uri ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~get_elements_by_tag_name.
    DATA lo_ixml_node_collection TYPE REF TO if_ixml_node_collection.

    lo_ixml_node_collection = ixml_element->get_elements_by_tag_name( name = name ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_collection ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~get_elements_by_tag_name_ns.
    DATA lo_ixml_node_collection TYPE REF TO if_ixml_node_collection.

    lo_ixml_node_collection = ixml_element->get_elements_by_tag_name_ns( name = name
                                                                         uri  = uri ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_collection ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~remove_attribute_ns.
    ixml_element->remove_attribute_ns( name = name ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~set_attribute.
    ixml_element->set_attribute( name      = name
                                 namespace = namespace
                                 value     = value ).
  ENDMETHOD.

  METHOD zif_excel_xml_element~set_attribute_ns.
    ixml_element->set_attribute_ns( name   = name
                                    prefix = prefix
                                    value  = value ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_istream_string IMPLEMENTATION.
  METHOD zif_excel_xml_stream~close.
    lif_wrap_ixml_istream~ixml_istream->close( ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_istream_xstring IMPLEMENTATION.
  METHOD zif_excel_xml_stream~close.
    lif_wrap_ixml_istream~ixml_istream->close( ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_named_node_map IMPLEMENTATION.
  METHOD zif_excel_xml_named_node_map~create_iterator.
    DATA lo_ixml_node_iterator TYPE REF TO if_ixml_node_iterator.

    lo_ixml_node_iterator = ixml_named_node_map->create_iterator( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_iterator ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_node IMPLEMENTATION.
  METHOD zif_excel_xml_node~append_child.
    DATA lo_wrap_ixml_node TYPE REF TO lcl_wrap_ixml_node.
    DATA lo_ixml_node      TYPE REF TO if_ixml_node.

    lo_wrap_ixml_node ?= new_child.
    lo_ixml_node ?= lcl_wrap_ixml=>unwrap_ixml( lo_wrap_ixml_node ).
    ixml_node->append_child( lo_ixml_node ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~clone.
    DATA lo_ixml_node TYPE REF TO if_ixml_node.

    lo_ixml_node = ixml_node->clone( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~create_iterator.
    DATA lo_ixml_node_iterator TYPE REF TO if_ixml_node_iterator.

    lo_ixml_node_iterator = ixml_node->create_iterator( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_iterator ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_attributes.
    DATA lo_ixml_named_node_map TYPE REF TO if_ixml_named_node_map.

    lo_ixml_named_node_map = ixml_node->get_attributes( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_named_node_map ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_children.
    DATA lo_ixml_node_list TYPE REF TO if_ixml_node_list.

    lo_ixml_node_list = ixml_node->get_children( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_list ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_first_child.
    DATA lo_ixml_node TYPE REF TO if_ixml_node.

    lo_ixml_node = ixml_node->get_first_child( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_name.
    rval = ixml_node->get_name( ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_namespace_prefix.
    rval = ixml_node->get_namespace_prefix( ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_namespace_uri.
    rval = ixml_node->get_namespace_uri( ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_next.
    DATA lo_ixml_node TYPE REF TO if_ixml_node.

    lo_ixml_node = ixml_node->get_next( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~get_value.
    rval = ixml_node->get_value( ).
  ENDMETHOD.

  METHOD zif_excel_xml_node~set_value.
    ixml_node->set_value( value ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_node_collection IMPLEMENTATION.
  METHOD zif_excel_xml_node_collection~create_iterator.
    DATA lo_ixml_node_iterator TYPE REF TO if_ixml_node_iterator.

    lo_ixml_node_iterator = ixml_node_collection->create_iterator( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_iterator ).
  ENDMETHOD.

  METHOD zif_excel_xml_node_collection~get_length.
    rval = ixml_node_collection->get_length( ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_node_iterator IMPLEMENTATION.
  METHOD zif_excel_xml_node_iterator~get_next.
    DATA lo_ixml_node TYPE REF TO if_ixml_node.

    lo_ixml_node = ixml_node_iterator->get_next( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_node_list IMPLEMENTATION.
  METHOD zif_excel_xml_node_list~create_iterator.
    DATA lo_ixml_node_iterator TYPE REF TO if_ixml_node_iterator.

    lo_ixml_node_iterator = ixml_node_list->create_iterator( ).
    rval ?= lcl_wrap_ixml=>wrap_ixml( lo_ixml_node_iterator ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_ostream_string IMPLEMENTATION.

ENDCLASS.


CLASS lcl_wrap_ixml_parser IMPLEMENTATION.
  METHOD zif_excel_xml_parser~add_strip_space_element.
* Method PARSE_STRING of class ZCL_EXCEL_THEME_FMT_SCHEME
*    li_parser->add_strip_space_element( ).
    ixml_parser->add_strip_space_element( ).
  ENDMETHOD.

  METHOD zif_excel_xml_parser~parse.
    ixml_parser->parse( ).
  ENDMETHOD.

  METHOD zif_excel_xml_parser~set_normalizing.
    ixml_parser->set_normalizing( is_normalizing ).
  ENDMETHOD.

  METHOD zif_excel_xml_parser~set_validating.
    ixml_parser->set_validating( mode ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_renderer IMPLEMENTATION.
  METHOD zif_excel_xml_renderer~render.
    ixml_renderer->render( ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_stream IMPLEMENTATION.
  METHOD zif_excel_xml_stream~close.
    RAISE EXCEPTION TYPE lcx_unexpected.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_stream_factory IMPLEMENTATION.
  METHOD zif_excel_xml_stream_factory~create_istream_string.
* Method PARSE_STRING of class ZCL_EXCEL_THEME_FMT_SCHEME:
*    li_ixml = cl_ixml=>create( ).
*    li_document = li_ixml->create_document( ).
*    li_factory = li_ixml->create_stream_factory( ).
*    li_istream = li_factory->create_istream_string( iv_string ).

    DATA lo_wrap_ixml_istream_string TYPE REF TO lcl_wrap_ixml_istream_string.

    CREATE OBJECT lo_wrap_ixml_istream_string.
    lo_wrap_ixml_istream_string->lif_wrap_ixml_istream~ixml_istream = ixml_stream_factory->create_istream_string( string ).
    rval = lo_wrap_ixml_istream_string.
  ENDMETHOD.

  METHOD zif_excel_xml_stream_factory~create_istream_xstring.
* Method GET_IXML_FROM_ZIP_ARCHIVE of ZCL_EXCEL_READER_2007
*     lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).

    DATA lo_wrap_ixml_istream_xstring TYPE REF TO lcl_wrap_ixml_istream_xstring.

    CREATE OBJECT lo_wrap_ixml_istream_xstring.
    lo_wrap_ixml_istream_xstring->lif_wrap_ixml_istream~ixml_istream = ixml_stream_factory->create_istream_xstring(
                                                                           string ).
    rval = lo_wrap_ixml_istream_xstring.
  ENDMETHOD.

  METHOD zif_excel_xml_stream_factory~create_ostream_cstring.
* Method RENDER_XML_DOCUMENT of class ZCL_EXCEL_WRITER_2007:
*    lo_streamfactory = me->ixml->create_stream_factory( ).
*    lo_ostream = lo_streamfactory->create_ostream_cstring( string = lv_string ).
*    lo_renderer = me->ixml->create_renderer( ostream  = lo_ostream document = io_document ).
*    lo_renderer->render( ).

    DATA lo_wrap_ixml_ostream_string TYPE REF TO lcl_wrap_ixml_ostream_string.

    CREATE OBJECT lo_wrap_ixml_ostream_string.
    lo_wrap_ixml_ostream_string->lif_wrap_ixml_ostream~ixml_ostream = ixml_stream_factory->create_ostream_cstring(
                                                                          string->* ).
    rval = lo_wrap_ixml_ostream_string.
  ENDMETHOD.

  METHOD zif_excel_xml_stream_factory~create_ostream_xstring.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM:
*    CLEAR ep_content.
*    lo_ixml = cl_ixml=>create( ).
*    lo_streamfactory = lo_ixml->create_stream_factory( ).
*    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
*    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
*    lo_renderer->render( ).

    DATA lo_wrap_ixml_ostream_xstring TYPE REF TO lcl_wrap_ixml_ostream_xstring.

    CREATE OBJECT lo_wrap_ixml_ostream_xstring.
    lo_wrap_ixml_ostream_xstring->lif_wrap_ixml_ostream~ixml_ostream = ixml_stream_factory->create_ostream_xstring(
                                                                           string->* ).
    rval = lo_wrap_ixml_ostream_xstring.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_wrap_ixml_text IMPLEMENTATION.
ENDCLASS.


CLASS lcl_wrap_ixml_unknown IMPLEMENTATION.
  METHOD zif_excel_xml_unknown~query_interface.
    RAISE EXCEPTION TYPE lcx_unexpected.
  ENDMETHOD.
ENDCLASS.
