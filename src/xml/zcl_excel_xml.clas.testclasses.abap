*"* use this source file for your ABAP unit test classes

CLASS ltc_isxmlixml DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PROTECTED SECTION.

    TYPES tt_ixml_and_isxml TYPE STANDARD TABLE OF REF TO zif_excel_xml WITH DEFAULT KEY.

    CLASS-DATA isxml TYPE REF TO zif_excel_xml.
    CLASS-DATA ixml  TYPE REF TO zif_excel_xml.
    DATA attribute       TYPE REF TO zif_excel_xml_attribute.
    DATA document        TYPE REF TO zif_excel_xml_document.
    DATA element         TYPE REF TO zif_excel_xml_element.
    DATA encoding        TYPE REF TO zif_excel_xml_encoding.
    DATA istream         TYPE REF TO zif_excel_xml_istream.
    DATA ixml_or_isxml   TYPE REF TO zif_excel_xml.
    DATA ixml_and_isxml  TYPE tt_ixml_and_isxml.
    DATA length          TYPE i.
    DATA named_node_map  TYPE REF TO zif_excel_xml_named_node_map.
    DATA node            TYPE REF TO zif_excel_xml_node.
    DATA node_collection TYPE REF TO zif_excel_xml_node_collection.
    DATA node_iterator   TYPE REF TO zif_excel_xml_node_iterator.
    DATA node_list       TYPE REF TO zif_excel_xml_node_list.
    DATA ostream         TYPE REF TO zif_excel_xml_ostream.
    DATA parser          TYPE REF TO zif_excel_xml_parser.
    DATA rc              TYPE i.
    DATA ref_string      TYPE REF TO string.
    DATA ref_xstring     TYPE REF TO xstring.
    DATA renderer        TYPE REF TO zif_excel_xml_renderer.
    DATA stream_factory  TYPE REF TO zif_excel_xml_stream_factory.
    DATA string          TYPE string.
    DATA text            TYPE REF TO zif_excel_xml_text.
    DATA type            TYPE i.
    DATA value           TYPE string.
    DATA xstring         TYPE xstring.

    CLASS-METHODS get_ixml_and_isxml
      RETURNING
        VALUE(rt_result) TYPE tt_ixml_and_isxml.

    METHODS render
      RETURNING
        VALUE(rv_result) TYPE string.

  PRIVATE SECTION.

    METHODS create_encoding FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


"! Test of ZIF_EXCEL_IXML_DOCUMENT methods
CLASS ltc_isxmlixml_complex DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS reassign_to_other_parent FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


CLASS ltc_isxmlixml_document DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS create_element FOR TESTING RAISING cx_static_check.
    METHODS create_simple_element FOR TESTING RAISING cx_static_check.
    METHODS create_simple_element_ns FOR TESTING RAISING cx_static_check.
    METHODS find_from_name FOR TESTING RAISING cx_static_check.
    METHODS find_from_name_ns FOR TESTING RAISING cx_static_check.
    METHODS get_elements_by_tag_name FOR TESTING RAISING cx_static_check.
    METHODS get_elements_by_tag_name_ns FOR TESTING RAISING cx_static_check.
    METHODS get_root_element FOR TESTING RAISING cx_static_check.
    METHODS set_encoding FOR TESTING RAISING cx_static_check.
    METHODS set_standalone FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


"! Test of ZIF_EXCEL_IXML_ELEMENT methods
CLASS ltc_isxmlixml_element DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS find_from_name_default_ns FOR TESTING RAISING cx_static_check.
    METHODS find_from_name_level_1      FOR TESTING RAISING cx_static_check.
    METHODS find_from_name_level_2      FOR TESTING RAISING cx_static_check.
    METHODS find_from_name_ns_depth_0   FOR TESTING RAISING cx_static_check.
    METHODS find_from_name_ns_depth_1   FOR TESTING RAISING cx_static_check.
    METHODS find_from_name_ns_depth_2   FOR TESTING RAISING cx_static_check.
    METHODS get_attribute               FOR TESTING RAISING cx_static_check.
    METHODS get_attribute_node_ns       FOR TESTING RAISING cx_static_check.
    METHODS get_attribute_ns            FOR TESTING RAISING cx_static_check.
    METHODS get_elements_by_tag_name    FOR TESTING RAISING cx_static_check.
    METHODS get_elements_by_tag_name_ns FOR TESTING RAISING cx_static_check.
    METHODS remove_attribute_ns         FOR TESTING RAISING cx_static_check.
    METHODS set_attribute               FOR TESTING RAISING cx_static_check.
    METHODS set_attribute_ns            FOR TESTING RAISING cx_static_check.

    METHODS find_from_name_ns
      IMPORTING
        iv_depth     TYPE i
        iv_exp_found TYPE abap_bool
        iv_exp_value TYPE string OPTIONAL.
    METHODS setup.

ENDCLASS.


CLASS ltc_isxmlixml_named_node_map DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS create_iterator FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


"! Test of ZIF_EXCEL_IXML_NODE methods
CLASS ltc_isxmlixml_node DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS append_child FOR TESTING RAISING cx_static_check.
    METHODS clone FOR TESTING RAISING cx_static_check.
    METHODS create_iterator FOR TESTING RAISING cx_static_check.
    METHODS get_attributes FOR TESTING RAISING cx_static_check.
    METHODS get_children FOR TESTING RAISING cx_static_check.
    METHODS get_first_child FOR TESTING RAISING cx_static_check.
    METHODS get_name FOR TESTING RAISING cx_static_check.
    METHODS get_namespace_prefix FOR TESTING RAISING cx_static_check.
    METHODS get_namespace_uri FOR TESTING RAISING cx_static_check.
    METHODS get_next FOR TESTING RAISING cx_static_check.
    METHODS get_value FOR TESTING RAISING cx_static_check.
    METHODS set_value FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


CLASS ltc_isxmlixml_node_collection DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS create_iterator FOR TESTING RAISING cx_static_check.
    METHODS get_length FOR TESTING RAISING cx_static_check.

ENDCLASS.


CLASS ltc_xml_node_iterator DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS get_next FOR TESTING RAISING cx_static_check.

ENDCLASS.


CLASS ltc_isxmlixml_node_list DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS create_iterator FOR TESTING RAISING cx_static_check.

ENDCLASS.


CLASS ltc_isxmlixml_parse_and_render DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS create_docprops_app FOR TESTING RAISING cx_static_check.
    METHODS namespace FOR TESTING RAISING cx_static_check.
    METHODS space_normalizing_left_right FOR TESTING RAISING cx_static_check.
    METHODS space_normalizing_off FOR TESTING RAISING cx_static_check.
    METHODS space_normalizing_on FOR TESTING RAISING cx_static_check.
    METHODS space_normalizing_on_strip_on FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


CLASS ltc_isxmlixml_parser DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PUBLIC SECTION.

    INTERFACES lif_isxml_all_friends.

  PRIVATE SECTION.

    METHODS set_validating FOR TESTING RAISING cx_static_check.
    METHODS several_children FOR TESTING RAISING cx_static_check.
    METHODS space_normalizing_off FOR TESTING RAISING cx_static_check.
    METHODS space_normalizing_on FOR TESTING RAISING cx_static_check.
    METHODS space_normalizing_on_strip_on FOR TESTING RAISING cx_static_check.
    METHODS text_node FOR TESTING RAISING cx_static_check.
    METHODS two_ixml_instances FOR TESTING RAISING cx_static_check.
    METHODS two_ixml_stream_factories FOR TESTING RAISING cx_static_check.
    METHODS two_ixml_encodings FOR TESTING RAISING cx_static_check.
*    METHODS two_parsers FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


CLASS ltc_isxmlixml_render DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS most_simple_valid_xml FOR TESTING RAISING cx_static_check.
    METHODS namespace FOR TESTING RAISING cx_static_check.
    METHODS namespace_2 FOR TESTING RAISING cx_static_check.
    METHODS namespace_3 FOR TESTING RAISING cx_static_check.

    METHODS setup.
ENDCLASS.


CLASS ltc_isxmlixml_stream DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS close FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


CLASS ltc_isxmlixml_stream_factory DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS create_istream_string  FOR TESTING RAISING cx_static_check.
    METHODS create_istream_xstring FOR TESTING RAISING cx_static_check.
    METHODS create_ostream_cstring FOR TESTING RAISING cx_static_check.
    METHODS create_ostream_xstring FOR TESTING RAISING cx_static_check.

    METHODS parse
      IMPORTING
        io_ixml_or_isxml TYPE REF TO zif_excel_xml
        io_istream       TYPE REF TO zif_excel_xml_istream
      RETURNING
        VALUE(ro_result) TYPE REF TO zif_excel_xml_document.
    METHODS prepare
      IMPORTING
        io_ixml_or_isxml TYPE REF TO zif_excel_xml.
    METHODS setup.
    METHODS render_ostream
      IMPORTING
        io_ixml_or_isxml TYPE REF TO zif_excel_xml
        io_ostream       TYPE REF TO zif_excel_xml_ostream.

ENDCLASS.


CLASS ltc_isxmlonly_parser DEFINITION
      INHERITING FROM ltc_isxmlixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PUBLIC SECTION.

    INTERFACES lif_isxml_all_friends.

  PRIVATE SECTION.

    METHODS namespace FOR TESTING RAISING cx_static_check.

    METHODS setup.

ENDCLASS.


CLASS ltc_rewrite_xml_via_sxml DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS default_namespace FOR TESTING RAISING cx_static_check.
    METHODS default_namespace_removed FOR TESTING RAISING cx_static_check.
    METHODS namespace FOR TESTING RAISING cx_static_check.
    METHODS namespace_2 FOR TESTING RAISING cx_static_check.
    METHODS namespace_3 FOR TESTING RAISING cx_static_check.

    DATA parsed_element_index TYPE i.
    DATA string               TYPE string.

    METHODS get_expected_attribute
      IMPORTING
        iv_name          TYPE string
        iv_namespace     TYPE string DEFAULT ``
        iv_prefix        TYPE string DEFAULT ``
      RETURNING
        VALUE(rs_result) TYPE lcl_rewrite_xml_via_sxml=>ts_attribute.

    METHODS get_expected_element
      IMPORTING
        iv_name          TYPE string
        iv_namespace     TYPE string DEFAULT ``
        iv_prefix        TYPE string DEFAULT ``
      RETURNING
        VALUE(rs_result) TYPE lcl_rewrite_xml_via_sxml=>ts_element.

    METHODS get_expected_nsbinding
      IMPORTING
        iv_prefix        TYPE string DEFAULT ``
        iv_nsuri         TYPE string DEFAULT ``
      RETURNING
        VALUE(rs_result) TYPE lcl_rewrite_xml_via_sxml=>ts_nsbinding.

    METHODS get_parsed_element
      RETURNING
        VALUE(rs_result) TYPE lcl_rewrite_xml_via_sxml=>ts_element.

    METHODS get_parsed_element_attribute
      IMPORTING
        iv_index         TYPE i
      RETURNING
        VALUE(rs_result) TYPE lcl_rewrite_xml_via_sxml=>ts_attribute.

    METHODS get_parsed_element_nsbinding
      IMPORTING
        iv_index         TYPE i
      RETURNING
        VALUE(rs_result) TYPE lcl_rewrite_xml_via_sxml=>ts_nsbinding.

    METHODS rewrite_xml_via_sxml
      IMPORTING
        iv_xml_string TYPE string
      RETURNING
        VALUE(rv_string) TYPE string.

    METHODS set_current_parsed_element
      IMPORTING
        iv_index TYPE i.

ENDCLASS.


CLASS ltc_sxml_reader DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS bom FOR TESTING RAISING cx_static_check.
    METHODS empty_object_oriented_parsing FOR TESTING RAISING cx_static_check.
    METHODS empty_token_based_parsing FOR TESTING RAISING cx_static_check.
    METHODS empty_xml FOR TESTING RAISING cx_static_check.
    METHODS invalid_xml FOR TESTING RAISING cx_static_check.
    METHODS invalid_xml_eof_reached FOR TESTING RAISING cx_static_check.
    METHODS invalid_xml_not_wellformed FOR TESTING RAISING cx_static_check.
    METHODS keep_whitespace FOR TESTING RAISING cx_static_check.
    METHODS normalization FOR TESTING RAISING cx_static_check.
    METHODS object_oriented_parsing FOR TESTING RAISING cx_static_check.
    METHODS token_based_parsing FOR TESTING RAISING cx_static_check.
    METHODS xml_header_is_ignored FOR TESTING RAISING cx_static_check.

    DATA node         TYPE REF TO if_sxml_node.
    DATA reader       TYPE REF TO if_sxml_reader.
    DATA xstring      TYPE xstring.
    DATA parse_error  TYPE REF TO cx_sxml_parse_error.
    DATA error_rtti   TYPE REF TO cl_abap_typedescr.
    DATA open_element TYPE REF TO if_sxml_open_element.
    DATA error        TYPE REF TO cx_root.
    DATA node_attr    TYPE REF TO if_sxml_value.
    DATA value_node   TYPE REF TO if_sxml_value_node.
ENDCLASS.


CLASS ltc_sxml_writer DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS attribute_namespace FOR TESTING RAISING cx_static_check.
    "! It's mandatory to indicate nsuri = 'http://www.w3.org/XML/1998/namespace'
    "! for the standard "xml" namespace, but the URI is not rendered.
    "! e.g. open_element->set_attribute( name = 'space' nsuri = 'http://www.w3.org/XML/1998/namespace' prefix = 'xml' value = 'preserve' ).
    "! will render &lt;A xml:space="preserve"/>; as expected, there's no
    "! xmlns:xml="http://www.w3.org/XML/1998/namespace".
    METHODS attribute_xml_namespace FOR TESTING RAISING cx_static_check.
    METHODS most_simple_valid_xml FOR TESTING RAISING cx_static_check.
    METHODS namespace FOR TESTING RAISING cx_static_check.
    METHODS namespace_default FOR TESTING RAISING cx_static_check.
    METHODS namespace_default_by_attribute FOR TESTING RAISING cx_static_check.
    METHODS namespace_inheritance FOR TESTING RAISING cx_static_check.
    METHODS namespace_set_prefix FOR TESTING RAISING cx_static_check.
    METHODS object_oriented_rendering FOR TESTING RAISING cx_static_check.
    METHODS token_based_rendering FOR TESTING RAISING cx_static_check.
    METHODS write_namespace_declaration FOR TESTING RAISING cx_static_check.
    "! Order between namespace declarations, 1 then 2
    METHODS write_namespace_declaration_2 FOR TESTING RAISING cx_static_check.
    "! Order between namespace declarations, 2 then 1
    METHODS write_namespace_declaration_3 FOR TESTING RAISING cx_static_check.
    "! write_namespace_declaration called right before write_node( element )
    "! will position namespace declarations at the beginning of element,
    "! before default namespace and attributes
    METHODS write_namespace_declaration_4 FOR TESTING RAISING cx_static_check.
    "! write_namespace_declaration called right after write_node( element )
    "! will position namespace declarations at the end of element
    METHODS write_namespace_declaration_5 FOR TESTING RAISING cx_static_check.
    METHODS write_namespace_declaration_6 FOR TESTING RAISING cx_static_check.
    METHODS order_of_xmlns_and_attributes FOR TESTING RAISING cx_static_check.

    DATA open_element  TYPE REF TO if_sxml_open_element.
    DATA writer        TYPE REF TO if_sxml_writer.
    DATA string        TYPE string.
    DATA value_node    TYPE REF TO if_sxml_value_node.
    DATA close_element TYPE REF TO if_sxml_close_element.

    METHODS get_output
      IMPORTING
        io_writer        TYPE REF TO if_sxml_writer
      RETURNING
        VALUE(rv_result) TYPE string.

    METHODS setup.

ENDCLASS.


CLASS lth_isxmlixml DEFINITION.

  PUBLIC SECTION.

    CLASS-DATA document       TYPE REF TO zif_excel_xml_document       READ-ONLY.
    CLASS-DATA istream        TYPE REF TO zif_excel_xml_istream        READ-ONLY.
    CLASS-DATA ixml           TYPE REF TO zif_excel_xml                READ-ONLY.
    CLASS-DATA stream_factory TYPE REF TO zif_excel_xml_stream_factory READ-ONLY.

    CLASS-METHODS parse
      IMPORTING
        io_ixml_or_isxml TYPE REF TO zif_excel_xml
        iv_xml_string TYPE csequence
        iv_normalizing TYPE abap_bool DEFAULT abap_true
        iv_preserve_space_element TYPE abap_bool DEFAULT abap_true
        iv_validating TYPE i DEFAULT zif_excel_xml_parser=>co_no_validation
      RETURNING
        VALUE(ro_result) TYPE REF TO zif_excel_xml_document.

    CLASS-METHODS render
      IMPORTING
        ixml_or_isxml TYPE REF TO zif_excel_xml
      RETURNING
        VALUE(rv_result) TYPE string.

    CLASS-METHODS create_document.

ENDCLASS.


CLASS ltc_isxmlixml IMPLEMENTATION.
  METHOD create_encoding.
*ZCL_EXCEL_THEME
*    lo_encoding = lo_ixml->create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
*                                            character_set = 'UTF-8' ).
*    lo_document = lo_ixml->create_document( ).
*    lo_document->set_encoding( lo_encoding ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      encoding = ixml_or_isxml->create_encoding( byte_order    = zif_excel_xml_encoding=>co_platform_endian
                                                 character_set = 'UTF-8' ).
      document->set_encoding( encoding ).
      stream_factory = ixml_or_isxml->create_stream_factory( ).
      GET REFERENCE OF xstring INTO ref_xstring.
      ostream = stream_factory->create_ostream_xstring( ref_xstring ).
      renderer = ixml_or_isxml->create_renderer( ostream  = ostream
                                                 document = document ).
      document->create_simple_element( name   = 'é'
                                       parent = document ).
      CLEAR xstring.
      renderer->render( ).
      string = cl_abap_codepage=>convert_from( xstring ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<?xml version="1.0" encoding="utf-8"?><é/>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_ixml_and_isxml.
    TRY.
        ixml = zcl_excel_ixml=>create( ).
      CATCH cx_sy_dyn_call_illegal_class ##NO_HANDLER.
        " CL_IXML does not exist = ABAP Cloud
    ENDTRY.
    IF ixml IS BOUND.
      INSERT ixml INTO TABLE rt_result.
    ENDIF.
    isxml = zcl_excel_xml=>create( ).
    INSERT isxml INTO TABLE rt_result.
  ENDMETHOD.

  METHOD render.
    DATA lr_string   TYPE REF TO string.
    DATA lo_ostream  TYPE REF TO zif_excel_xml_ostream.
    DATA lo_renderer TYPE REF TO zif_excel_xml_renderer.

    stream_factory = ixml_or_isxml->create_stream_factory( ).
    GET REFERENCE OF rv_result INTO lr_string.
    lo_ostream = stream_factory->create_ostream_cstring( lr_string ).
    lo_renderer = ixml_or_isxml->create_renderer( ostream  = lo_ostream
                                                  document = document ).
    document->set_declaration( abap_false ).
    " Fills RV_RESULT
    lo_renderer->render( ).
    " remove the UTF-16 BOM (i.e. remove the first character)
    SHIFT rv_result LEFT BY 1 PLACES.

    " Normalize XML according to SXML limitations in order to compare IXML
    " and SXML results by simple string comparison.
    rv_result = lcl_rewrite_xml_via_sxml=>execute( rv_result ).
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_complex IMPLEMENTATION.
  METHOD reassign_to_other_parent.
    " As done in method CREATE_DOCPROPS_APP of class ZCL_EXCEL_WRITER_2007, e.g. LinksUpToDate and SharedDoc.

    DATA lo_element_root TYPE REF TO zif_excel_xml_element.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A/>' ).
      lo_element_root = document->get_root_element( ).
      element = document->create_simple_element( name   = 'LinksUpToDate'
                                                 parent = document ).
      lo_element_root->append_child( new_child = element ).
      element = document->create_simple_element( name   = 'SharedDoc'
                                                 parent = document ).
      lo_element_root->append_child( new_child = element ).
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = '<A><LinksUpToDate/><SharedDoc/></A>' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_document IMPLEMENTATION.
  METHOD create_element.
* (only at 2 places in test classes)
* Method SET_CELL of local class LTC_COLUMN_FORMULA of class ZCL_EXCEL_WRITER_2007.
*     lo_cell = lo_document->create_element( 'c' ).
*     lo_cell->set_attribute( name = 'r' value = |R{ is_cell_data-cell_row }C{ is_cell_data-cell_column }| ).
*     lo_root->append_child( lo_cell ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      element = document->create_element( name = 'A' ).
      document->append_child( element ).
      string = render( ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A/>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create_simple_element.
* Method ADD_HYPERLINKS of local class LCL_CREATE_XL_SHEET of class ZCL_EXCEL_WRITER_2007.
*      lo_element = o_document->create_simple_element( name   = 'hyperlinks'
*                                                      parent = o_document ).
*        lo_element_2 = o_document->create_simple_element( name   = 'hyperlink'
*                                                          parent = lo_element ).
*        lo_element_2->set_attribute_ns( name  = 'ref'
*                                        value = lv_value ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      document->create_simple_element( name   = 'A'
                                       parent = document ).
      string = render( ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A/>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create_simple_element_ns.
* Method WRITE_THEME of class ZCL_EXCEL_THEME
*    CONSTANTS c_theme TYPE string VALUE 'theme'.            "#EC NOTEXT
*    CONSTANTS c_theme_xmlns TYPE string VALUE 'xmlns:a'.    "#EC NOTEXT
*    CONSTANTS c_theme_prefix TYPE string VALUE 'a'.         "#EC NOTEXT
*    CONSTANTS c_theme_xmlns_val TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main'. "#EC NOTEXT
*    lo_document->set_namespace_prefix( prefix = 'a' ).
*    lo_element_root = lo_document->create_simple_element_ns( prefix = c_theme_prefix
*                                                             name   = c_theme
*                                                             parent = lo_document ).
*    lo_element_root->set_attribute_ns( name  = c_theme_xmlns
*                                       value = c_theme_xmlns_val ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      element = document->create_simple_element_ns( name   = 'A'
                                                    parent = document
                                                    prefix = 'a' ).
      element->set_attribute_ns( name  = 'xmlns:a'
                                 value = 'http://schemas.openxmlformats.org/drawingml/2006/main' ).
      string = render( ).
      cl_abap_unit_assert=>assert_equals(
          act = string
          exp = `<a:A xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD find_from_name.
*Method LOAD_WORKSHEET_TABLES of ZCL_EXCEL_READER_2007
*    lo_ixml_table_style ?= lo_ixml_table->find_from_name( 'tableStyleInfo' ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = |<A><B/></A>| ).
      element = document->find_from_name( name = 'B' ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'B' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD find_from_name_ns.
* Method LOAD_CHART_ATTRIBUTES of class ZCL_EXCEL_DRAWING
*     CONSTANTS: BEGIN OF namespace,
*                  c   TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
*                END OF namespace.
*     node ?= ip_chart->if_ixml_node~get_first_child( ).
*     node2 ?= node->find_from_name_ns( name = 'date1904' uri = namespace-c ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = |<A><B>B1</B><a:B xmlns:a="a">B2</a:B></A>| ).
      element = document->find_from_name_ns( name = 'B'
                                             uri  = 'a' ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'B' ).
      cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                          exp = 'B2' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_elements_by_tag_name.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM
*    lo_collection = lo_document->get_elements_by_tag_name( 'Override' ).

    DATA lo_isxml_node_collection TYPE REF TO zif_excel_xml_node_collection.
    DATA lo_isxml_node_iterator   TYPE REF TO zif_excel_xml_node_iterator.
    DATA lo_isxml_node            TYPE REF TO zif_excel_xml_node.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
                     io_ixml_or_isxml = ixml_or_isxml
                     iv_xml_string    = |<A><B xmlns="dnsuri">B1</B><a:B xmlns:a="a">B2</a:B><B>B3</B></A>| ).
      lo_isxml_node_collection = document->get_elements_by_tag_name( 'B' ).
      lo_isxml_node_iterator = lo_isxml_node_collection->create_iterator( ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B1' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B3' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = lo_isxml_node ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_elements_by_tag_name_ns.
* Method LOAD_WORKSHEET of class ZCL_EXCEL_READER_2007
*    lo_ixml_rows = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'row' uri = namespace-main ).

    DATA lo_isxml_node_collection TYPE REF TO zif_excel_xml_node_collection.
    DATA lo_isxml_node_iterator   TYPE REF TO zif_excel_xml_node_iterator.
    DATA lo_isxml_node            TYPE REF TO zif_excel_xml_node.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
          io_ixml_or_isxml = ixml_or_isxml
          iv_xml_string    = |<A xmlns:a="nsuri"><B>B1</B><a:B>B2</a:B><C xmlns="dnsuri"><B>B3</B></C></A>| ).

      lo_isxml_node_collection = document->get_elements_by_tag_name_ns( name = 'B' ).
      lo_isxml_node_iterator = lo_isxml_node_collection->create_iterator( ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B1' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = lo_isxml_node ).

      lo_isxml_node_collection = document->get_elements_by_tag_name_ns( name = 'B'
                                                                        uri  = 'nsuri' ).
      lo_isxml_node_iterator = lo_isxml_node_collection->create_iterator( ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B2' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = lo_isxml_node ).

      lo_isxml_node_collection = document->get_elements_by_tag_name_ns( name = 'B'
                                                                        uri  = 'dnsuri' ).
      lo_isxml_node_iterator = lo_isxml_node_collection->create_iterator( ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B3' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = lo_isxml_node ).

    ENDLOOP.
  ENDMETHOD.

  METHOD get_root_element.
* Method READ_THEME of class ZCL_EXCEL_THEME:
*    lo_node_theme  = io_theme_xml->get_root_element( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = |<A><B>B1</B></A>| ).
      element = document->get_root_element( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD set_encoding.
* Method CREATE_XML_DOCUMENT of class ZCL_EXCEL_WRITER_2007:
*    DATA lo_encoding TYPE REF TO if_ixml_encoding.
*    lo_encoding = me->ixml->create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
*                                             character_set = 'utf-8' ).
*    ro_document = me->ixml->create_document( ).
*    ro_document->set_encoding( lo_encoding ).

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      encoding = ixml_or_isxml->create_encoding( byte_order    = zif_excel_xml_encoding=>co_platform_endian
                                                 character_set = 'utf-8' ).
      document->set_encoding( encoding ).
      element = document->create_simple_element( name   = 'ROOT'
                                                 parent = document ).
      stream_factory = ixml_or_isxml->create_stream_factory( ).
      GET REFERENCE OF xstring INTO ref_xstring.
      ostream = stream_factory->create_ostream_xstring( ref_xstring ).
      renderer = ixml_or_isxml->create_renderer( ostream  = ostream
                                                 document = document ).
      CLEAR xstring.
      renderer->render( ).
      string = cl_abap_codepage=>convert_from( xstring ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = '<?xml version="1.0" encoding="utf-8"?><ROOT/>' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD set_standalone.
* Method CREATE_XML_DOCUMENT of class ZCL_EXCEL_WRITER_2007:
*    DATA lo_encoding TYPE REF TO if_ixml_encoding.
*    lo_encoding = me->ixml->create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
*                                             character_set = 'utf-8' ).
*    ro_document = me->ixml->create_document( ).
*    ro_document->set_encoding( lo_encoding ).
*    ro_document->set_standalone( abap_true ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      encoding = ixml_or_isxml->create_encoding( byte_order    = zif_excel_xml_encoding=>co_platform_endian
                                                 character_set = 'utf-8' ).
      document->set_encoding( encoding ).
      document->set_standalone( abap_true ).
      element = document->create_simple_element( name   = 'ROOT'
                                                 parent = document ).
      stream_factory = ixml_or_isxml->create_stream_factory( ).
      GET REFERENCE OF xstring INTO ref_xstring.
      ostream = stream_factory->create_ostream_xstring( ref_xstring ).
      renderer = ixml_or_isxml->create_renderer( ostream  = ostream
                                                 document = document ).
      CLEAR xstring.
      renderer->render( ).
      string = cl_abap_codepage=>convert_from( xstring ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><ROOT/>' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_element IMPLEMENTATION.
  METHOD find_from_name_default_ns.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
          io_ixml_or_isxml = ixml_or_isxml
          iv_xml_string    = |<A><C xmlns="dnsuri" xmlns:b="nsuri_b"><b:B>B1</b:B><B>B2</B></C><B>B3</B><B xmlns="dnsuri">B4</B></A>| ).
      element = document->get_root_element( ).
      element = element->find_from_name( name = 'B' ).
      cl_abap_unit_assert=>assert_bound( act = element ).
      cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                          exp = 'B2' ).

      element = document->get_root_element( ).
      element = element->find_from_name_ns( name = 'B' ).
      cl_abap_unit_assert=>assert_bound( act = element ).
      cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                          exp = 'B3' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD find_from_name_level_1.
* Method LOAD_WORKSHEET_TABLES of class ZCL_EXCEL_READER_2007:
*    DATA lo_ixml_table TYPE REF TO if_ixml_element.
*      lo_ixml_table_style ?= lo_ixml_table->find_from_name( 'tableStyleInfo' ).

    DATA lo_element TYPE REF TO zif_excel_xml_element.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = |<A xmlns:a="a"><C>C1</C><B><a:C>C2</a:C><C>C3</C></B></A>| ).
      lo_element = document->get_root_element( ).
      lo_element ?= lo_element->get_first_child( ).
      lo_element ?= lo_element->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_element->get_name( )
                                          exp = 'B' ).
      element = lo_element->find_from_name( name = 'C' ).
      cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                          exp = 'C3' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD find_from_name_level_2.
* Method LOAD_WORKSHEET_TABLES of class ZCL_EXCEL_READER_2007:
*    DATA lo_ixml_table TYPE REF TO if_ixml_element.
*      lo_ixml_table_style ?= lo_ixml_table->find_from_name( 'tableStyleInfo' ).

    DATA lo_element TYPE REF TO zif_excel_xml_element.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
                     io_ixml_or_isxml = ixml_or_isxml
                     iv_xml_string    = |<A xmlns:a="a"><C>C1</C><B><a:C>C2</a:C><D><C>C3</C></D></B></A>| ).
      lo_element = document->get_root_element( ).
      lo_element ?= lo_element->get_first_child( ).
      lo_element ?= lo_element->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_element->get_name( )
                                          exp = 'B' ).
      element = lo_element->find_from_name( name = 'C' ).
      cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                          exp = 'C3' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD find_from_name_ns.
* Method LOAD_CHART_ATTRIBUTES of class ZCL_EXCEL_DRAWING:
*    DATA: node                TYPE REF TO if_ixml_element.
*    node2 ?= node->find_from_name_ns( name = 'date1904' uri = namespace-c ).
*        node2 ?= node->find_from_name_ns( name = 'marker' uri = namespace-c depth = '1' ).

    DATA lo_element TYPE REF TO zif_excel_xml_element.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
                     io_ixml_or_isxml = ixml_or_isxml
                     iv_xml_string    = |<A xmlns:pc="uc"><B><C>C1</C><D><C>C2</C><pc:C>C3</pc:C></D></B></A>| ).
      lo_element = document->get_root_element( ).
      lo_element ?= lo_element->get_first_child( ).
      cl_abap_unit_assert=>assert_equals( act = lo_element->get_name( )
                                          exp = 'B' ).
      element = lo_element->find_from_name_ns( depth = iv_depth
                                               name  = 'C'
                                               uri   = 'uc' ).
      IF iv_exp_found = abap_false.
        cl_abap_unit_assert=>assert_not_bound( element ).
      ELSE.
        cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                            exp = iv_exp_value ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD find_from_name_ns_depth_0.
    find_from_name_ns( iv_depth     = 0
                       iv_exp_found = abap_true
                       iv_exp_value = 'C3' ).
  ENDMETHOD.

  METHOD find_from_name_ns_depth_1.
    find_from_name_ns( iv_depth     = 1
                       iv_exp_found = abap_false ).
  ENDMETHOD.

  METHOD find_from_name_ns_depth_2.
    find_from_name_ns( iv_depth     = 2
                       iv_exp_found = abap_true
                       iv_exp_value = 'C3' ).
  ENDMETHOD.

  METHOD get_attribute.
    " Method LOAD_STYLE_BORDERS of class ZCL_EXCEL_READER_2007:
    "      IF lo_node_border->get_attribute( 'diagonalDown' ) IS NOT INITIAL.
    " Method READ_THEME of class ZCL_EXCEL_THEME:
    "    CONSTANTS c_theme_xmlns TYPE string VALUE 'xmlns:a'.    "#EC NOTEXT
    "      xmls_a = lo_node_theme->get_attribute( name = c_theme_xmlns ).
    " NB: above GET_ATTRIBUTE with NAME = 'xmlns:XXXX' always returns no attribute found, it's also not used
    "     so cleanup done -> https://github.com/abap2xlsx/abap2xlsx/pull/1183.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
                     io_ixml_or_isxml = ixml_or_isxml
                     iv_xml_string    = `<A xmlns="default" xmlns:nsprefix="nsuri" nsprefix:attr="A1" attr="A2"/>` ).
      element = document->get_root_element( ).
      string = element->get_attribute( name = 'xmlns:nsprefix' ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = '' ).
      string = element->get_attribute( name = 'nsprefix:attr' ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = '' ).
      string = element->get_attribute( name = 'attr' ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = 'A2' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_attribute_node_ns.
* Method LOAD_COMMENTS of class ZCL_EXCEL_READER_2007:
*      lo_attr = lo_node_comment->get_attribute_node_ns( name = 'ref' ).
*      lv_attr_value  = lo_attr->get_value( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A xmlns:nsprefix="nsuri" nsprefix:attr="A1" attr="A2"/>` ).
      element = document->get_root_element( ).
      attribute = element->get_attribute_node_ns( name = 'attr'
                                                  uri  = 'nsuri' ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_value( )
                                          exp = 'A1' ).
      attribute = element->get_attribute_node_ns( name = 'attr' ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_value( )
                                          exp = 'A2' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_attribute_ns.
* Method LOAD_WORKSHEET_HYPERLINKS of class ZCL_EXCEL_READER_2007:
*     ls_hyperlink-tooltip  = lo_ixml_hyperlink->get_attribute_ns( 'tooltip' ).
*     ls_hyperlink-r_id     = lo_ixml_hyperlink->get_attribute_ns( name = 'id' uri = namespace-r ).

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A xmlns:nsprefix="nsuri" nsprefix:attr="A1" attr="A2"/>` ).
      element = document->get_root_element( ).
      string = element->get_attribute_ns( name = 'nsprefix'
                                          uri  = 'nsuri' ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = '' ).
      string = element->get_attribute_ns( name = 'attr'
                                          uri  = 'nsuri' ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = 'A1' ).
      string = element->get_attribute_ns( name = 'attr' ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = 'A2' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_elements_by_tag_name.
* Method LOAD_WORKSHEET_TABLES of class ZCL_EXCEL_READER_2007:
*      lo_ixml_table_columns =  lo_ixml_table->get_elements_by_tag_name( name = 'tableColumn' ).

    DATA lo_isxml_node_collection TYPE REF TO zif_excel_xml_node_collection.
    DATA lo_isxml_node_iterator   TYPE REF TO zif_excel_xml_node_iterator.
    DATA lo_isxml_node            TYPE REF TO zif_excel_xml_node.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
          io_ixml_or_isxml = ixml_or_isxml
          iv_xml_string    = |<A><B>B1</B><a:B xmlns:a="a">B2</a:B><B>B3</B><C><B>B4</B><a:B xmlns:a="a">B5</a:B><B xmlns="dnsuri">B6</B></C></A>| ).

      element = document->get_root_element( ).
      lo_isxml_node_collection = element->get_elements_by_tag_name( 'B' ).
      lo_isxml_node_iterator = lo_isxml_node_collection->create_iterator( ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B1' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B3' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B4' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B6' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = lo_isxml_node ).

      element = document->find_from_name( name = 'C' ).
      lo_isxml_node_collection = element->get_elements_by_tag_name( 'B' ).
      lo_isxml_node_iterator = lo_isxml_node_collection->create_iterator( ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B4' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B6' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = lo_isxml_node ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_elements_by_tag_name_ns.
* Method LOAD_DXF_STYLES of class ZCL_EXCEL_READER_2007:
*    lo_nodes_dxf ?= lo_node_dxfs->get_elements_by_tag_name_ns( name = 'dxf' uri = namespace-main ).

    DATA lo_isxml_node_collection TYPE REF TO zif_excel_xml_node_collection.
    DATA lo_isxml_node_iterator   TYPE REF TO zif_excel_xml_node_iterator.
    DATA lo_isxml_node            TYPE REF TO zif_excel_xml_node.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
                     io_ixml_or_isxml = ixml_or_isxml
                     iv_xml_string    = |<A xmlns:a="nsa"><B>B1</B><a:B>B2</a:B><B>B3</B><a:B>B4</a:B></A>| ).
      element = document->get_root_element( ).
      lo_isxml_node_collection = element->get_elements_by_tag_name_ns( name = 'B'
                                                                       uri  = 'nsa' ).
      lo_isxml_node_iterator = lo_isxml_node_collection->create_iterator( ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B2' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_isxml_node->get_value( )
                                          exp = 'B4' ).
      lo_isxml_node = lo_isxml_node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = lo_isxml_node ).
    ENDLOOP.
  ENDMETHOD.

  METHOD remove_attribute_ns.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM:
*        lo_element->remove_attribute_ns( lc_xml_attr_contenttype ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
                     io_ixml_or_isxml = ixml_or_isxml
                     iv_xml_string    = `<A xmlns="nsuri" xmlns:nsprefix="nsuri2" nsprefix:attr="A1" attr="A2"/>` ).
      element = document->get_root_element( ).
      element->remove_attribute_ns( name = 'attr' ).
      string = render( ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A nsprefix:attr="A1" xmlns="nsuri" xmlns:nsprefix="nsuri2"/>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD set_attribute.
* Method CREATE_XL_SHAREDSTRINGS of class ZCL_EXCEL_WRITER_2007:
*           lo_sub_element->set_attribute( name = 'space' namespace = 'xml' value = 'preserve' ).
* Method CREATE_XL_SHEET_COLUMN_FORMULA of class ZCL_EXCEL_WRITER_2007:
*      eo_element->set_attribute( name  = 't'
*                                 value = <ls_column_formula_used>-t ).

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      element = document->create_simple_element( name   = 'A'
                                                 parent = document ).
      element->set_attribute( name  = 'a'
                              value = '1' ).
      element->set_attribute( name      = 'space'
                              namespace = 'xml'
                              value     = 'preserve' ).
      string = render( ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A a="1" xml:space="preserve"/>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD set_attribute_ns.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM:
*        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
*                                      value = lc_xml_node_workb_ct ).
* Method CLONE_IXML_WITH_NAMESPACES of class ZCL_EXCEL_COMMON:
*      result->set_attribute_ns( prefix = 'xmlns' name = <xmlns>-name value = <xmlns>-value ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      element = document->create_simple_element_ns( name   = 'A'
                                                    parent = document
                                                    prefix = 'a' ).
      element->set_attribute_ns( name  = 'xmlns:a'
                                 value = 'nsuri_a' ).
      element = document->create_simple_element_ns( name   = 'B'
                                                    parent = element
                                                    prefix = 'b' ).
      element->set_attribute_ns( name   = 'b'
                                 prefix = 'xmlns'
                                 value  = 'nsuri_b' ).
      string = render( ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<a:A xmlns:a="nsuri_a"><b:B xmlns:b="nsuri_b"/></a:A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_named_node_map IMPLEMENTATION.
  METHOD create_iterator.
* Method FILL_STRUCT_FROM_ATTRIBUTES of class ZCL_EXCEL_READER_2007:
*    lo_attributes  = ip_element->get_attributes( ).
*    lo_iterator    = lo_attributes->create_iterator( ).
*    lo_attribute  ?= lo_iterator->get_next( ).
*    WHILE lo_attribute IS BOUND.
*      lo_attribute ?= lo_iterator->get_next( ).
*    ENDWHILE.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A a="1" b="2"/>` ).
      element ?= document->get_root_element( ).
      named_node_map = element->get_attributes( ).
      node_iterator = named_node_map->create_iterator( ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_name( )
                                          exp = `a` ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_value( )
                                          exp = `1` ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_name( )
                                          exp = `b` ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_value( )
                                          exp = `2` ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = attribute ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = attribute ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_node IMPLEMENTATION.
  METHOD append_child.
* Method ADD_1_VAL_CHILD_NODE of class ZCL_EXCEL_WRITER_2007:
*    io_parent->append_child( new_child = lo_child ).

    DATA lo_element TYPE REF TO zif_excel_xml_element.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<ROOT/>' ).
      element = document->get_root_element( ).
      lo_element = document->create_element( name = 'A' ).
      element->append_child( lo_element ).
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<ROOT><A/></ROOT>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD clone.
* Method CLONE_IXML_WITH_NAMESPACES of class ZCL_EXCEL_COMMON:
*    result ?= element->clone( ).

    DATA lo_element TYPE REF TO zif_excel_xml_element.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<ROOT><A><B/></A></ROOT>` ).
      element ?= document->get_root_element( ).
      lo_element ?= element->get_first_child( )->clone( ).
      element->append_child( lo_element ).
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<ROOT><A><B/></A><A><B/></A></ROOT>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create_iterator.
* Method CLONE_IXML_WITH_NAMESPACES of class ZCL_EXCEL_COMMON:
*    iterator = element->create_iterator( ).
*    node = iterator->get_next( ).
*    WHILE node IS BOUND.
*      node = iterator->get_next( ).
*    ENDWHILE.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A><B><C/></B><D/></A>` ).
      element ?= document->get_root_element( ).
      node_iterator = element->create_iterator( ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `A` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `B` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `C` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `D` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = element ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = element ).

      element ?= document->get_root_element( )->get_first_child( )->get_first_child( ).
      node_iterator = element->create_iterator( ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `C` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = element ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_attributes.
* Method FILL_STRUCT_FROM_ATTRIBUTES of class ZCL_EXCEL_READER_2007:
*    lo_attributes  = ip_element->get_attributes( ).
*    lo_iterator    = lo_attributes->create_iterator( ).
*    lo_attribute  ?= lo_iterator->get_next( ).
*    WHILE lo_attribute IS BOUND.
*      lo_attribute ?= lo_iterator->get_next( ).
*    ENDWHILE.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A a="1" b="2"/>` ).
      element ?= document->get_root_element( ).
      named_node_map = element->get_attributes( ).
      node_iterator = named_node_map->create_iterator( ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_name( )
                                          exp = `a` ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_value( )
                                          exp = `1` ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_name( )
                                          exp = `b` ).
      cl_abap_unit_assert=>assert_equals( act = attribute->get_value( )
                                          exp = `2` ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = attribute ).
      attribute ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = attribute ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_children.
* Method READ_THEME of class ZCL_EXCEL_THEME:
*      lo_theme_children = lo_node_theme->get_children( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A><B><C/></B><D/></A>' ).
      " WHEN
      element = document->get_root_element( ).
      node_list = element->get_children( ).
      node_iterator = node_list->create_iterator( ).
      " THEN
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `B` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `D` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = element ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = element ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_first_child.
* Method LOAD_COMMENTS of class ZCL_EXCEL_READER_2007:
*      lo_node_comment_child ?= lo_node_comment->get_first_child( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A><B/></A>' ).
      " WHEN
      element ?= document->get_first_child( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `A` ).
      " WHEN
      element ?= element->get_first_child( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `B` ).
      " WHEN
      element ?= element->get_first_child( ).
      " THEN
      cl_abap_unit_assert=>assert_not_bound( act = element ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_name.
* Method FILL_STRUCT_FROM_ATTRIBUTES of class ZCL_EXCEL_READER_2007:
*      lv_name = lo_attribute->get_name( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A a="1"/>' ).
      element ?= document->get_root_element( ).
      node = element->get_attribute_node_ns( name = 'a' ).
      " WHEN
      string = node->get_name( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `a` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_namespace_prefix.
* Method CLONE_IXML_WITH_NAMESPACES of class ZCL_EXCEL_COMMON:
*      xmlns-name = node->get_namespace_prefix( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<a:A xmlns:a="nsuri"/>' ).
      node = document->get_root_element( ).
      " WHEN
      string = node->get_namespace_prefix( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `a` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_namespace_uri.
* Method CLONE_IXML_WITH_NAMESPACES of class ZCL_EXCEL_COMMON:
*      xmlns-value = node->get_namespace_uri( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<a:A xmlns:a="nsuri"/>' ).
      node = document->get_root_element( ).
      " WHEN
      string = node->get_namespace_uri( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `nsuri` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_next.
* Method LOAD_WORKBOOK of class ZCL_EXCEL_READER_2007:
*      lo_node ?= lo_node->get_next( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A><B><C/></B><D><E/></D><F/></A>' ).
      node = document->get_root_element( )->get_first_child( ).
      " WHEN
      node = node->get_next( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = node->get_name( )
                                          exp = `D` ).
      " WHEN
      node = node->get_next( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = node->get_name( )
                                          exp = `F` ).
      " WHEN
      node = node->get_next( ).
      " THEN
      cl_abap_unit_assert=>assert_not_bound( act = node ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_value.
* Method FILL_STRUCT_FROM_ATTRIBUTES of class ZCL_EXCEL_READER_2007:
*        <component> = lo_attribute->get_value( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A>1<B>2</B>3</A>' ).
      " WHEN
      node = document.
      value = node->get_value( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = value
                                          exp = '' ).
      " WHEN
      node = document->get_root_element( ).
      value = node->get_value( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = value
                                          exp = '123' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD set_value.
* Method CREATE_XL_WORKBOOK of class ZCL_EXCEL_WRITER_2007
*        lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
*                                                                parent = lo_document ).
*        lo_sub_element->set_value( value = lv_value ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      element = document->create_simple_element_ns( name   = 'A'
                                                    parent = document ).
      element->set_value( '1' ).
      string = render( ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A>1</A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_node_collection IMPLEMENTATION.
  METHOD create_iterator.
* Method LOAD_WORKSHEET_DRAWING of class ZCL_EXCEL_READER_2007:
*    coll_length = anchors->get_length( ).
*    iterator = anchors->create_iterator( ).
*    DO coll_length TIMES.
*      anchor_elem ?= iterator->get_next( ).
*    ENDDO.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = |<A><B>B1</B><a:B xmlns:a="a">B2</a:B><B>B3</B></A>| ).
      node_collection = document->get_elements_by_tag_name( 'B' ).
      length = node_collection->get_length( ).
      cl_abap_unit_assert=>assert_equals( act = length
                                          exp = 2 ).
      node_iterator = node_collection->create_iterator( ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = node->get_value( )
                                          exp = 'B1' ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = node->get_value( )
                                          exp = 'B3' ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = node ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_length.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM
*    lo_collection = lo_document->get_elements_by_tag_name( 'Override' ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = |<A><B>B1</B><a:B xmlns:a="a">B2</a:B><B>B3</B></A>| ).
      node_collection = document->get_elements_by_tag_name( 'B' ).
      length = node_collection->get_length( ).
      cl_abap_unit_assert=>assert_equals( act = length
                                          exp = 2 ).
      node_iterator = node_collection->create_iterator( ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = node->get_value( )
                                          exp = 'B1' ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = node->get_value( )
                                          exp = 'B3' ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = node ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_xml_node_iterator IMPLEMENTATION.
  METHOD get_next.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM
*    lo_collection = lo_document->get_elements_by_tag_name( 'Override' ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = |<A><B>B1</B><a:B xmlns:a="a">B2</a:B><B>B3</B></A>| ).
      node_collection = document->get_elements_by_tag_name( 'B' ).
      node_iterator = node_collection->create_iterator( ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = node->get_value( )
                                          exp = 'B1' ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = node->get_value( )
                                          exp = 'B3' ).
      node = node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = node ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_node_list IMPLEMENTATION.
  METHOD create_iterator.
* Method READ_THEME of class ZCL_EXCEL_THEME:
*      lo_theme_children = lo_node_theme->get_children( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A><B><C/></B><D/></A>' ).
      " WHEN
      element = document->get_root_element( ).
      node_list = element->get_children( ).
      node_iterator = node_list->create_iterator( ).
      " THEN
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `B` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `D` ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = element ).
      element ?= node_iterator->get_next( ).
      cl_abap_unit_assert=>assert_not_bound( act = element ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_parse_and_render IMPLEMENTATION.
  METHOD create_docprops_app.
    DATA lv_xml_string TYPE string.

    " XML generated by method CREATE_DOCPROPS_APP of class ZCL_EXCEL_WRITER_2007
    " GIVEN
    lv_xml_string = ``
&& `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Ex`
&& `cel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPair`
&& `s><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion`
&& `></Properties>`.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = lv_xml_string ).
      " WHEN
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = lv_xml_string ).
    ENDLOOP.
  ENDMETHOD.

  METHOD namespace.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse(
          io_ixml_or_isxml = ixml_or_isxml
          iv_xml_string    = `<nsprefix:A xmlns="dnsuri" xmlns:nsprefix="nsuri" nsprefix:attr="1" attr="2"><B attr="3"/></nsprefix:A>` ).
      " WHEN
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      " THEN
      cl_abap_unit_assert=>assert_equals(
          act = string
          exp = `<nsprefix:A nsprefix:attr="1" attr="2" xmlns="dnsuri" xmlns:nsprefix="nsuri"><B attr="3"/></nsprefix:A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.

  METHOD space_normalizing_left_right.
* Method GET_IXML_FROM_ZIP_ARCHIVE of class ZCL_EXCEL_READER_2007:
*    lo_comments_xml = me->get_ixml_from_zip_archive( ip_path ).
*    METHODS get_ixml_from_zip_archive
*      IMPORTING
*        !i_filename     TYPE string
*        !is_normalizing TYPE abap_bool DEFAULT 'X'
*  METHOD get_ixml_from_zip_archive.
*    lo_ixml           = cl_ixml=>create( ).
*    lo_streamfactory  = lo_ixml->create_stream_factory( ).
*    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
*    r_ixml            = lo_ixml->create_document( ).
*    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
*                                                istream        = lo_istream
*                                                document       = r_ixml ).
*    lo_parser->set_normalizing( is_normalizing ).
*    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
*    lo_parser->parse( ).
* All calls don't pass IS_NORMALIZING except the call in the method LOAD_SHARED_STRINGS of ZCL_EXCEL_READER_2007:
*    lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
*                                                           is_normalizing = space ).  " NO!!! normalizing - otherwise leading blanks will be omitted and that is not really desired for the stringtable
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A>  <B>  1  </B>  </A>`
                                       iv_normalizing   = abap_true ).
      " WHEN
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A><B>1</B></A>` ).
    ENDLOOP.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A>  <B>  1  </B>  </A>`
                                       iv_normalizing   = abap_false ).
      " WHEN
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A><B>  1  </B></A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD space_normalizing_off.
* Method GET_IXML_FROM_ZIP_ARCHIVE of class ZCL_EXCEL_READER_2007:
*    lo_ixml           = cl_ixml=>create( ).
*    lo_streamfactory  = lo_ixml->create_stream_factory( ).
*    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
*    r_ixml            = lo_ixml->create_document( ).
*    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
*                                                istream        = lo_istream
*                                                document       = r_ixml ).
*    lo_parser->set_normalizing( is_normalizing ).
*    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
*    lo_parser->parse( ).
* ALL CALLS ARE DONE WITH is_normalizing = 'X' except this one in method LOAD_SHARED_STRINGS of ZCL_EXCEL_READER_2007:
*    lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
*                                                           is_normalizing = space ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml          = ixml_or_isxml
                                       iv_xml_string             = `<A>  <B>  </B>  </A>`
                                       iv_normalizing            = abap_false
                                       iv_preserve_space_element = abap_true ).
      " WHEN
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A><B>  </B></A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD space_normalizing_on.
* Method GET_IXML_FROM_ZIP_ARCHIVE of class ZCL_EXCEL_READER_2007:
*    lo_ixml           = cl_ixml=>create( ).
*    lo_streamfactory  = lo_ixml->create_stream_factory( ).
*    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
*    r_ixml            = lo_ixml->create_document( ).
*    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
*                                                istream        = lo_istream
*                                                document       = r_ixml ).
*    lo_parser->set_normalizing( is_normalizing ).
*    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
*    lo_parser->parse( ).
* ALL CALLS ARE DONE WITH is_normalizing = 'X' except this one in method LOAD_SHARED_STRINGS of ZCL_EXCEL_READER_2007:
*    lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
*                                                           is_normalizing = space ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml          = ixml_or_isxml
                                       iv_xml_string             = `<A>  <B>  </B>  </A>`
                                       iv_normalizing            = abap_true
                                       iv_preserve_space_element = abap_true ).
      " WHEN
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A><B>  </B></A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD space_normalizing_on_strip_on.
* Method PARSE_STRING of class ZCL_EXCEL_THEME_FMT_SCHEME:
*    li_ixml = cl_ixml=>create( ).
*    li_document = li_ixml->create_document( ).
*    li_factory = li_ixml->create_stream_factory( ).
*    li_istream = li_factory->create_istream_string( iv_string ).
*    li_parser = li_ixml->create_parser(
*      stream_factory = li_factory
*      istream        = li_istream
*      document       = li_document ).
*    li_parser->add_strip_space_element( ).
*    li_parser->parse( ).
*    li_istream->close( ).
*    ri_node = li_document->get_first_child( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml          = ixml_or_isxml
                                       iv_xml_string             = `<A>  <B>  </B>  </A>`
                                       iv_normalizing            = abap_true
                                       iv_preserve_space_element = abap_false ).
      " WHEN
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A><B/></A>` ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_parser IMPLEMENTATION.
  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.

  METHOD set_validating.
* Method GET_IXML_FROM_ZIP_ARCHIVE of class ZCL_EXCEL_READER_2007:
*    lo_ixml           = cl_ixml=>create( ).
*    lo_streamfactory  = lo_ixml->create_stream_factory( ).
*    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
*    r_ixml            = lo_ixml->create_document( ).
*    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
*                                                istream        = lo_istream
*                                                document       = r_ixml ).
*    lo_parser->set_normalizing( is_normalizing ).
*    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
*    lo_parser->parse( ).
* Method PARSE_STRING of class ZCL_EXCEL_THEME_FMT_SCHEME:
*    li_ixml = cl_ixml=>create( ).
*    li_document = li_ixml->create_document( ).
*    li_factory = li_ixml->create_stream_factory( ).
*    li_istream = li_factory->create_istream_string( iv_string ).
*    li_parser = li_ixml->create_parser(
*      stream_factory = li_factory
*      istream        = li_istream
*      document       = li_document ).
*    li_parser->add_strip_space_element( ).
*    li_parser->parse( ).
*    li_istream->close( ).
*    ri_node = li_document->get_first_child( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A/>`
                                       iv_validating    = zif_excel_xml_parser=>co_no_validation ).
      element = document->get_root_element( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD several_children.
    DATA lo_element TYPE REF TO zif_excel_xml_element.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A>T<B>U</B><C/></A>` ).
      cl_abap_unit_assert=>assert_equals( act = rc
                                          exp = zif_excel_xml_constants=>ixml_mr-dom_ok ).
      element = document->get_root_element( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
      text ?= element->get_first_child( ).
      cl_abap_unit_assert=>assert_equals( act = text->get_value( )
                                          exp = 'T' ).
      lo_element ?= text->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_element->get_name( )
                                          exp = 'B' ).
      text ?= lo_element->get_first_child( ).
      cl_abap_unit_assert=>assert_equals( act = text->get_value( )
                                          exp = 'U' ).
      lo_element ?= lo_element->get_next( ).
      cl_abap_unit_assert=>assert_equals( act = lo_element->get_name( )
                                          exp = 'C' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD space_normalizing_off.
* FOR NOW, I CAN'T MAKE WORK NORMALIZING OFF and PRESERVE SPACE ELEMENT.
** Method GET_IXML_FROM_ZIP_ARCHIVE of class ZCL_EXCEL_READER_2007:
**    lo_ixml           = cl_ixml=>create( ).
**    lo_streamfactory  = lo_ixml->create_stream_factory( ).
**    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
**    r_ixml            = lo_ixml->create_document( ).
**    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
**                                                istream        = lo_istream
**                                                document       = r_ixml ).
**    lo_parser->set_normalizing( is_normalizing ).
**    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
**    lo_parser->parse( ).
** ALL CALLS ARE DONE WITH is_normalizing = 'X' except this one in method LOAD_SHARED_STRINGS of ZCL_EXCEL_READER_2007:
**    lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
**                                                           is_normalizing = space ).
*    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
*      " GIVEN
*      document = lth_isxmlixml=>parse( io_ixml_or_isxml          = ixml_or_isxml
*                                       iv_xml_string             = `<A>  <B>  </B>  </A>`
*                                       iv_normalizing            = abap_false
*                                       iv_preserve_space_element = abap_true ).
*      " WHEN
*      element = document->get_root_element( ).
*      " THEN
*      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
*                                          exp = 'A' ).
*      node = element->get_first_child( ).
*      cl_abap_unit_assert=>assert_equals( act = node->get_value( )
*                                          exp = `  ` ).
*      element ?= node->get_next( ).
*      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
*                                          exp = 'B' ).
*    ENDLOOP.
  ENDMETHOD.

  METHOD space_normalizing_on.
* Method GET_IXML_FROM_ZIP_ARCHIVE of class ZCL_EXCEL_READER_2007:
*    lo_ixml           = cl_ixml=>create( ).
*    lo_streamfactory  = lo_ixml->create_stream_factory( ).
*    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
*    r_ixml            = lo_ixml->create_document( ).
*    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
*                                                istream        = lo_istream
*                                                document       = r_ixml ).
*    lo_parser->set_normalizing( is_normalizing ).
*    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
*    lo_parser->parse( ).
* ALL CALLS ARE DONE WITH is_normalizing = 'X' except this one in method LOAD_SHARED_STRINGS of ZCL_EXCEL_READER_2007:
*    lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
*                                                           is_normalizing = space ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml          = ixml_or_isxml
                                       iv_xml_string             = `<A>  <B>  </B>  </A>`
                                       iv_normalizing            = abap_true
                                       iv_preserve_space_element = abap_true ).
      " WHEN
      element = document->get_root_element( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
      " WHEN
      element ?= element->get_first_child( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'B' ).
      " WHEN
      text ?= element->get_first_child( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = text->get_value( )
                                          exp = `  ` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD space_normalizing_on_strip_on.
* Method PARSE_STRING of class ZCL_EXCEL_THEME_FMT_SCHEME:
*    li_ixml = cl_ixml=>create( ).
*    li_document = li_ixml->create_document( ).
*    li_factory = li_ixml->create_stream_factory( ).
*    li_istream = li_factory->create_istream_string( iv_string ).
*    li_parser = li_ixml->create_parser(
*      stream_factory = li_factory
*      istream        = li_istream
*      document       = li_document ).
*    li_parser->add_strip_space_element( ).
*    li_parser->parse( ).
*    li_istream->close( ).
*    ri_node = li_document->get_first_child( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      " GIVEN
      document = lth_isxmlixml=>parse( io_ixml_or_isxml          = ixml_or_isxml
                                       iv_xml_string             = `<A>  <B>  </B>  </A>`
                                       iv_normalizing            = abap_true
                                       iv_preserve_space_element = abap_false ).
      " WHEN
      element = document->get_root_element( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
      " WHEN
      element ?= element->get_first_child( ).
      " THEN
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'B' ).
      " WHEN
      text ?= element->get_first_child( ).
      " THEN
      cl_abap_unit_assert=>assert_not_bound( act = text ).
    ENDLOOP.
  ENDMETHOD.

  METHOD text_node.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A>B</A>` ).
      cl_abap_unit_assert=>assert_equals( act = rc
                                          exp = zif_excel_xml_constants=>ixml_mr-dom_ok ).
      element = document->get_root_element( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
      text ?= element->get_first_child( ).
      cl_abap_unit_assert=>assert_equals( act = text->get_value( )
                                          exp = 'B' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD two_ixml_encodings.
    DATA lo_encoding TYPE REF TO zif_excel_xml_encoding.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      encoding = ixml_or_isxml->create_encoding( byte_order    = zif_excel_xml_encoding=>co_platform_endian
                                                 character_set = 'UTF-8' ).
      lo_encoding = ixml_or_isxml->create_encoding( byte_order    = zif_excel_xml_encoding=>co_platform_endian
                                                    character_set = 'UTF-8' ).
      cl_abap_unit_assert=>assert_true( boolc( encoding <> lo_encoding ) ).
    ENDLOOP.
  ENDMETHOD.

  METHOD two_ixml_instances.
    DATA lo_ixml TYPE REF TO zif_excel_xml.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      CASE ixml_or_isxml.
        WHEN ixml.
          lo_ixml = zcl_excel_ixml=>create( ).
        WHEN isxml.
          lo_ixml = zcl_excel_xml=>create( ).
      ENDCASE.
      cl_abap_unit_assert=>assert_equals( act = lo_ixml
                                          exp = ixml_or_isxml ).
    ENDLOOP.
  ENDMETHOD.

  METHOD two_ixml_stream_factories.
    DATA lo_streamfactory TYPE REF TO zif_excel_xml_stream_factory.

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      stream_factory = ixml_or_isxml->create_stream_factory( ).
      lo_streamfactory = ixml_or_isxml->create_stream_factory( ).
      cl_abap_unit_assert=>assert_true( boolc( lo_streamfactory <> stream_factory ) ).
    ENDLOOP.
  ENDMETHOD.
*  METHOD two_parsers.
*    DATA lo_istream_2 TYPE REF TO zif_excel_ixml_istream.
*    DATA lo_parser_2  TYPE REF TO zif_excel_ixml_parser.
*
*    LOOP AT ixml_and_isxml INTO ixml_or_isxml
*         WHERE table_line = ixml
**         WHERE table_line = isxml
*         .
*      document = ixml_or_isxml->create_document( ).
*      stream_factory = ixml_or_isxml->create_stream_factory( ).
*
*      xstring = cl_abap_codepage=>convert_to( |<B/>| ).
*      istream = stream_factory->create_istream_xstring( xstring ).
*      parser = ixml_or_isxml->create_parser( stream_factory = stream_factory
*                                             istream        = istream
*                                             document       = document ).
*      parser->set_normalizing( abap_true ).
*      parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
*      rc = parser->parse( ).
*
*      string = '<A/>'.
*      lo_istream_2 = stream_factory->create_istream_string( string ).
*      parser = ixml_or_isxml->create_parser( stream_factory = stream_factory
*                                             istream        = lo_istream_2
*                                             document       = document ).
*      rc = parser->parse( ).
*
*      element = document->get_root_element( ).
*      string = element->get_name( ).
*
*      " The second parsing is ignored
*      cl_abap_unit_assert=>assert_equals( act = string
*                                          exp = 'B' ).
*    ENDLOOP.
*  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_render IMPLEMENTATION.
  METHOD most_simple_valid_xml.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = ixml_or_isxml->create_document( ).
      element = document->create_simple_element( name   = 'ROOT'
                                                 parent = document ).
      stream_factory = ixml_or_isxml->create_stream_factory( ).
      GET REFERENCE OF xstring INTO ref_xstring.
      ostream = stream_factory->create_ostream_xstring( ref_xstring ).
      renderer = ixml_or_isxml->create_renderer( ostream  = ostream
                                                 document = document ).
      CLEAR xstring.
      renderer->render( ).
      string = cl_abap_codepage=>convert_from( xstring ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = '<?xml version="1.0"?><ROOT/>' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD namespace.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse(
          io_ixml_or_isxml = ixml_or_isxml
          iv_xml_string    = `<nsprefix:A xmlns="dnsuri" xmlns:nsprefix="nsuri" nsprefix:attr="1" attr="2"><B attr="3"/></nsprefix:A>` ).
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      cl_abap_unit_assert=>assert_equals(
          act = string
          exp = `<nsprefix:A nsprefix:attr="1" attr="2" xmlns="dnsuri" xmlns:nsprefix="nsuri"><B attr="3"/></nsprefix:A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD namespace_2.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A xmlns:a="nsuri"><a:B/></A>` ).
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A xmlns:a="nsuri"><a:B/></A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD namespace_3.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = `<A xmlns:a="nsuri"><B a:a=""/></A>` ).
      string = lth_isxmlixml=>render( ixml_or_isxml ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = `<A xmlns:a="nsuri"><B a:a=""/></A>` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_stream IMPLEMENTATION.
  METHOD close.
* Method PARSE_STRING of class ZCL_EXCEL_THEME_FMT_SCHEME:
*    li_ixml = cl_ixml=>create( ).
*    li_document = li_ixml->create_document( ).
*    li_factory = li_ixml->create_stream_factory( ).
*    li_istream = li_factory->create_istream_string( iv_string ).
*    li_parser = li_ixml->create_parser(
*      stream_factory = li_factory
*      istream        = li_istream
*      document       = li_document ).
*    li_parser->add_strip_space_element( ).
*    li_parser->parse( ).
*    li_istream->close( ).
*    ri_node = li_document->get_first_child( ).
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      document = lth_isxmlixml=>parse( io_ixml_or_isxml = ixml_or_isxml
                                       iv_xml_string    = '<A/>' ).
      lth_isxmlixml=>istream->close( ).
      element = document->get_root_element( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = `A` ).
    ENDLOOP.
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlixml_stream_factory IMPLEMENTATION.
  METHOD create_istream_string.
* Method PARSE_STRING of class ZCL_EXCEL_THEME_FMT_SCHEME:
*    li_ixml = cl_ixml=>create( ).
*    li_document = li_ixml->create_document( ).
*    li_factory = li_ixml->create_stream_factory( ).
*    li_istream = li_factory->create_istream_string( iv_string ).
*    li_parser = li_ixml->create_parser(
*      stream_factory = li_factory
*      istream        = li_istream
*      document       = li_document ).
*    li_parser->add_strip_space_element( ).
*    li_parser->parse( ).
*    li_istream->close( ).
*    ri_node = li_document->get_first_child( ).

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      prepare( ixml_or_isxml ).
      istream = stream_factory->create_istream_string( `<A/>` ).
      document = parse( io_ixml_or_isxml = ixml_or_isxml
                        io_istream       = istream ).
      element = document->get_root_element( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create_istream_xstring.
* Method GET_IXML_FROM_ZIP_ARCHIVE of class ZCL_EXCEL_READER_2007:
*    lo_ixml           = cl_ixml=>create( ).
*    lo_streamfactory  = lo_ixml->create_stream_factory( ).
*    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
*    r_ixml            = lo_ixml->create_document( ).
*    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
*                                                istream        = lo_istream
*                                                document       = r_ixml ).
*    lo_parser->set_normalizing( is_normalizing ).
*    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
*    lo_parser->parse( ).

    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      prepare( ixml_or_isxml ).
      istream = stream_factory->create_istream_xstring( cl_abap_codepage=>convert_to( `<A/>` ) ).
      document = parse( io_ixml_or_isxml = ixml_or_isxml
                        io_istream       = istream ).
      element = document->get_root_element( ).
      cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                          exp = 'A' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create_ostream_cstring.
* Method RENDER_XML_DOCUMENT of class ZCL_EXCEL_WRITER_2007:
*    lo_streamfactory = me->ixml->create_stream_factory( ).
*    lo_ostream = lo_streamfactory->create_ostream_cstring( string = lv_string ).
*    lo_renderer = me->ixml->create_renderer( ostream  = lo_ostream document = io_document ).
*    lo_renderer->render( ).
    GET REFERENCE OF string INTO ref_string.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      prepare( ixml_or_isxml ).
      document->create_simple_element( name   = 'A'
                                       parent = document ).
      ostream = stream_factory->create_ostream_cstring( ref_string ).
      CLEAR string.
      render_ostream( io_ixml_or_isxml = ixml_or_isxml
                      io_ostream       = ostream ).
      cl_abap_unit_assert=>assert_equals( act = string
                                          exp = |{ lcl_bom_utf16_as_character=>system_value }<A/>| ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create_ostream_xstring.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM:
*    CLEAR ep_content.
*    lo_ixml = cl_ixml=>create( ).
*    lo_streamfactory = lo_ixml->create_stream_factory( ).
*    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
*    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
*    lo_renderer->render( ).
    GET REFERENCE OF xstring INTO ref_xstring.
    LOOP AT ixml_and_isxml INTO ixml_or_isxml.
      prepare( ixml_or_isxml ).
      document->create_simple_element( name   = 'A'
                                       parent = document ).
      ostream = stream_factory->create_ostream_xstring( ref_xstring ).
      CLEAR xstring.
      render_ostream( io_ixml_or_isxml = ixml_or_isxml
                      io_ostream       = ostream ).
      cl_abap_unit_assert=>assert_equals( act = xstring
                                          exp = cl_abap_codepage=>convert_to( `<A/>` ) ).
    ENDLOOP.
  ENDMETHOD.

  METHOD parse.
    parser = io_ixml_or_isxml->create_parser( stream_factory = stream_factory
                                              istream        = io_istream
                                              document       = document ).
    parser->parse( ).
    ro_result = document.
  ENDMETHOD.

  METHOD prepare.
    document = io_ixml_or_isxml->create_document( ).
    stream_factory = io_ixml_or_isxml->create_stream_factory( ).
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.

  METHOD render_ostream.
    DATA lo_renderer TYPE REF TO zif_excel_xml_renderer.

    lo_renderer = ixml_or_isxml->create_renderer( ostream  = io_ostream
                                                  document = document ).
    document->set_declaration( abap_false ).
    lo_renderer->render( ).
*    " remove the UTF-16 BOM (i.e. remove the first character)
*    SHIFT rv_result LEFT BY 1 PLACES.
*
*    " Normalize XML according to SXML limitations in order to compare IXML
*    " and SXML results by simple string comparison.
*    rv_result = lcl_rewrite_xml_via_sxml=>execute( rv_result ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_isxmlonly_parser IMPLEMENTATION.
  METHOD namespace.
    DATA lo_isxml_attribute TYPE REF TO lcl_isxml_attribute.

    " GIVEN
    document = lth_isxmlixml=>parse( io_ixml_or_isxml = isxml
                                     iv_xml_string    = `<A xmlns:a="nsuri"><B a:a=""/></A>` ).
    " WHEN
    lo_isxml_attribute ?= document->get_root_element( )->get_first_child( )->get_attributes( )->create_iterator( )->get_next( ). "->get_next( ).
    " THEN
    cl_abap_unit_assert=>assert_bound( act = lo_isxml_attribute ).
    cl_abap_unit_assert=>assert_equals( act  = lo_isxml_attribute->prefix
                                        exp  = 'a'
                                        msg  = 'prefix'
                                        quit = if_aunit_constants=>no ).
    cl_abap_unit_assert=>assert_equals( act  = lo_isxml_attribute->name
                                        exp  = 'a'
                                        msg  = 'name'
                                        quit = if_aunit_constants=>no ).
*    cl_abap_unit_assert=>assert_equals( act  = lo_isxml_attribute->nsuri
*                                        exp  = 'nsuri'
*                                        msg  = 'namespace'
*                                        quit = if_aunit_constants=>no ).
  ENDMETHOD.

  METHOD setup.
    ixml_and_isxml = get_ixml_and_isxml( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_rewrite_xml_via_sxml IMPLEMENTATION.
  METHOD default_namespace.
    string = rewrite_xml_via_sxml(
                 `<A xmlns="dnsuri" xmlns:nsprefix="nsuri" nsprefix:attr="1" attr="2"><B attr="3"/></A>` ).
    cl_abap_unit_assert=>assert_equals(
        act = string
        exp = `<A nsprefix:attr="1" attr="2" xmlns="dnsuri" xmlns:nsprefix="nsuri"><B attr="3"/></A>` ).

    set_current_parsed_element( iv_index = 1 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'A'
                                                                    iv_namespace = 'dnsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_prefix = 'nsprefix'
                                                                      iv_nsuri  = 'nsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 2 )
                                        exp = get_expected_nsbinding( iv_nsuri = 'dnsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_attribute( iv_index = 1 )
                                        exp = get_expected_attribute( iv_name      = 'attr'
                                                                      iv_namespace = 'nsuri'
                                                                      iv_prefix    = 'nsprefix' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_attribute( iv_index = 2 )
                                        exp = get_expected_attribute( iv_name = 'attr' ) ).
    set_current_parsed_element( iv_index = 2 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'B'
                                                                    iv_namespace = 'dnsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_prefix = 'nsprefix'
                                                                      iv_nsuri  = 'nsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 2 )
                                        exp = get_expected_nsbinding( iv_nsuri = 'dnsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_attribute( iv_index = 1 )
                                        exp = get_expected_attribute( iv_name = 'attr' ) ).
  ENDMETHOD.

  METHOD default_namespace_removed.
    string = rewrite_xml_via_sxml( `<A><B xmlns="dnsuri"><C xmlns=""><D/></C></B></A>` ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = `<A><B xmlns="dnsuri"><C xmlns=""><D/></C></B></A>` ).

    set_current_parsed_element( iv_index = 1 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name = 'A' ) ).
    set_current_parsed_element( iv_index = 2 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'B'
                                                                    iv_namespace = 'dnsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_nsuri = 'dnsuri' ) ).
    set_current_parsed_element( iv_index = 3 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name = 'C' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_nsuri = '' ) ).
    set_current_parsed_element( iv_index = 4 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name = 'D' ) ).
  ENDMETHOD.

  METHOD get_expected_attribute.
    rs_result-name      = iv_name.
    rs_result-namespace = iv_namespace.
    rs_result-prefix    = iv_prefix.
  ENDMETHOD.

  METHOD get_expected_element.
    rs_result-name      = iv_name.
    rs_result-namespace = iv_namespace.
    rs_result-prefix    = iv_prefix.
  ENDMETHOD.

  METHOD get_expected_nsbinding.
    rs_result-prefix = iv_prefix.
    rs_result-nsuri  = iv_nsuri.
  ENDMETHOD.

  METHOD get_parsed_element.
    DATA lr_complete_parsed_element TYPE REF TO lcl_rewrite_xml_via_sxml=>ts_complete_element.

    READ TABLE lcl_rewrite_xml_via_sxml=>complete_parsed_elements REFERENCE INTO lr_complete_parsed_element INDEX parsed_element_index.
    IF sy-subrc = 0.
      rs_result = lr_complete_parsed_element->element.
    ENDIF.
  ENDMETHOD.

  METHOD get_parsed_element_attribute.
    DATA lr_complete_parsed_element TYPE REF TO lcl_rewrite_xml_via_sxml=>ts_complete_element.

    READ TABLE lcl_rewrite_xml_via_sxml=>complete_parsed_elements REFERENCE INTO lr_complete_parsed_element INDEX parsed_element_index.
    IF sy-subrc = 0.
      READ TABLE lr_complete_parsed_element->attributes INTO rs_result INDEX iv_index.
    ENDIF.
  ENDMETHOD.

  METHOD get_parsed_element_nsbinding.
    DATA lr_complete_parsed_element TYPE REF TO lcl_rewrite_xml_via_sxml=>ts_complete_element.

    READ TABLE lcl_rewrite_xml_via_sxml=>complete_parsed_elements REFERENCE INTO lr_complete_parsed_element INDEX parsed_element_index.
    IF sy-subrc = 0.
      READ TABLE lr_complete_parsed_element->nsbindings INTO rs_result INDEX iv_index.
    ENDIF.
  ENDMETHOD.

  METHOD namespace.
    string = rewrite_xml_via_sxml(
        `<nsprefix:A xmlns="dnsuri" xmlns:nsprefix="nsuri" nsprefix:attr="1" attr="2"><B attr="3"/></nsprefix:A>` ).
    cl_abap_unit_assert=>assert_equals(
        act = string
        exp = `<nsprefix:A nsprefix:attr="1" attr="2" xmlns="dnsuri" xmlns:nsprefix="nsuri"><B attr="3"/></nsprefix:A>` ).

    cl_abap_unit_assert=>assert_equals( act = rewrite_xml_via_sxml( `<A xmlns=""/>` )
                                        exp = `<A xmlns=""/>` ).

    cl_abap_unit_assert=>assert_equals( act = rewrite_xml_via_sxml( `<A xmlns=""><B/></A>` )
                                        exp = `<A xmlns=""><B/></A>` ).

    cl_abap_unit_assert=>assert_equals( act = rewrite_xml_via_sxml( `<A><B xmlns=""/></A>` )
                                        exp = `<A><B xmlns=""/></A>` ).
  ENDMETHOD.

  METHOD namespace_2.
    string = rewrite_xml_via_sxml( `<A xmlns="dnsuri" xmlns:nsprefix="nsuri"><B/><nsprefix:C/></A>` ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = `<A xmlns="dnsuri" xmlns:nsprefix="nsuri"><B/><nsprefix:C/></A>` ).

    set_current_parsed_element( iv_index = 1 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'A'
                                                                    iv_namespace = 'dnsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_prefix = 'nsprefix'
                                                                      iv_nsuri  = 'nsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 2 )
                                        exp = get_expected_nsbinding( iv_nsuri = 'dnsuri' ) ).
    set_current_parsed_element( iv_index = 2 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'B'
                                                                    iv_namespace = 'dnsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_prefix = 'nsprefix'
                                                                      iv_nsuri  = 'nsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 2 )
                                        exp = get_expected_nsbinding( iv_nsuri = 'dnsuri' ) ).
    set_current_parsed_element( iv_index = 3 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'C'
                                                                    iv_namespace = 'nsuri'
                                                                    iv_prefix    = 'nsprefix' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_prefix = 'nsprefix'
                                                                      iv_nsuri  = 'nsuri' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 2 )
                                        exp = get_expected_nsbinding( iv_nsuri = 'dnsuri' ) ).
  ENDMETHOD.

  METHOD namespace_3.
    string = rewrite_xml_via_sxml( `<a:A xmlns:a="nsuri_a"><b:B xmlns:b="nsuri_b"/></a:A>` ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = `<a:A xmlns:a="nsuri_a"><b:B xmlns:b="nsuri_b"/></a:A>` ).

    set_current_parsed_element( iv_index = 1 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'A'
                                                                    iv_namespace = 'nsuri_a'
                                                                    iv_prefix    = 'a' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_prefix = 'a'
                                                                      iv_nsuri  = 'nsuri_a' ) ).
    set_current_parsed_element( iv_index = 2 ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element( )
                                        exp = get_expected_element( iv_name      = 'B'
                                                                    iv_namespace = 'nsuri_b'
                                                                    iv_prefix    = 'b' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 1 )
                                        exp = get_expected_nsbinding( iv_prefix = 'b'
                                                                      iv_nsuri  = 'nsuri_b' ) ).
    cl_abap_unit_assert=>assert_equals( act = get_parsed_element_nsbinding( iv_index = 2 )
                                        exp = get_expected_nsbinding( iv_prefix = 'a'
                                                                      iv_nsuri  = 'nsuri_a' ) ).
  ENDMETHOD.

  METHOD rewrite_xml_via_sxml.
    rv_string = lcl_rewrite_xml_via_sxml=>execute( iv_xml_string = iv_xml_string
                                                   iv_trace      = abap_true ).
  ENDMETHOD.

  METHOD set_current_parsed_element.
    parsed_element_index = iv_index.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_sxml_reader IMPLEMENTATION.
  METHOD bom.
    xstring = cl_abap_codepage=>convert_to(
        '<?xml version="1.0"  standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="24" count="24">' &&
        '<si><t>Olá Mundo</t></si><si><t>Привет, мир</t></si><si><t>مرحبا بالعالم</t></si><si><t>ہیلو دنیا</t></si><si><t>नमस्ते दुनिया</t></si><si><t>ওহে বিশ্ব</t></si><si><t>你好，世界</t></si><si><t>👋🌎, 👋🌍, 👋🌏</t></si></sst>' ).
    CONCATENATE cl_abap_char_utilities=>byte_order_mark_utf8
                xstring
                INTO xstring
                IN BYTE MODE.
    " xstring = 'EFBBBF3C3F786D6C2076657273696F6E3D22312E302220207374616E64616C6F6E653D22796573223F3E3C73737420636F756E743D2232342220756E69717565436F756E743D2232342220786D6C6E733D22687474703A2F2F736368656D61732E6F70656E786D6C666F726D6174732E6F72672F73707
    "2' &&
    " '65616473686565746D6C2F323030362F6D61696E223E3C73693E3C743E28417261626963293C2F743E3C2F73693E3C73693E3C743E2842656E67616C69293C2F743E3C2F73693E3C73693E3C743E284368696E657365293C2F743E3C2F73693E3C73693E3C743E28456D6F6A6920776176696E672068616E64202B2
    "0' &&
    " '33207061727473206F662074686520776F726C64293C2F743E3C2F73693E3C73693E3C743E284672656E6368293C2F743E3C2F73693E3C73693E3C743E2848696E6469293C2F743E3C2F73693E3C73693E3C743E28506F7274756775657365293C2F743E3C2F73693E3C73693E3C743E285275737369616E293C2F7
    "4' &&
    " '3E3C2F73693E3C73693E3C743E285370616E697368293C2F743E3C2F73693E3C73693E3C743E2855726475293C2F743E3C2F73693E3C73693E3C743E426F6E6A6F7572206C65206D6F6E64653C2F743E3C2F73693E3C73693E3C743E436C69636B206865726520746F207669736974206162617032786C737820686
    "F' &&
    " '6D65706167653C2F743E3C2F73693E3C73693E3C743E48656C6C6F20776F726C643C2F743E3C2F73693E3C73693E3C743E486F6C61204D756E646F3C2F743E3C2F73693E3C73693E3C743E4F6CC3A1204D756E646F3C2F743E3C2F73693E3C73693E3C743ED09FD180D0B8D0B2D0B5D1822C20D0BCD0B8D1803C2F7
    "4' &&
    " '3E3C2F73693E3C73693E3C743ED985D8B1D8ADD8A8D8A720D8A8D8A7D984D8B9D8A7D984D9853C2F743E3C2F73693E3C73693E3C743EDB81DB8CD984D98820D8AFD986DB8CD8A73C2F743E3C2F73693E3C73693E3C743EE0A4A8E0A4AEE0A4B8E0A58DE0A4A4E0A58720E0A4A6E0A581E0A4A8E0A4BFE0A4AFE0A4B
    "E' &&
    " '3C2F743E3C2F73693E3C73693E3C743EE0A693E0A6B9E0A78720E0A6ACE0A6BFE0A6B6E0A78DE0A6AC3C2F743E3C2F73693E3C73693E3C743EE4BDA0E5A5BDEFBC8CE4B896E7958C3C2F743E3C2F73693E3C73693E3C743EF09F918BF09F8C8E2C20F09F918BF09F8C8D2C20F09F918BF09F8C8F3C2F743E3C2F736
    "9' &&
    " '3E3C2F7373743E'.
    reader = cl_sxml_string_reader=>create( xstring ).
    reader->set_option( if_sxml_reader=>co_opt_keep_whitespace ).
    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_not_bound( act = node ).
*    cl_abap_unit_assert=>assert_equals( act = node->type
*                                        exp = if_sxml_node=>co_nt_element_open ).
*    cl_abap_unit_assert=>assert_equals( act = reader->name
*                                        exp = 'sst' ).
  ENDMETHOD.

  METHOD empty_object_oriented_parsing.
    xstring = cl_abap_codepage=>convert_to( '<ROOTNODE/>' ).
    reader = cl_sxml_string_reader=>create( xstring ).
    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_bound( node ).
    cl_abap_unit_assert=>assert_equals( act = node->type
                                        exp = node->co_nt_element_open ).
    open_element ?= node.
    cl_abap_unit_assert=>assert_equals( act = open_element->qname-name
                                        exp = 'ROOTNODE' ).
    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_bound( node ).
    cl_abap_unit_assert=>assert_equals( act = node->type
                                        exp = node->co_nt_element_close ).
    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_not_bound( node ).
  ENDMETHOD.

  METHOD empty_token_based_parsing.
    xstring = cl_abap_codepage=>convert_to( '<ROOTNODE/>' ).
    reader = cl_sxml_string_reader=>create( xstring ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_initial ).
    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_element_open ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ROOTNODE' ).
    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_element_close ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ROOTNODE' ).
    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_final ).
  ENDMETHOD.

  METHOD empty_xml.
    DATA parse_error TYPE REF TO cx_sxml_parse_error.

    CLEAR xstring.
    reader = cl_sxml_string_reader=>create( xstring ).
    TRY.
        node = reader->read_next_node( ).
        cl_abap_unit_assert=>fail( msg = 'should have failed' ).
      CATCH cx_root INTO error.
        error_rtti = cl_abap_typedescr=>describe_by_object_ref( error ).
        cl_abap_unit_assert=>assert_equals( act = error_rtti->get_relative_name( )
                                            exp = 'CX_SXML_PARSE_ERROR' ).
        parse_error ?= error.
        cl_abap_unit_assert=>assert_equals( act = parse_error->textid
                                            exp = parse_error->kernel_parser ).
        cl_abap_unit_assert=>assert_equals( act = parse_error->error_text
                                            exp = 'BOM / charset detection failed' ).
    ENDTRY.
  ENDMETHOD.

  METHOD invalid_xml.
    xstring = cl_abap_codepage=>convert_to( '<' ).
    reader = cl_sxml_string_reader=>create( xstring ).
    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_not_bound( node ).
  ENDMETHOD.

  METHOD invalid_xml_eof_reached.
    xstring = cl_abap_codepage=>convert_to( '<A>' ).
    reader = cl_sxml_string_reader=>create( xstring ).
    node = reader->read_next_node( ).
    TRY.
        node = reader->read_next_node( ).
        cl_abap_unit_assert=>assert_not_bound( node ).
        cl_abap_unit_assert=>fail( msg = 'should have failed' ).
      CATCH cx_root INTO error.
        error_rtti = cl_abap_typedescr=>describe_by_object_ref( error ).
        cl_abap_unit_assert=>assert_equals( act = error_rtti->get_relative_name( )
                                            exp = 'CX_SXML_PARSE_ERROR' ).
        parse_error ?= error.
        cl_abap_unit_assert=>assert_equals( act = parse_error->textid
                                            exp = parse_error->kernel_parser ).
        cl_abap_unit_assert=>assert_equals( act = parse_error->error_text
                                            exp = '<EOF> reached' ).
    ENDTRY.
  ENDMETHOD.

  METHOD invalid_xml_not_wellformed.
    xstring = cl_abap_codepage=>convert_to( '<A></B>' ).
    reader = cl_sxml_string_reader=>create( xstring ).
    node = reader->read_next_node( ).
    TRY.
        node = reader->read_next_node( ).
        cl_abap_unit_assert=>assert_not_bound( node ).
        cl_abap_unit_assert=>fail( msg = 'should have failed' ).
      CATCH cx_root INTO error.
        error_rtti = cl_abap_typedescr=>describe_by_object_ref( error ).
        cl_abap_unit_assert=>assert_equals( act = error_rtti->get_relative_name( )
                                            exp = 'CX_SXML_PARSE_ERROR' ).
        parse_error ?= error.
        cl_abap_unit_assert=>assert_equals( act = parse_error->textid
                                            exp = parse_error->kernel_parser ).
        cl_abap_unit_assert=>assert_equals( act = parse_error->error_text
                                            exp = 'document not wellformed' ).
    ENDTRY.
  ENDMETHOD.

  METHOD keep_whitespace.
    xstring = cl_abap_codepage=>convert_to(
        '<?xml version="1.0"  standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="24" count="24">' &&
        '<si><t>Olá Mundo</t></si><si><t>Привет, мир</t></si><si><t>مرحبا بالعالم</t></si><si><t>ہیلو دنیا</t></si><si><t>नमस्ते दुनिया</t></si><si><t>ওহে বিশ্ব</t></si><si><t>你好，世界</t></si><si><t>👋🌎, 👋🌍, 👋🌏</t></si></sst>' ).
    CONCATENATE cl_abap_char_utilities=>byte_order_mark_utf8
                xstring
                INTO xstring
                IN BYTE MODE.
    reader = cl_sxml_string_reader=>create( xstring ).
    reader->set_option( if_sxml_reader=>co_opt_keep_whitespace ).
    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_not_bound( act = node ).
  ENDMETHOD.

  METHOD normalization.
    xstring = cl_abap_codepage=>convert_to( |<A><B>  1 \n 2  </B></A>| ).
    reader = cl_sxml_string_reader=>create( xstring ).
    reader->set_option( if_sxml_reader=>co_opt_normalizing ).
    open_element ?= reader->read_next_node( ).
    cl_abap_unit_assert=>assert_equals( act = open_element->qname-name
                                        exp = `A` ).
    open_element ?= reader->read_next_node( ).
    cl_abap_unit_assert=>assert_equals( act = open_element->qname-name
                                        exp = `B` ).
    value_node ?= reader->read_next_node( ).
    cl_abap_unit_assert=>assert_equals( act = value_node->get_value( )
                                        exp = |  1 \n 2  | ).
  ENDMETHOD.

  METHOD object_oriented_parsing.
    xstring = cl_abap_codepage=>convert_to( '<ROOTNODE ATTR="Efe=">Efe=</ROOTNODE>' ).
    reader = cl_sxml_string_reader=>create( xstring ).

    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_bound( node ).
    cl_abap_unit_assert=>assert_equals( act = node->type
                                        exp = node->co_nt_element_open ).
    open_element ?= node.
    cl_abap_unit_assert=>assert_equals( act = open_element->qname-name
                                        exp = 'ROOTNODE' ).

    node_attr = open_element->get_attribute_value( 'ATTR' ).
    cl_abap_unit_assert=>assert_bound( node_attr ).
    cl_abap_unit_assert=>assert_equals( act = node_attr->type
                                        exp = node_attr->co_vt_text ).
    cl_abap_unit_assert=>assert_equals( act = node_attr->get_value( )
                                        exp = 'Efe=' ).
    xstring = 'E0'.
    cl_abap_unit_assert=>assert_equals( act = node_attr->get_value_raw( )
                                        exp = xstring ).

    node = reader->read_current_node( ).

    node = reader->read_next_node( ).
    cl_abap_unit_assert=>assert_bound( node ).
    cl_abap_unit_assert=>assert_equals( act = node->type
                                        exp = node->co_nt_value ).
    value_node ?= node.
    cl_abap_unit_assert=>assert_equals( act = value_node->get_value( )
                                        exp = 'Efe=' ).
    xstring = 'E0'.
    cl_abap_unit_assert=>assert_equals( act = value_node->get_value_raw( )
                                        exp = xstring ).
  ENDMETHOD.

  METHOD token_based_parsing.
    xstring = cl_abap_codepage=>convert_to( '<ROOTNODE ATTR="UFE=">UFE=</ROOTNODE>' ).
    reader = cl_sxml_string_reader=>create( xstring ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_initial ).

    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_element_open ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ROOTNODE' ).

    reader->next_attribute( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_attribute ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ATTR' ).
    cl_abap_unit_assert=>assert_equals( act = reader->value_type
                                        exp = if_sxml_value=>co_vt_text ).
    cl_abap_unit_assert=>assert_equals( act = reader->value
                                        exp = 'UFE=' ).
    CLEAR xstring.
    cl_abap_unit_assert=>assert_equals( act = reader->value_raw
                                        exp = xstring ).

    reader->next_attribute( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_element_open ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ROOTNODE' ).
    reader->current_node( ).

    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_value ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ROOTNODE' ).
    cl_abap_unit_assert=>assert_equals( act = reader->value_type
                                        exp = if_sxml_value=>co_vt_text ).
    cl_abap_unit_assert=>assert_equals( act = reader->value
                                        exp = 'UFE=' ).
    CLEAR xstring.
    cl_abap_unit_assert=>assert_equals( act = reader->value_raw
                                        exp = xstring ).

    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_element_close ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ROOTNODE' ).

    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_final ).
  ENDMETHOD.

  METHOD xml_header_is_ignored.
    xstring = cl_abap_codepage=>convert_to( '<?xml version="1.0" encoding="utf-8" standalone="yes"?><ROOT/>' ).
    reader = cl_sxml_string_reader=>create( xstring ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_initial ).

    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_element_open ).
    cl_abap_unit_assert=>assert_equals( act = reader->name
                                        exp = 'ROOT' ).

    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_element_close ).

    reader->next_node( ).
    cl_abap_unit_assert=>assert_equals( act = reader->node_type
                                        exp = if_sxml_node=>co_nt_final ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_sxml_writer IMPLEMENTATION.
  METHOD attribute_namespace.
    open_element = writer->new_open_element( name = 'A' ).
    open_element->set_attribute( name   = 'attr'
                                 nsuri  = 'http://...'
                                 prefix = 'a'
                                 value  = '1' ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A a:attr="1" xmlns:a="http://..."/>' ).
  ENDMETHOD.

  METHOD attribute_xml_namespace.
    open_element = writer->new_open_element( name = 'A' ).
    open_element->set_attribute( name   = 'space'
                                 nsuri  = 'http://www.w3.org/XML/1998/namespace'
                                 prefix = 'xml'
                                 value  = 'preserve' ).
    writer->write_node( open_element ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A xml:space="preserve"/>' ).
  ENDMETHOD.

  METHOD get_output.
    DATA lo_string_writer TYPE REF TO cl_sxml_string_writer.
    DATA lv_xstring       TYPE xstring.

    lo_string_writer ?= io_writer.
    lv_xstring = lo_string_writer->get_output( ).
    rv_result = cl_abap_codepage=>convert_from( lv_xstring ).
  ENDMETHOD.

  METHOD most_simple_valid_xml.
    open_element = writer->new_open_element( 'A' ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A/>' ).
  ENDMETHOD.

  METHOD namespace.
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'http://...'
                                             prefix = 'a' ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<a:A xmlns:a="http://..."/>' ).
  ENDMETHOD.

  METHOD namespace_default.
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'http://...'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_node( open_element ).
    open_element = writer->new_open_element( name   = 'B'
                                             nsuri  = 'http://...'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A xmlns="http://..."><B/></A>' ).
  ENDMETHOD.

  METHOD namespace_default_by_attribute.
    open_element = writer->new_open_element( name = 'A' ).
    open_element->set_attribute( name  = 'xmlns'
                                 value = 'http://...' ).
    writer->write_node( open_element ).
    open_element = writer->new_open_element( name = 'B' ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A xmlns="http://..."><B/></A>' ).
  ENDMETHOD.

  METHOD namespace_inheritance.
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'http://...'
                                             prefix = 'a' ).
    writer->write_node( open_element ).
    open_element = writer->new_open_element( name   = 'B'
                                             nsuri  = 'http://...'
                                             prefix = 'a' ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<a:A xmlns:a="http://..."><a:B/></a:A>' ).
  ENDMETHOD.

  METHOD namespace_set_prefix.
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'http://...'
                                             prefix = 'a' ).
    open_element->set_prefix( 'b' ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<b:A xmlns:b="http://..."/>' ).
  ENDMETHOD.

  METHOD object_oriented_rendering.
    " WRITE_NODE and NEW_* methods
    " NB: NEW_* methods are static methods (i.e. it's valid: DATA(open_element) = cl_sxml_writer=>if_sxml_writer~new_open_element( 'ROOTNODE' ).)
    open_element = writer->new_open_element( 'ROOTNODE' ).
    open_element->set_attribute( name  = 'ATTR'
                                 value = '5' ).
    writer->write_node( open_element ).
    " WRITE_ATTRIBUTE and WRITE_ATTRIBUTE_RAW can also be used, but only after the WRITE_NODE of an element opening tag
    writer->write_attribute( name  = 'ATTR2'
                             value = 'A' ).
    " WRITE_ATTRIBUTE_RAW writes in Base64
    writer->write_attribute_raw( name  = 'ATTR3'
                                 value = '5051' ).
    value_node = writer->new_value( ).
    value_node->set_value( 'HELLO' ).
    writer->write_node( value_node ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<ROOTNODE ATTR="5" ATTR2="A" ATTR3="UFE=">HELLO</ROOTNODE>' ).
  ENDMETHOD.

  METHOD setup.
    writer = cl_sxml_string_writer=>create( ).
  ENDMETHOD.

  METHOD token_based_rendering.
    writer->open_element( 'ROOTNODE' ).
    " WRITE_ATTRIBUTE and WRITE_ATTRIBUTE_RAW can be used only after OPEN_ELEMENT
    writer->write_attribute( name  = 'ATTR'
                             value = '5' ).
    writer->write_attribute_raw( name  = 'ATTR2'
                                 value = '5051' ).
    writer->write_value( 'HELLO' ).
    writer->open_element( 'NODE' ).
    writer->write_value_raw( '5051' ).
    writer->close_element( ).
    writer->close_element( ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<ROOTNODE ATTR="5" ATTR2="UFE=">HELLO<NODE>UFE=</NODE></ROOTNODE>' ).
  ENDMETHOD.

  METHOD write_namespace_declaration.
    open_element = writer->new_open_element( name = 'A' ).
    writer->write_namespace_declaration( nsuri  = 'nsuri'
                                         prefix = 'nsprefix' ).
    writer->write_node( open_element ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A xmlns:nsprefix="nsuri"/>' ).
  ENDMETHOD.

  METHOD write_namespace_declaration_2.
    " <A xmlns:nsprefix="nsuri"/>
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'dnsuri'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_1'
                                         prefix = 'nsprefix_1' ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_2'
                                         prefix = 'nsprefix_2' ).
    writer->write_node( open_element ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals(
        act = string
        exp = '<A xmlns:nsprefix_1="nsuri_1" xmlns:nsprefix_2="nsuri_2" xmlns="dnsuri"/>' ).
  ENDMETHOD.

  METHOD write_namespace_declaration_3.
    " <A xmlns:nsprefix="nsuri"/>
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'dnsuri'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_2'
                                         prefix = 'nsprefix_2' ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_1'
                                         prefix = 'nsprefix_1' ).
    writer->write_node( open_element ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals(
        act = string
        exp = '<A xmlns:nsprefix_2="nsuri_2" xmlns:nsprefix_1="nsuri_1" xmlns="dnsuri"/>' ).
  ENDMETHOD.

  METHOD write_namespace_declaration_4.
    " <A xmlns:nsprefix="nsuri"/>
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'dnsuri'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_1'
                                         prefix = 'nsprefix_1' ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_2'
                                         prefix = 'nsprefix_2' ).
    open_element->set_attribute( name  = 'a'
                                 value = '1' ).
    writer->write_node( open_element ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals(
        act = string
        exp = '<A a="1" xmlns:nsprefix_1="nsuri_1" xmlns:nsprefix_2="nsuri_2" xmlns="dnsuri"/>' ).
  ENDMETHOD.

  METHOD write_namespace_declaration_5.
    " <A xmlns:nsprefix="nsuri"/>
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'dnsuri'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    open_element->set_attribute( name  = 'a'
                                 value = '1' ).
    writer->write_node( open_element ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_1'
                                         prefix = 'nsprefix_1' ).
    writer->write_namespace_declaration( nsuri  = 'nsuri_2'
                                         prefix = 'nsprefix_2' ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals(
        act = string
        exp = '<A a="1" xmlns="dnsuri" xmlns:nsprefix_1="nsuri_1" xmlns:nsprefix_2="nsuri_2"/>' ).
  ENDMETHOD.

  METHOD write_namespace_declaration_6.
    " <A xmlns="dnsuri" xmlns:nsprefix="nsuri"><B/><nsprefix:C/></A>
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'dnsuri'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_namespace_declaration( nsuri  = 'nsuri'
                                         prefix = 'nsprefix' ).
    writer->write_node( open_element ).

    open_element = writer->new_open_element( name   = 'B'
                                             nsuri  = 'dnsuri'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_node( open_element ).
    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    open_element = writer->new_open_element( name   = 'C'
                                             nsuri  = 'nsuri'
                                             prefix = 'nsprefix' ).
    writer->write_node( open_element ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A xmlns:nsprefix="nsuri" xmlns="dnsuri"><B/><nsprefix:C/></A>' ).
  ENDMETHOD.

  METHOD order_of_xmlns_and_attributes.
    open_element = writer->new_open_element( name   = 'A'
                                             nsuri  = 'dnsuri'
                                             prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_node( open_element ).
    writer->write_namespace_declaration( nsuri = 'dnsuri' prefix = if_sxml_named=>co_use_default_xmlns ).
    writer->write_namespace_declaration( nsuri = 'nsuri_b' prefix = 'b' ).
*    writer->write_namespace_declaration( nsuri = 'nsuri_a' prefix = 'a' ).
*    writer->write_attribute( prefix = 'xmlns' name = 'c' value = 'nsuri_c' ).
    writer->write_attribute( name = 'b' nsuri = 'dnsuri' prefix = if_sxml_named=>co_use_default_xmlns value = '1' ).
    writer->write_attribute( name = 'a' nsuri = 'nsuri_a' prefix = 'a' value = '2' ).

    close_element = writer->new_close_element( ).
    writer->write_node( close_element ).

    string = get_output( writer ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A b="1" a:a="2" xmlns="dnsuri" xmlns:b="nsuri_b" xmlns:a="nsuri_a"/>' ).
  ENDMETHOD.
ENDCLASS.


CLASS lth_isxmlixml IMPLEMENTATION.
  METHOD create_document.
    ixml = zcl_excel_xml=>create( ).
    document = ixml->create_document( ).
  ENDMETHOD.

  METHOD parse.
    DATA lv_xstring TYPE xstring.
    DATA lo_parser  TYPE REF TO zif_excel_xml_parser.

    " code inspired from the method GET_IXML_FROM_ZIP_ARCHIVE of ZCL_EXCEL_READER_2007.
    document = io_ixml_or_isxml->create_document( ).
    stream_factory = io_ixml_or_isxml->create_stream_factory( ).

    lv_xstring = cl_abap_codepage=>convert_to( iv_xml_string ).
    istream = stream_factory->create_istream_xstring( lv_xstring ).
    lo_parser = io_ixml_or_isxml->create_parser( stream_factory = stream_factory
                                                 istream        = istream
                                                 document       = document ).
    lo_parser->set_normalizing( iv_normalizing ).
    IF iv_preserve_space_element = abap_false.
      lo_parser->add_strip_space_element( ).
    ENDIF.
    IF iv_validating <> zif_excel_xml_parser=>co_no_validation.
      lo_parser->set_validating( mode = iv_validating ).
    ENDIF.
    lo_parser->parse( ).
    ro_result = document.
  ENDMETHOD.

  METHOD render.
    DATA lr_string   TYPE REF TO string.
    DATA lo_ostream  TYPE REF TO zif_excel_xml_ostream.
    DATA lo_renderer TYPE REF TO zif_excel_xml_renderer.

    stream_factory = ixml_or_isxml->create_stream_factory( ).
    GET REFERENCE OF rv_result INTO lr_string.
    lo_ostream = stream_factory->create_ostream_cstring( lr_string ).
    lo_renderer = ixml_or_isxml->create_renderer( ostream  = lo_ostream
                                                  document = document ).
    " remove the XML declaration
    document->set_declaration( abap_false ).
    " Fills RV_RESULT
    lo_renderer->render( ).
    " remove the UTF-16 BOM (i.e. remove the first character)
    SHIFT rv_result LEFT BY 1 PLACES.

    " Normalize XML according to SXML limitations in order to compare IXML
    " and SXML results by simple string comparison.
    rv_result = lcl_rewrite_xml_via_sxml=>execute( rv_result ).
  ENDMETHOD.
ENDCLASS.
