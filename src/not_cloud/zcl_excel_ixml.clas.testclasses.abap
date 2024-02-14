*"* use this source file for your ABAP unit test classes

"! SXML is less tolerant than IXML
CLASS ltc_diff_ixml_sxml_parser DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    "! Parsing of &lt;a:A/> without xmlns:a="" is possible with IXML, not with SXML
    METHODS missing_namespace_declaration FOR TESTING RAISING cx_static_check.

    DATA error               TYPE REF TO cx_root.
    DATA error_rtti          TYPE REF TO cl_abap_typedescr.
    DATA ixml                TYPE REF TO if_ixml.
    DATA ixml_document       TYPE REF TO if_ixml_document.
    DATA ixml_element        TYPE REF TO if_ixml_element.
    DATA ixml_istream        TYPE REF TO if_ixml_istream.
    DATA ixml_node           TYPE REF TO if_ixml_node.
    DATA ixml_parser         TYPE REF TO if_ixml_parser.
    DATA ixml_stream_factory TYPE REF TO if_ixml_stream_factory.
    DATA ixml_text           TYPE REF TO if_ixml_text.
    DATA sxml_open_element   TYPE REF TO if_sxml_open_element.
    DATA sxml_node           TYPE REF TO if_sxml_node.
    DATA sxml_reader         TYPE REF TO if_sxml_reader.
    DATA sxml_parse_error    TYPE REF TO cx_sxml_parse_error.
    DATA sxml_node_open      TYPE REF TO if_sxml_open_element.
    DATA sxml_node_attr      TYPE REF TO if_sxml_value.
    DATA sxml_value_node     TYPE REF TO if_sxml_value_node.
    DATA string              TYPE string.
    DATA xstring             TYPE xstring.

    METHODS parse_ixml.

ENDCLASS.


CLASS ltc_ixml DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PROTECTED SECTION.

    DATA document       TYPE REF TO if_ixml_document.
    DATA element        TYPE REF TO if_ixml_element.
    DATA error          TYPE REF TO if_ixml_parse_error.
    DATA istream        TYPE REF TO if_ixml_istream.
    DATA ixml           TYPE REF TO if_ixml.
    DATA num_errors     TYPE i.
    DATA parser         TYPE REF TO if_ixml_parser.
    DATA rc             TYPE i.
    DATA reason         TYPE string.
    DATA stream_factory TYPE REF TO if_ixml_stream_factory.
    DATA string         TYPE string.
    DATA text           TYPE REF TO if_ixml_text.
    DATA xstring        TYPE xstring.

ENDCLASS.


CLASS ltc_ixml_complex DEFINITION
      INHERITING FROM ltc_ixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS bom FOR TESTING RAISING cx_static_check.
    METHODS reassign_to_other_parent FOR TESTING RAISING cx_static_check.
    METHODS invalid_multiple_root_elements FOR TESTING RAISING cx_static_check.

ENDCLASS.


CLASS ltc_ixml_element DEFINITION
      INHERITING FROM ltc_ixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS get_attribute_xmlns FOR TESTING RAISING cx_static_check.
    METHODS set_attribute_ns_name_xmlns FOR TESTING RAISING cx_static_check.

ENDCLASS.


CLASS ltc_ixml_node DEFINITION
      INHERITING FROM ltc_ixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS append_child FOR TESTING RAISING cx_static_check.
    METHODS append_child_not_bound FOR TESTING RAISING cx_static_check.

ENDCLASS.


CLASS ltc_ixml_parser DEFINITION
      INHERITING FROM ltc_ixml
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS add_preserve_space_element FOR TESTING RAISING cx_static_check.
    METHODS add_strip_space_element FOR TESTING RAISING cx_static_check.
    METHODS empty_xml FOR TESTING RAISING cx_static_check.
    METHODS end_tag_doesnt_match_begin_tag FOR TESTING RAISING cx_static_check.
    METHODS most_simple_valid_xml FOR TESTING RAISING cx_static_check.
    METHODS normalization FOR TESTING RAISING cx_static_check.
    METHODS set_validating_on FOR TESTING RAISING cx_static_check.

    METHODS init_parser.

ENDCLASS.


CLASS lth_ixml DEFINITION.

  PUBLIC SECTION.

    CLASS-DATA document       TYPE REF TO if_ixml_document.
    CLASS-DATA encoding       TYPE REF TO if_ixml_encoding.
    CLASS-DATA ixml           TYPE REF TO if_ixml.
    CLASS-DATA stream_factory TYPE REF TO if_ixml_stream_factory.

    CLASS-METHODS create_document.

    CLASS-METHODS parse
      IMPORTING
        iv_xml_string  TYPE csequence OPTIONAL
        iv_xml_xstring TYPE xsequence OPTIONAL
        iv_validating  TYPE i DEFAULT zif_excel_xml_parser=>co_no_validation
      PREFERRED PARAMETER
        iv_xml_string
      RETURNING
        VALUE(ro_result) TYPE REF TO if_ixml_document.

    CLASS-METHODS render
      RETURNING
        VALUE(rv_result) TYPE string.

ENDCLASS.


CLASS ltc_diff_ixml_sxml_parser IMPLEMENTATION.
  METHOD missing_namespace_declaration.
    string = '<a:A><B/></a:A>'.
    xstring = cl_abap_codepage=>convert_to( string ).
    sxml_reader = cl_sxml_string_reader=>create( xstring ).
    TRY.
        sxml_node = sxml_reader->read_next_node( ).
        cl_abap_unit_assert=>fail( msg = 'missing namespace declaration xmlns:a="..." -> should have failed' ).
      CATCH cx_sxml_parse_error.
        cl_abap_unit_assert=>assert_not_bound( act = sxml_node ).
        cl_abap_unit_assert=>assert_equals( act = sxml_reader->name
                                            exp = '' ).
    ENDTRY.
    sxml_open_element ?= sxml_reader->read_next_node( ).
    cl_abap_unit_assert=>assert_equals( act = sxml_open_element->qname-name
                                        exp = 'B' ).
  ENDMETHOD.

  METHOD parse_ixml.
    ixml = cl_ixml=>create( ).
    ixml_document = ixml->create_document( ).
    ixml_stream_factory = ixml->create_stream_factory( ).

    xstring = cl_abap_codepage=>convert_to( string ).
    ixml_istream = ixml_stream_factory->create_istream_xstring( xstring ).
    ixml_parser = ixml->create_parser( stream_factory = ixml_stream_factory
                                       istream        = ixml_istream
                                       document       = ixml_document ).
    ixml_parser->set_normalizing( abap_true ).
    ixml_parser->set_validating( mode = zif_excel_xml_parser=>co_no_validation ).
    ixml_parser->parse( ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_ixml IMPLEMENTATION.
ENDCLASS.


CLASS ltc_ixml_complex IMPLEMENTATION.
  METHOD bom.
    xstring = cl_abap_codepage=>convert_to(
        '<?xml version="1.0"  standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/20' &&
        '06/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.' &&
        'xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' ).
    CONCATENATE cl_abap_char_utilities=>byte_order_mark_utf8
                xstring
                INTO xstring
                IN BYTE MODE.
    document = lth_ixml=>parse( iv_xml_xstring = xstring ).
    element = document->get_root_element( ).
    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = 'Relationships' ).
  ENDMETHOD.

  METHOD invalid_multiple_root_elements.
    " As done in method CREATE_DOCPROPS_APP of class ZCL_EXCEL_WRITER_2007, e.g. LinksUpToDate and SharedDoc.

    DATA lo_element_root TYPE REF TO if_ixml_element.

    document = lth_ixml=>parse( '<A/>' ).
    lo_element_root = document->get_root_element( ).
    element = document->create_simple_element( name   = 'LinksUpToDate'
                                               parent = document ).
    element = document->create_simple_element( name   = 'SharedDoc'
                                               parent = document ).
    string = lth_ixml=>render( ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A/><LinksUpToDate/><SharedDoc/>' ).
  ENDMETHOD.

  METHOD reassign_to_other_parent.
    " As done in method CREATE_DOCPROPS_APP of class ZCL_EXCEL_WRITER_2007, e.g. LinksUpToDate and SharedDoc.

    DATA lo_element_root TYPE REF TO if_ixml_element.

    document = lth_ixml=>parse( '<A/>' ).
    lo_element_root = document->get_root_element( ).
    element = document->create_simple_element( name   = 'LinksUpToDate'
                                               parent = document ).
    lo_element_root->append_child( new_child = element ).
    element = document->create_simple_element( name   = 'SharedDoc'
                                               parent = document ).
    lo_element_root->append_child( new_child = element ).
    string = lth_ixml=>render( ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A><LinksUpToDate/><SharedDoc/></A>' ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_ixml_element IMPLEMENTATION.
  METHOD get_attribute_xmlns.
    " Method READ_THEME of class ZCL_EXCEL_THEME:
    "    CONSTANTS c_theme_xmlns TYPE string VALUE 'xmlns:a'.    "#EC NOTEXT
    "      xmls_a = lo_node_theme->get_attribute( name = c_theme_xmlns ).
    " NB: In fact, xmls_a IS NOT USED, it can be deleted!
    " NB: GET_ATTRIBUTE with NAME = 'xmlns:XXXX' always returns no attribute found!
    document = lth_ixml=>parse( '<A xmlns:nsprefix="nsuri" nsprefix:attr="A1"/>' ).
    element = document->get_root_element( ).
    string = element->get_attribute( name = 'xmlns:nsprefix' ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '' ).
  ENDMETHOD.

  METHOD set_attribute_ns_name_xmlns.
    " Two ways to set attributes from a namespace
    document = lth_ixml=>parse( '<A/>' ).
    element = document->get_root_element( ).
    element->set_attribute_ns( name  = 'xmlns:nsprefix'
                               value = 'nsuri' ).
    element->set_attribute_ns( name  = 'nsprefix:attr'
                               value = 'A1' ).
    string = lth_ixml=>render( ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A xmlns:nsprefix="nsuri" nsprefix:attr="A1"/>' ).

    " TODO: abaplint deps is missing parameter uri of set_attribute_ns
    document = lth_ixml=>parse( '<A/>' ).
    element = document->get_root_element( ).
    element->set_attribute_ns( name   = 'nsprefix'
                               prefix = 'xmlns'
                               uri    = ''
                               value  = 'nsuri' ).
    element->set_attribute_ns( name   = 'attr'
                               prefix = 'nsprefix'
                               uri    = 'nsuri'
                               value  = 'A1' ).
    string = lth_ixml=>render( ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A xmlns:nsprefix="nsuri" nsprefix:attr="A1"/>' ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_ixml_node IMPLEMENTATION.
  METHOD append_child.
    " NB: append to the root element is not the same as append to the document.
    DATA lo_element TYPE REF TO if_ixml_element.

    document = lth_ixml=>parse( '<A/>' ).
    element = document->get_root_element( ).
    lo_element = document->create_element( name = 'B' ).
    element->append_child( lo_element ).
    string = lth_ixml=>render( ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A><B/></A>' ).

    document = lth_ixml=>parse( '<A/>' ).
    element = document->get_root_element( ).
    lo_element = document->create_element( name = 'B' ).
    document->append_child( lo_element ).
    string = lth_ixml=>render( ).
    cl_abap_unit_assert=>assert_equals( act = string
                                        exp = '<A/><B/>' ).
  ENDMETHOD.

  METHOD append_child_not_bound.
    DATA rval       TYPE i.
    DATA lo_element TYPE REF TO if_ixml_element.

    document = lth_ixml=>parse( '<A/>' ).
    element = document->get_root_element( ).
    rval = element->append_child( lo_element ).
    cl_abap_unit_assert=>assert_equals( act = rval
                                        exp = zif_excel_xml_constants=>ixml_mr-dom_invalid_arg ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_ixml_parser IMPLEMENTATION.
  METHOD add_preserve_space_element.
    string = |  <A>  <B>   </B>  </A>  |.
    init_parser( ).
    rc = parser->parse( ).
    element = document->get_root_element( ).
    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = `A` ).
    cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                        exp = `   ` ).
    element ?= element->get_first_child( ).
    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = `B` ).
    cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                        exp = `   ` ).
  ENDMETHOD.

  METHOD add_strip_space_element.
    string = |  <A>  <B>  </B>  </A>  |.
    init_parser( ).
    parser->add_strip_space_element( ).
    rc = parser->parse( ).
    element = document->get_root_element( ).
    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = `A` ).
    cl_abap_unit_assert=>assert_equals( act = element->get_value( )
                                        exp = `` ).
    element ?= element->get_first_child( ).
  ENDMETHOD.

  METHOD empty_xml.
    CLEAR string.
    init_parser( ).
    rc = parser->parse( ).
    cl_abap_unit_assert=>assert_equals( act = rc
                                        exp = zif_excel_xml_constants=>ixml_mr-parser_error ).
    num_errors = parser->num_errors( ).
    cl_abap_unit_assert=>assert_equals( act = num_errors
                                        exp = 1 ).
    error = parser->get_error( index = 0 ).
    reason = error->get_reason( ).
    cl_abap_unit_assert=>assert_equals( act = reason
                                        exp = `unexpected end-of-file` ).
  ENDMETHOD.

  METHOD end_tag_doesnt_match_begin_tag.
    string = |<A></B>|.
    init_parser( ).
    rc = parser->parse( ).
    cl_abap_unit_assert=>assert_equals( act = rc
                                        exp = zif_excel_xml_constants=>ixml_mr-parser_error ).
    num_errors = parser->num_errors( ).
    cl_abap_unit_assert=>assert_equals( act = num_errors
                                        exp = 1 ).
    error = parser->get_error( index = 0 ).
    reason = error->get_reason( ).
    cl_abap_unit_assert=>assert_equals( act = reason
                                        exp = `end tag 'B' does not match begin tag 'A'` ).
  ENDMETHOD.

  METHOD init_parser.
    ixml = cl_ixml=>create( ).
    stream_factory = ixml->create_stream_factory( ).
    xstring = cl_abap_codepage=>convert_to( string ).
    istream = stream_factory->create_istream_xstring( xstring ).
    document = ixml->create_document( ).
    parser = ixml->create_parser( stream_factory = stream_factory
                                  istream        = istream
                                  document       = document ).
  ENDMETHOD.

  METHOD most_simple_valid_xml.
    string = |<A/>|.
    init_parser( ).
    rc = parser->parse( ).
    cl_abap_unit_assert=>assert_equals( act = rc
                                        exp = zif_excel_xml_constants=>ixml_mr-dom_ok ).
    num_errors = parser->num_errors( ).
    cl_abap_unit_assert=>assert_equals( act = num_errors
                                        exp = 0 ).
    element = document->get_root_element( ).
    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = `A` ).
  ENDMETHOD.

  METHOD normalization.
    string = |<A><B>  1  \n  2  </B></A>|.
    init_parser( ).
    parser->set_normalizing( abap_true ). " default = abap_true
    rc = parser->parse( ).
    element = document->get_root_element( ).
    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = `A` ).
    element ?= element->get_first_child( ).
    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = `B` ).
    text ?= element->get_first_child( ).
    cl_abap_unit_assert=>assert_equals( act = text->get_value( )
                                        exp = |1  \n  2| ).
  ENDMETHOD.

  METHOD set_validating_on.
    string = '<A></B>'.
    init_parser( ).
    parser->set_validating( if_ixml_parser=>co_validate ).
    rc = parser->parse( ).
    cl_abap_unit_assert=>assert_equals( act = rc
                                        exp = zif_excel_xml_constants=>ixml_mr-parser_error ).
    num_errors = parser->num_errors( ).
    cl_abap_unit_assert=>assert_equals( act = num_errors
                                        exp = 1 ).
    error = parser->get_error( index = 0 ).
    reason = error->get_reason( ).
    cl_abap_unit_assert=>assert_equals( act = reason
                                        exp = `no DTD specified, can't validate` ).
  ENDMETHOD.
ENDCLASS.


CLASS lth_ixml IMPLEMENTATION.
  METHOD create_document.
    ixml = cl_ixml=>create( ).
    document = ixml->create_document( ).
  ENDMETHOD.

  METHOD parse.
    DATA lv_xstring TYPE xstring.
    DATA lo_istream TYPE REF TO if_ixml_istream.
    DATA lo_parser  TYPE REF TO if_ixml_parser.

    " code inspired from the method GET_IXML_FROM_ZIP_ARCHIVE of ZCL_EXCEL_READER_2007.
    ixml = cl_ixml=>create( ).
    document = ixml->create_document( ).
    encoding = ixml->create_encoding( byte_order    = if_ixml_encoding=>co_none "co_platform_endian
                                      character_set = 'UTF-16' ).
    document->set_encoding( encoding ).
    stream_factory = ixml->create_stream_factory( ).

    IF iv_xml_string IS NOT INITIAL.
      lv_xstring = cl_abap_codepage=>convert_to( iv_xml_string ).
      lo_istream = stream_factory->create_istream_xstring( lv_xstring ).
    ELSE.
      lo_istream = stream_factory->create_istream_xstring( iv_xml_xstring ).
    ENDIF.
    lo_parser = ixml->create_parser( stream_factory = stream_factory
                                     istream        = lo_istream
                                     document       = document ).
    lo_parser->set_normalizing( abap_true ).
    lo_parser->set_validating( mode = iv_validating ).
    lo_parser->parse( ).

    ro_result = document.
  ENDMETHOD.

  METHOD render.
    DATA lo_ostream  TYPE REF TO if_ixml_ostream.
    DATA lo_renderer TYPE REF TO if_ixml_renderer.

    stream_factory = ixml->create_stream_factory( ).
    lo_ostream = stream_factory->create_ostream_cstring( rv_result ).
    lo_renderer = ixml->create_renderer( ostream  = lo_ostream
                                         document = document ).
    document->set_declaration( abap_false ).
    " Fills RV_RESULT
    lo_renderer->render( ).
    " remove the UTF-16 BOM (i.e. remove the first character)
    SHIFT rv_result LEFT BY 1 PLACES.
  ENDMETHOD.
ENDCLASS.
