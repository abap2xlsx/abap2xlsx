CLASS zcl_excel_theme DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    CONSTANTS c_theme_elements TYPE string VALUE 'themeElements'. "#EC NOTEXT
    CONSTANTS c_theme_object_def TYPE string VALUE 'objectDefaults'. "#EC NOTEXT
    CONSTANTS c_theme_extra_color TYPE string VALUE 'extraClrSchemeLst'. "#EC NOTEXT
    CONSTANTS c_theme_extlst TYPE string VALUE 'extLst'.    "#EC NOTEXT
    CONSTANTS c_theme TYPE string VALUE 'theme'.            "#EC NOTEXT
    CONSTANTS c_theme_name TYPE string VALUE 'name'.        "#EC NOTEXT
    CONSTANTS c_theme_xmlns TYPE string VALUE 'xmlns:a'.    "#EC NOTEXT
    CONSTANTS c_theme_prefix TYPE string VALUE 'a'.         "#EC NOTEXT
    CONSTANTS c_theme_prefix_write TYPE string VALUE 'a:'.  "#EC NOTEXT
    CONSTANTS c_theme_xmlns_val TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main'. "#EC NOTEXT

    METHODS constructor .
    METHODS read_theme
      IMPORTING
        VALUE(io_theme_xml) TYPE REF TO if_ixml_document .
    METHODS write_theme
      RETURNING
        VALUE(rv_xstring) TYPE xstring .
    METHODS set_color
      IMPORTING
        VALUE(iv_type)         TYPE string
        VALUE(iv_srgb)         TYPE zcl_excel_theme_color_scheme=>t_srgb OPTIONAL
        VALUE(iv_syscolorname) TYPE string OPTIONAL
        VALUE(iv_syscolorlast) TYPE zcl_excel_theme_color_scheme=>t_srgb OPTIONAL .
    METHODS set_color_scheme_name
      IMPORTING
        VALUE(iv_name) TYPE string .
    METHODS set_font
      IMPORTING
        VALUE(iv_type)     TYPE string
        VALUE(iv_script)   TYPE string
        VALUE(iv_typeface) TYPE string .
    METHODS set_latin_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS set_ea_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS set_cs_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS set_font_scheme_name
      IMPORTING
        VALUE(iv_name) TYPE string .
    METHODS set_theme_name
      IMPORTING
        VALUE(iv_name) TYPE string .
  PROTECTED SECTION.

    DATA elements TYPE REF TO zcl_excel_theme_elements .
    DATA objectdefaults TYPE REF TO zcl_excel_theme_objectdefaults .
    DATA extclrschemelst TYPE REF TO zcl_excel_theme_eclrschemelst .
    DATA extlst TYPE REF TO zcl_excel_theme_extlst .
  PRIVATE SECTION.

    DATA theme_changed TYPE abap_bool .
    DATA theme_read TYPE abap_bool .
    DATA name TYPE string .
    DATA xmls_a TYPE string .
ENDCLASS.



CLASS zcl_excel_theme IMPLEMENTATION.


  METHOD constructor.
    CREATE OBJECT elements.
    CREATE OBJECT objectdefaults.
    CREATE OBJECT extclrschemelst.
    CREATE OBJECT extlst.
  ENDMETHOD.                    "class_constructor


  METHOD read_theme.
    DATA: lo_node_theme TYPE REF TO if_ixml_element.
    DATA: lo_theme_children TYPE REF TO if_ixml_node_list.
    DATA: lo_theme_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_theme_element TYPE REF TO if_ixml_element.
    CHECK io_theme_xml IS NOT INITIAL.

    lo_node_theme  = io_theme_xml->get_root_element( )."   find_from_name( name = c_theme ).
    IF lo_node_theme IS BOUND.
      name = lo_node_theme->get_attribute( name = c_theme_name ).
      xmls_a = lo_node_theme->get_attribute( name = c_theme_xmlns ).
      lo_theme_children = lo_node_theme->get_children( ).
      lo_theme_iterator = lo_theme_children->create_iterator( ).
      lo_theme_element ?= lo_theme_iterator->get_next( ).
      WHILE lo_theme_element IS BOUND.
        CASE lo_theme_element->get_name( ).
          WHEN c_theme_elements.
            elements->load( io_elements = lo_theme_element ).
          WHEN c_theme_object_def.
            objectdefaults->load( io_object_def = lo_theme_element ).
          WHEN c_theme_extra_color.
            extclrschemelst->load( io_extra_color = lo_theme_element ).
          WHEN c_theme_extlst.
            extlst->load( io_extlst = lo_theme_element ).
        ENDCASE.
        lo_theme_element ?= lo_theme_iterator->get_next( ).
      ENDWHILE.
    ENDIF.
  ENDMETHOD.                    "read_theme


  METHOD set_color.
    elements->color_scheme->set_color(
      EXPORTING
        iv_type         = iv_type
        iv_srgb         = iv_srgb
        iv_syscolorname = iv_syscolorname
        iv_syscolorlast = iv_syscolorlast
    ).
  ENDMETHOD.                    "set_color


  METHOD set_color_scheme_name.
    elements->color_scheme->set_name( iv_name = iv_name ).
  ENDMETHOD.                    "set_color_scheme_name


  METHOD set_cs_font.
    elements->font_scheme->modify_cs_font(
      EXPORTING
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "set_cs_font


  METHOD set_ea_font.
    elements->font_scheme->modify_ea_font(
      EXPORTING
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "set_ea_font


  METHOD set_font.
    elements->font_scheme->modify_font(
      EXPORTING
        iv_type     = iv_type
        iv_script   = iv_script
        iv_typeface = iv_typeface
    ).
  ENDMETHOD.                    "set_font


  METHOD set_font_scheme_name.
    elements->font_scheme->set_name( iv_name = iv_name ).
  ENDMETHOD.                    "set_font_scheme_name


  METHOD set_latin_font.
    elements->font_scheme->modify_latin_font(
      EXPORTING
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "set_latin_font


  METHOD set_theme_name.
    name = iv_name.
  ENDMETHOD.


  METHOD write_theme.
    DATA: lo_ixml         TYPE REF TO if_ixml,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_encoding     TYPE REF TO if_ixml_encoding.
    DATA: lo_streamfactory  TYPE REF TO if_ixml_stream_factory.
    DATA: lo_ostream TYPE REF TO if_ixml_ostream.
    DATA: lo_renderer TYPE REF TO if_ixml_renderer.
    DATA: lo_document TYPE REF TO if_ixml_document.
    lo_ixml = cl_ixml=>create( ).

    lo_encoding = lo_ixml->create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
                                            character_set = 'UTF-8' ).
    lo_document = lo_ixml->create_document( ).
    lo_document->set_encoding( lo_encoding ).
    lo_document->set_standalone( abap_true ).
    lo_document->set_namespace_prefix( prefix = 'a' ).

    lo_element_root = lo_document->create_simple_element_ns( prefix = c_theme_prefix
                                                             name   = c_theme
                                                            parent = lo_document
                                                            ).
    lo_element_root->set_attribute_ns( name  = c_theme_xmlns
                                       value = c_theme_xmlns_val ).
    lo_element_root->set_attribute_ns( name  = c_theme_name
                                       value = name ).

    elements->build_xml( io_document = lo_document ).
    objectdefaults->build_xml( io_document = lo_document ).
    extclrschemelst->build_xml( io_document = lo_document ).
    extlst->build_xml( io_document = lo_document ).

    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = rv_xstring ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  ENDMETHOD.                    "write_theme
ENDCLASS.
