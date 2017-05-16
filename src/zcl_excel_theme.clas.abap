class ZCL_EXCEL_THEME definition
  public
  create public .

public section.

  constants C_THEME_ELEMENTS type STRING value 'themeElements'. "#EC NOTEXT
  constants C_THEME_OBJECT_DEF type STRING value 'objectDefaults'. "#EC NOTEXT
  constants C_THEME_EXTRA_COLOR type STRING value 'extraClrSchemeLst'. "#EC NOTEXT
  constants C_THEME_EXTLST type STRING value 'extLst'. "#EC NOTEXT
  constants C_THEME type STRING value 'theme'. "#EC NOTEXT
  constants C_THEME_NAME type STRING value 'name'. "#EC NOTEXT
  constants C_THEME_XMLNS type STRING value 'xmlns:a'. "#EC NOTEXT
  constants C_THEME_PREFIX type STRING value 'a'. "#EC NOTEXT
  constants C_THEME_PREFIX_WRITE type STRING value 'a:'. "#EC NOTEXT
  constants C_THEME_XMLNS_VAL type STRING value 'http://schemas.openxmlformats.org/drawingml/2006/main'. "#EC NOTEXT

  methods CONSTRUCTOR .
  methods READ_THEME
    importing
      value(IO_THEME_XML) type ref to IF_IXML_DOCUMENT .
  methods WRITE_THEME
    returning
      value(RV_XSTRING) type XSTRING .
  methods SET_COLOR
    importing
      value(IV_TYPE) type STRING
      value(IV_SRGB) type ZCL_EXCEL_THEME_COLOR_SCHEME=>T_SRGB optional
      value(IV_SYSCOLORNAME) type STRING optional
      value(IV_SYSCOLORLAST) type ZCL_EXCEL_THEME_COLOR_SCHEME=>T_SRGB optional .
  methods SET_COLOR_SCHEME_NAME
    importing
      value(IV_NAME) type STRING .
  methods SET_FONT
    importing
      value(IV_TYPE) type STRING
      value(IV_SCRIPT) type STRING
      value(IV_TYPEFACE) type STRING .
  methods SET_LATIN_FONT
    importing
      value(IV_TYPE) type STRING
      value(IV_TYPEFACE) type STRING
      value(IV_PANOSE) type STRING optional
      value(IV_PITCHFAMILY) type STRING optional
      value(IV_CHARSET) type STRING optional .
  methods SET_EA_FONT
    importing
      value(IV_TYPE) type STRING
      value(IV_TYPEFACE) type STRING
      value(IV_PANOSE) type STRING optional
      value(IV_PITCHFAMILY) type STRING optional
      value(IV_CHARSET) type STRING optional .
  methods SET_CS_FONT
    importing
      value(IV_TYPE) type STRING
      value(IV_TYPEFACE) type STRING
      value(IV_PANOSE) type STRING optional
      value(IV_PITCHFAMILY) type STRING optional
      value(IV_CHARSET) type STRING optional .
  methods SET_FONT_SCHEME_NAME
    importing
      value(IV_NAME) type STRING .
  methods SET_THEME_NAME
    importing
      value(IV_NAME) type STRING .
protected section.

  data ELEMENTS type ref to ZCL_EXCEL_THEME_ELEMENTS .
  data OBJECTDEFAULTS type ref to ZCL_EXCEL_THEME_OBJECTDEFAULTS .
  data EXTCLRSCHEMELST type ref to ZCL_EXCEL_THEME_ECLRSCHEMELST .
  data EXTLST type ref to ZCL_EXCEL_THEME_EXTLST .
private section.

  data THEME_CHANGED type ABAP_BOOL .
  data THEME_READ type ABAP_BOOL .
  data NAME type STRING .
  data XMLS_A type STRING .
ENDCLASS.



CLASS ZCL_EXCEL_THEME IMPLEMENTATION.


method constructor.
    create object elements.
    create object objectdefaults.
    create object extclrschemelst.
    create object extlst.
  endmethod.                    "class_constructor


method read_theme.
    data: lo_node_theme type ref to if_ixml_element.
    data: lo_theme_children type ref to if_ixml_node_list.
    data: lo_theme_iterator type ref to if_ixml_node_iterator.
    data: lo_theme_element type ref to if_ixml_element.
    check io_theme_xml is not initial.

    lo_node_theme  = io_theme_xml->get_root_element( )."   find_from_name( name = c_theme ).
    if lo_node_theme is bound.
      name = lo_node_theme->get_attribute( name = c_theme_name ).
      xmls_a = lo_node_theme->get_attribute( name = c_theme_xmlns ).
      lo_theme_children = lo_node_theme->get_children( ).
      lo_theme_iterator = lo_theme_children->create_iterator( ).
      lo_theme_element ?= lo_theme_iterator->get_next( ).
      while lo_theme_element is bound.
        case lo_theme_element->get_name( ).
          when c_theme_elements.
            elements->load( io_elements = lo_theme_element ).
          when c_theme_object_def.
            objectdefaults->load( io_object_def = lo_theme_element ).
          when c_theme_extra_color.
            extclrschemelst->load( io_extra_color = lo_theme_element ).
          when c_theme_extlst.
            extlst->load( io_extlst = lo_theme_element ).
        endcase.
        lo_theme_element ?= lo_theme_iterator->get_next( ).
      endwhile.
    endif.
  endmethod.                    "read_theme


method set_color.
    elements->color_scheme->set_color(
      exporting
        iv_type         = iv_type
        iv_srgb         = iv_srgb
        iv_syscolorname = iv_syscolorname
        iv_syscolorlast = iv_syscolorlast
    ).
  endmethod.                    "set_color


method set_color_scheme_name.
    elements->color_scheme->set_name( iv_name = iv_name ).
  endmethod.                    "set_color_scheme_name


method set_cs_font.
    elements->font_scheme->modify_cs_font(
      exporting
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  endmethod.                    "set_cs_font


method set_ea_font.
    elements->font_scheme->modify_ea_font(
      exporting
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  endmethod.                    "set_ea_font


method set_font.
    elements->font_scheme->modify_font(
      exporting
        iv_type     = iv_type
        iv_script   = iv_script
        iv_typeface = iv_typeface
    ).
  endmethod.                    "set_font


method set_font_scheme_name.
    elements->font_scheme->set_name( iv_name = iv_name ).
  endmethod.                    "set_font_scheme_name


method set_latin_font.
    elements->font_scheme->modify_latin_font(
      exporting
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  endmethod.                    "set_latin_font


method set_theme_name.
    name = iv_name.
  endmethod.


method write_theme.
    data:   lo_ixml           type ref to if_ixml,
            lo_element_root   type ref to if_ixml_element,
            lo_encoding       type ref to if_ixml_encoding.
    data: lo_streamfactory  TYPE REF TO if_ixml_stream_factory.
    data: lo_ostream TYPE REF TO if_ixml_ostream.
    data: lo_renderer TYPE REF TO if_ixml_renderer.
    data: lo_document type ref to if_ixml_document.
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

  endmethod.                    "write_theme
ENDCLASS.
