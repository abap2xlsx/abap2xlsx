class ZCL_EXCEL_THEME_ELEMENTS definition
  public
  final
  create public

  global friends ZCL_EXCEL_THEME .

public section.

  constants C_COLOR_SCHEME type STRING value 'clrScheme'. "#EC NOTEXT
  constants C_FONT_SCHEME type STRING value 'fontScheme'. "#EC NOTEXT
  constants C_FMT_SCHEME type STRING value 'fmtScheme'. "#EC NOTEXT
  constants C_THEME_ELEMENTS type STRING value 'themeElements'. "#EC NOTEXT

  methods CONSTRUCTOR .
  methods LOAD
    importing
      !IO_ELEMENTS type ref to IF_IXML_ELEMENT .
  methods BUILD_XML
    importing
      !IO_DOCUMENT type ref to IF_IXML_DOCUMENT .
protected section.

  data COLOR_SCHEME type ref to ZCL_EXCEL_THEME_COLOR_SCHEME .
  data FONT_SCHEME type ref to ZCL_EXCEL_THEME_FONT_SCHEME .
  data FMT_SCHEME type ref to ZCL_EXCEL_THEME_FMT_SCHEME .
private section.
ENDCLASS.



CLASS ZCL_EXCEL_THEME_ELEMENTS IMPLEMENTATION.


method build_xml.
    data: lo_theme_element type ref to if_ixml_element.
    data: lo_theme type ref to if_ixml_element.
    check io_document is bound.
    lo_theme ?= io_document->get_root_element( ).
    if lo_theme is bound.
      lo_theme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                 name   = c_theme_elements
                                                              parent = lo_theme ).

      color_scheme->build_xml( io_document = io_document ).
      font_scheme->build_xml( io_document = io_document ).
      fmt_scheme->build_xml( io_document = io_document ).
    endif.
  endmethod.


method constructor.
    create object color_scheme.
    create object font_scheme.
    create object fmt_scheme.
  endmethod.                    "constructor


method load.
    data: lo_elements_children type ref to if_ixml_node_list.
    data: lo_elements_iterator type ref to if_ixml_node_iterator.
    data: lo_elements_element type ref to if_ixml_element.
    check io_elements is not initial.

    lo_elements_children = io_elements->get_children( ).
    lo_elements_iterator = lo_elements_children->create_iterator( ).
    lo_elements_element ?= lo_elements_iterator->get_next( ).
    while lo_elements_element is bound.
      case lo_elements_element->get_name( ).
        when c_color_scheme.
            color_scheme->load( io_color_scheme = lo_elements_element ).
        when c_font_scheme.
            font_scheme->load( io_font_scheme = lo_elements_element ).
        when c_fmt_scheme.
            fmt_scheme->load( io_fmt_scheme = lo_elements_element ).
      endcase.
      lo_elements_element ?= lo_elements_iterator->get_next( ).
    endwhile.
  endmethod.                    "load
ENDCLASS.
