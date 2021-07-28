CLASS zcl_excel_theme_elements DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC

  GLOBAL FRIENDS zcl_excel_theme .

  PUBLIC SECTION.

    CONSTANTS c_color_scheme TYPE string VALUE 'clrScheme'. "#EC NOTEXT
    CONSTANTS c_font_scheme TYPE string VALUE 'fontScheme'. "#EC NOTEXT
    CONSTANTS c_fmt_scheme TYPE string VALUE 'fmtScheme'.   "#EC NOTEXT
    CONSTANTS c_theme_elements TYPE string VALUE 'themeElements'. "#EC NOTEXT

    METHODS constructor .
    METHODS load
      IMPORTING
        !io_elements TYPE REF TO if_ixml_element .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
  PROTECTED SECTION.

    DATA color_scheme TYPE REF TO zcl_excel_theme_color_scheme .
    DATA font_scheme TYPE REF TO zcl_excel_theme_font_scheme .
    DATA fmt_scheme TYPE REF TO zcl_excel_theme_fmt_scheme .
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_theme_elements IMPLEMENTATION.


  METHOD build_xml.
    DATA: lo_theme_element TYPE REF TO if_ixml_element.
    DATA: lo_theme TYPE REF TO if_ixml_element.
    CHECK io_document IS BOUND.
    lo_theme ?= io_document->get_root_element( ).
    IF lo_theme IS BOUND.
      lo_theme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                 name   = c_theme_elements
                                                              parent = lo_theme ).

      color_scheme->build_xml( io_document = io_document ).
      font_scheme->build_xml( io_document = io_document ).
      fmt_scheme->build_xml( io_document = io_document ).
    ENDIF.
  ENDMETHOD.


  METHOD constructor.
    CREATE OBJECT color_scheme.
    CREATE OBJECT font_scheme.
    CREATE OBJECT fmt_scheme.
  ENDMETHOD.                    "constructor


  METHOD load.
    DATA: lo_elements_children TYPE REF TO if_ixml_node_list.
    DATA: lo_elements_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_elements_element TYPE REF TO if_ixml_element.
    CHECK io_elements IS NOT INITIAL.

    lo_elements_children = io_elements->get_children( ).
    lo_elements_iterator = lo_elements_children->create_iterator( ).
    lo_elements_element ?= lo_elements_iterator->get_next( ).
    WHILE lo_elements_element IS BOUND.
      CASE lo_elements_element->get_name( ).
        WHEN c_color_scheme.
          color_scheme->load( io_color_scheme = lo_elements_element ).
        WHEN c_font_scheme.
          font_scheme->load( io_font_scheme = lo_elements_element ).
        WHEN c_fmt_scheme.
          fmt_scheme->load( io_fmt_scheme = lo_elements_element ).
      ENDCASE.
      lo_elements_element ?= lo_elements_iterator->get_next( ).
    ENDWHILE.
  ENDMETHOD.                    "load
ENDCLASS.
