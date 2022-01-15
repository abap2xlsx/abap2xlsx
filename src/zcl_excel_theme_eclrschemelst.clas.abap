CLASS zcl_excel_theme_eclrschemelst DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS load
      IMPORTING
        !io_extra_color TYPE REF TO if_ixml_element .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA extracolor TYPE REF TO if_ixml_element .
ENDCLASS.



CLASS zcl_excel_theme_eclrschemelst IMPLEMENTATION.


  METHOD build_xml.
    DATA: lo_theme TYPE REF TO if_ixml_element.
    DATA: lo_theme_objdef TYPE REF TO if_ixml_element.
    CHECK io_document IS BOUND.
    lo_theme ?= io_document->get_root_element( ).
    CHECK lo_theme IS BOUND.
    IF extracolor IS INITIAL.
      lo_theme_objdef ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                name   = zcl_excel_theme=>c_theme_extra_color
                                                                parent = lo_theme ).

    ELSE.
      lo_theme->append_child( new_child = extracolor ).
    ENDIF.

  ENDMETHOD.                    "build_xml


  METHOD load.
    extracolor = zcl_excel_common=>clone_ixml_with_namespaces( io_extra_color ).
  ENDMETHOD.                    "load
ENDCLASS.
