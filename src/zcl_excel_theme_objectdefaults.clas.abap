class ZCL_EXCEL_THEME_OBJECTDEFAULTS definition
  public
  final
  create public .

public section.

  methods LOAD
    importing
      !IO_OBJECT_DEF type ref to IF_IXML_ELEMENT .
  methods BUILD_XML
    importing
      !IO_DOCUMENT type ref to IF_IXML_DOCUMENT .
protected section.
private section.

  data OBJECTDEFAULTS type ref to IF_IXML_ELEMENT .
ENDCLASS.



CLASS ZCL_EXCEL_THEME_OBJECTDEFAULTS IMPLEMENTATION.


method build_xml.
    data: lo_theme type ref to if_ixml_element.
    data: lo_theme_objdef type ref to if_ixml_element.
    check io_document is bound.
    lo_theme ?= io_document->get_root_element( ).
    check lo_theme is bound.
    if objectdefaults is initial.
      lo_theme_objdef ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                name   = zcl_excel_theme=>c_theme_object_def
                                                                parent = lo_theme ).
    else.
      lo_theme->append_child( new_child = objectdefaults ).
    endif.
  endmethod.                    "build_xml


method load.
    "! so far copy only existing values
    objectdefaults ?= io_object_def.
  endmethod.                    "load
ENDCLASS.
