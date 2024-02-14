CLASS ltcl_test DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS.
  PRIVATE SECTION.
    METHODS build_xml FOR TESTING.

    DATA mi_ixml TYPE REF TO zif_excel_xml.
    DATA mi_document TYPE REF TO zif_excel_xml_document.
    METHODS setup.
    METHODS render
      RETURNING
        VALUE(rv_xml) TYPE string.
ENDCLASS.


CLASS ltcl_test IMPLEMENTATION.

  METHOD setup.
    mi_ixml = zcl_excel_xml=>create( ).
    mi_document = mi_ixml->create_document( ).
  ENDMETHOD.

  METHOD render.
    DATA li_ostream  TYPE REF TO zif_excel_xml_ostream.
    DATA li_renderer TYPE REF TO zif_excel_xml_renderer.
    DATA li_factory  TYPE REF TO zif_excel_xml_stream_factory.
    DATA lr_xml      TYPE REF TO string.

    li_factory = mi_ixml->create_stream_factory( ).
    GET REFERENCE OF rv_xml INTO lr_xml.
    li_ostream = li_factory->create_ostream_cstring( lr_xml ).
    li_renderer = mi_ixml->create_renderer(
      ostream  = li_ostream
      document = mi_document ).
    li_renderer->render( ).
  ENDMETHOD.

  METHOD build_xml.
    DATA lo_theme_fmt TYPE REF TO zcl_excel_theme_fmt_scheme.
    DATA li_ixml      TYPE REF TO zif_excel_xml.
    DATA li_document  TYPE REF TO zif_excel_xml_document.
    DATA lv_xml       TYPE string.

    mi_document->create_simple_element(
      name   = zcl_excel_theme=>c_theme_elements
      parent = mi_document ).

    CREATE OBJECT lo_theme_fmt.
    lo_theme_fmt->build_xml( mi_document ).

    lv_xml = render( ).

    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml
      exp = '*<a:fmtScheme name="Office" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">*' ).
  ENDMETHOD.

ENDCLASS.
