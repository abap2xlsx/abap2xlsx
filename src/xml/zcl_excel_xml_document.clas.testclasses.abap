*"* use this source file for your ABAP unit test classes

CLASS ltc_parse_xstring DEFINITION
    FOR TESTING
    DURATION SHORT
    RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS simple FOR TESTING RAISING cx_static_check.

    DATA xml_document TYPE REF TO zcl_excel_xml_document.
    DATA document     TYPE REF TO zif_excel_xml_document.
    DATA element      TYPE REF TO zif_excel_xml_element.
    DATA ref_xstring  TYPE REF TO xstring.
    DATA xstring      TYPE xstring.
    DATA string       TYPE string.

ENDCLASS.


CLASS ltc_parse_xstring IMPLEMENTATION.
  METHOD simple.
* Method CREATE_CONTENT_TYPES of class ZCL_EXCEL_WRITER_XLSM:
*    CREATE OBJECT lo_document_xml.
*    lo_document_xml->parse_xstring( ep_content ).
*    lo_document ?= lo_document_xml->m_document.
    string = '<A/>'.
    xstring = cl_abap_codepage=>convert_to( string ).
    GET REFERENCE OF xstring INTO ref_xstring.

    CREATE OBJECT xml_document.
    xml_document->parse_xstring( ref_xstring ).
    document ?= xml_document->m_document.
    element ?= document->get_first_child( ).

    cl_abap_unit_assert=>assert_equals( act = element->get_name( )
                                        exp = 'A' ).
  ENDMETHOD.
ENDCLASS.
