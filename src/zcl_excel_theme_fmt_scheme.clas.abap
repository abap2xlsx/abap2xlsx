CLASS zcl_excel_theme_fmt_scheme DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS load
      IMPORTING
        !io_fmt_scheme TYPE REF TO if_ixml_element .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA fmt_scheme TYPE REF TO if_ixml_element .

    METHODS get_default_fmt
      RETURNING
        VALUE(rv_string) TYPE string .

    METHODS parse_string
      IMPORTING iv_string      TYPE string
      RETURNING VALUE(ri_node) TYPE REF TO if_ixml_node.
ENDCLASS.



CLASS ZCL_EXCEL_THEME_FMT_SCHEME IMPLEMENTATION.


  METHOD build_xml.
    DATA: lo_node TYPE REF TO if_ixml_node.
    DATA: lo_elements TYPE REF TO if_ixml_element.
    CHECK io_document IS BOUND.
    lo_elements ?= io_document->find_from_name_ns( name = zcl_excel_theme=>c_theme_elements ).
    IF lo_elements IS BOUND.

      IF fmt_scheme IS INITIAL.
        lo_node = parse_string( get_default_fmt( ) ).
        lo_elements->append_child( new_child = lo_node ).
      ELSE.
        lo_elements->append_child( new_child = fmt_scheme ).
      ENDIF.
    ENDIF.
  ENDMETHOD.                    "build_xml


  METHOD get_default_fmt.
    CONCATENATE    '<a:fmtScheme name="Office">'
    '      <a:fillStyleLst>'
    '        <a:solidFill>'
    '          <a:schemeClr val="phClr"/>'
    '        </a:solidFill>'
    '        <a:gradFill rotWithShape="1">'
    '          <a:gsLst>'
    '            <a:gs pos="0">'
    '              <a:schemeClr val="phClr">'
    '                <a:lumMod val="110000"/>'
    '                <a:satMod val="105000"/>'
    '                <a:tint val="67000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '            <a:gs pos="50000">'
    '              <a:schemeClr val="phClr">'
    '                <a:lumMod val="105000"/>'
    '                <a:satMod val="103000"/>'
    '                <a:tint val="73000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '            <a:gs pos="100000">'
    '              <a:schemeClr val="phClr">'
    '                <a:lumMod val="105000"/>'
    '                <a:satMod val="109000"/>'
    '                <a:tint val="81000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '          </a:gsLst>'
    '          <a:lin ang="5400000" scaled="0"/>'
    '        </a:gradFill>'
    '        <a:gradFill rotWithShape="1">'
    '          <a:gsLst>'
    '            <a:gs pos="0">'
    '              <a:schemeClr val="phClr">'
    '                <a:satMod val="103000"/>'
    '                <a:lumMod val="102000"/>'
    '                <a:tint val="94000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '            <a:gs pos="50000">'
    '              <a:schemeClr val="phClr">'
    '                <a:satMod val="110000"/>'
    '                <a:lumMod val="100000"/>'
    '                <a:shade val="100000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '            <a:gs pos="100000">'
    '              <a:schemeClr val="phClr">'
    '                <a:lumMod val="99000"/>'
    '                <a:satMod val="120000"/>'
    '                <a:shade val="78000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '          </a:gsLst>'
    '          <a:lin ang="5400000" scaled="0"/>'
    '        </a:gradFill>'
    '      </a:fillStyleLst>'
    '      <a:lnStyleLst>'
    '        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">'
    '          <a:solidFill>'
    '            <a:schemeClr val="phClr"/>'
    '          </a:solidFill>'
    '          <a:prstDash val="solid"/>'
    '          <a:miter lim="800000"/>'
    '        </a:ln>'
    '        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">'
    '          <a:solidFill>'
    '            <a:schemeClr val="phClr"/>'
    '          </a:solidFill>'
    '          <a:prstDash val="solid"/>'
    '          <a:miter lim="800000"/>'
    '        </a:ln>'
    '        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">'
    '          <a:solidFill>'
    '            <a:schemeClr val="phClr"/>'
    '          </a:solidFill>'
    '          <a:prstDash val="solid"/>'
    '          <a:miter lim="800000"/>'
    '        </a:ln>'
    '      </a:lnStyleLst>'
    '      <a:effectStyleLst>'
    '        <a:effectStyle>'
    '          <a:effectLst/>'
    '        </a:effectStyle>'
    '        <a:effectStyle>'
    '          <a:effectLst/>'
    '        </a:effectStyle>'
    '        <a:effectStyle>'
    '          <a:effectLst>'
    '            <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">'
    '              <a:srgbClr val="000000">'
    '                <a:alpha val="63000"/>'
    '              </a:srgbClr>'
    '            </a:outerShdw>'
    '          </a:effectLst>'
    '        </a:effectStyle>'
    '      </a:effectStyleLst>'
    '      <a:bgFillStyleLst>'
    '        <a:solidFill>'
    '          <a:schemeClr val="phClr"/>'
    '        </a:solidFill>'
    '        <a:solidFill>'
    '          <a:schemeClr val="phClr">'
    '            <a:tint val="95000"/>'
    '            <a:satMod val="170000"/>'
    '          </a:schemeClr>'
    '        </a:solidFill>'
    '        <a:gradFill rotWithShape="1">'
    '          <a:gsLst>'
    '            <a:gs pos="0">'
    '              <a:schemeClr val="phClr">'
    '                <a:tint val="93000"/>'
    '                <a:satMod val="150000"/>'
    '                <a:shade val="98000"/>'
    '                <a:lumMod val="102000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '            <a:gs pos="50000">'
    '              <a:schemeClr val="phClr">'
    '                <a:tint val="98000"/>'
    '                <a:satMod val="130000"/>'
    '                <a:shade val="90000"/>'
    '                <a:lumMod val="103000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '            <a:gs pos="100000">'
    '              <a:schemeClr val="phClr">'
    '                <a:shade val="63000"/>'
    '                <a:satMod val="120000"/>'
    '              </a:schemeClr>'
    '            </a:gs>'
    '          </a:gsLst>'
    '          <a:lin ang="5400000" scaled="0"/>'
    '        </a:gradFill>'
    '      </a:bgFillStyleLst>'
    '    </a:fmtScheme>'
    INTO rv_string .
  ENDMETHOD.                    "get_default_fmt


  METHOD load.
    fmt_scheme = zcl_excel_common=>clone_ixml_with_namespaces( io_fmt_scheme ).
  ENDMETHOD.                    "load


  METHOD parse_string.
    DATA li_stream   TYPE REF TO if_ixml_istream.
    DATA li_ixml     TYPE REF TO if_ixml.
    DATA li_document TYPE REF TO if_ixml_document.
    DATA li_factory  TYPE REF TO if_ixml_stream_factory.
    DATA li_parser   TYPE REF TO if_ixml_parser.
    DATA li_istream  TYPE REF TO if_ixml_istream.

    li_ixml = cl_ixml=>create( ).
    li_document = li_ixml->create_document( ).
    li_factory = li_ixml->create_stream_factory( ).
    li_istream = li_factory->create_istream_string( iv_string ).
    li_parser = li_ixml->create_parser(
      stream_factory = li_factory
      istream        = li_istream
      document       = li_document ).
    li_parser->add_strip_space_element( ).
    li_parser->parse( ).
    li_istream->close( ).
    ri_node = li_document->get_first_child( ).

  ENDMETHOD.
ENDCLASS.
