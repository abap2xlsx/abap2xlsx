class ZCL_EXCEL_THEME_FMT_SCHEME definition
  public
  final
  create public .

public section.

  methods LOAD
    importing
      !IO_FMT_SCHEME type ref to IF_IXML_ELEMENT .
  methods BUILD_XML
    importing
      !IO_DOCUMENT type ref to IF_IXML_DOCUMENT .
protected section.
private section.

  data FMT_SCHEME type ref to IF_IXML_ELEMENT .

  methods GET_DEFAULT_FMT
    returning
      value(RV_STRING) type STRING .
ENDCLASS.



CLASS ZCL_EXCEL_THEME_FMT_SCHEME IMPLEMENTATION.


method build_xml.
    data: lo_xml type ref to cl_xml_document.
    data: lo_node type ref to if_ixml_node.
    data: lo_elements type ref to if_ixml_element.
    check io_document is bound.
    lo_elements ?= io_document->find_from_name_ns( name = zcl_excel_theme=>c_theme_elements ).
    if lo_elements is bound.

      if fmt_scheme is initial.
        create object lo_xml.
        lo_xml->parse_string( get_default_fmt( ) ).
        lo_node = lo_xml->get_first_node( ).
        lo_elements->append_child( new_child = lo_node ).
      else.
        lo_elements->append_child( new_child = fmt_scheme ).
      endif.
    endif.
  endmethod.                    "build_xml


method get_default_fmt.
    concatenate    '<a:fmtScheme name="Office">'
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
    into rv_string .
  endmethod.                    "get_default_fmt


method load.
    "! so far copy only existing values
    fmt_scheme ?= io_fmt_scheme.
  endmethod.                    "load
ENDCLASS.
