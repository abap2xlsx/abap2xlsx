class ZCL_EXCEL_THEME_COLOR_SCHEME definition
  public
  final
  create public

  global friends ZCL_EXCEL_THEME
                 ZCL_EXCEL_THEME_ELEMENTS .

public section.

  types T_SRGB type STRING .
  types:
    begin of t_syscolor,
              val type string,
              lastclr type t_srgb,
             end of t_syscolor .
  types:
    begin of t_color,
             srgb type t_srgb,
             syscolor type t_syscolor,
             end of t_color .

  constants C_DARK1 type STRING value 'dk1'. "#EC NOTEXT
  constants C_DARK2 type STRING value 'dk2'. "#EC NOTEXT
  constants C_LIGHT1 type STRING value 'lt1'. "#EC NOTEXT
  constants C_LIGHT2 type STRING value 'lt2'. "#EC NOTEXT
  constants C_ACCENT1 type STRING value 'accent1'. "#EC NOTEXT
  constants C_ACCENT2 type STRING value 'accent2'. "#EC NOTEXT
  constants C_ACCENT3 type STRING value 'accent3'. "#EC NOTEXT
  constants C_ACCENT4 type STRING value 'accent4'. "#EC NOTEXT
  constants C_ACCENT5 type STRING value 'accent5'. "#EC NOTEXT
  constants C_ACCENT6 type STRING value 'accent6'. "#EC NOTEXT
  constants C_HLINK type STRING value 'hlink'. "#EC NOTEXT
  constants C_FOLHLINK type STRING value 'folHlink'. "#EC NOTEXT
  constants C_SYSCOLOR type STRING value 'sysClr'. "#EC NOTEXT
  constants C_SRGBCOLOR type STRING value 'srgbClr'. "#EC NOTEXT
  constants C_VAL type STRING value 'val'. "#EC NOTEXT
  constants C_LASTCLR type STRING value 'lastClr'. "#EC NOTEXT
  constants C_NAME type STRING value 'name'. "#EC NOTEXT
  constants C_SCHEME type STRING value 'clrScheme'. "#EC NOTEXT

  methods LOAD
    importing
      !IO_COLOR_SCHEME type ref to IF_IXML_ELEMENT .
  methods SET_COLOR
    importing
      value(IV_TYPE) type STRING
      value(IV_SRGB) type T_SRGB optional
      value(IV_SYSCOLORNAME) type STRING optional
      value(IV_SYSCOLORLAST) type T_SRGB .
  methods BUILD_XML
    importing
      !IO_DOCUMENT type ref to IF_IXML_DOCUMENT .
  methods CONSTRUCTOR .
  methods SET_NAME
    importing
      value(IV_NAME) type STRING .
protected section.

  data NAME type STRING .
  data DARK1 type T_COLOR .
  data DARK2 type T_COLOR .
  data LIGHT1 type T_COLOR .
  data LIGHT2 type T_COLOR .
  data ACCENT1 type T_COLOR .
  data ACCENT2 type T_COLOR .
  data ACCENT3 type T_COLOR .
  data ACCENT4 type T_COLOR .
  data ACCENT5 type T_COLOR .
  data ACCENT6 type T_COLOR .
  data HLINK type T_COLOR .
  data FOLHLINK type T_COLOR .
private section.

  methods GET_COLOR
    importing
      !IO_OBJECT type ref to IF_IXML_ELEMENT
    returning
      value(RV_COLOR) type T_COLOR .
  methods SET_DEFAULTS .
ENDCLASS.



CLASS ZCL_EXCEL_THEME_COLOR_SCHEME IMPLEMENTATION.


method build_xml.
    data: lo_scheme_element type ref to if_ixml_element.
    data: lo_color type ref to if_ixml_element.
    data: lo_syscolor type ref to if_ixml_element.
    data: lo_srgb type ref to if_ixml_element.
    data: lo_elements type ref to if_ixml_element.

    check io_document is bound.
    lo_elements ?= io_document->find_from_name_ns( name   = zcl_excel_theme=>c_theme_elements ).
    if lo_elements is bound.
      lo_scheme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                  name   = zcl_excel_theme_elements=>c_color_scheme
                                                               parent = lo_elements ).
      lo_scheme_element->set_attribute( name = c_name value = name ).

      "! Adding colors to scheme
      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                  name   = c_dark1
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if dark1-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = dark1-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = dark1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = dark1-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_light1
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if light1-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = light1-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = light1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = light1-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_dark2
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if dark2-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = dark2-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = dark2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = dark2-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_light2
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if light2-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = light2-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = light2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = light2-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent1
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if accent1-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent1-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent1-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent2
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if accent2-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent2-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent2-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent3
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if accent3-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent3-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent3-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent3-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent4
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if accent4-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent4-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent4-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent4-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent5
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if accent5-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent5-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent5-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent5-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent6
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if accent6-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent6-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent6-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent6-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_hlink
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if hlink-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = hlink-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = hlink-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = hlink-syscolor-lastclr ).
        endif.
        clear: lo_color, lo_srgb, lo_syscolor.
      endif.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_folhlink
                                                      parent = lo_scheme_element ).
      if lo_color is bound.
        if folhlink-srgb is not initial.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = folhlink-srgb ).
        else.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = folhlink-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = folhlink-syscolor-lastclr ).
        endif.
      endif.


    endif.
  endmethod.                    "build_xml


method constructor.
    set_defaults( ).
  endmethod.                    "constructor


method get_color.
    data: lo_color_children type ref to if_ixml_node_list.
    data: lo_color_iterator type ref to if_ixml_node_iterator.
    data: lo_color_element type ref to if_ixml_element.
    check io_object  is not initial.

    lo_color_children = io_object->get_children( ).
    lo_color_iterator = lo_color_children->create_iterator( ).
    lo_color_element ?= lo_color_iterator->get_next( ).
    if lo_color_element is bound.
      case lo_color_element->get_name( ).
        when c_srgbcolor.
          rv_color-srgb = lo_color_element->get_attribute( name = c_val ).
        when c_syscolor.
          rv_color-syscolor-val = lo_color_element->get_attribute( name = c_val ).
          rv_color-syscolor-lastclr = lo_color_element->get_attribute( name = c_lastclr ).
      endcase.
    endif.
  endmethod.                    "get_color


method load.
    data: lo_scheme_children type ref to if_ixml_node_list.
    data: lo_scheme_iterator type ref to if_ixml_node_iterator.
    data: lo_scheme_element type ref to if_ixml_element.
    check io_color_scheme  is not initial.

    name = io_color_scheme->get_attribute( name = c_name ).
    lo_scheme_children = io_color_scheme->get_children( ).
    lo_scheme_iterator = lo_scheme_children->create_iterator( ).
    lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    while lo_scheme_element is bound.
      case lo_scheme_element->get_name( ).
        when c_dark1.
          dark1 = me->get_color( lo_scheme_element ).
        when c_dark2.
          dark2 = me->get_color( lo_scheme_element ).
        when c_light1.
          light1 = me->get_color( lo_scheme_element ).
        when c_light2.
          light2 = me->get_color( lo_scheme_element ).
        when c_accent1.
          accent1 = me->get_color( lo_scheme_element ).
        when c_accent2.
          accent2 = me->get_color( lo_scheme_element ).
        when c_accent3.
          accent3 = me->get_color( lo_scheme_element ).
        when c_accent4.
          accent4 = me->get_color( lo_scheme_element ).
        when c_accent5.
          accent5 = me->get_color( lo_scheme_element ).
        when c_accent6.
          accent6 = me->get_color( lo_scheme_element ).
        when c_hlink.
          hlink = me->get_color( lo_scheme_element ).
        when c_folhlink.
          folhlink = me->get_color( lo_scheme_element ).
      endcase.
      lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    endwhile.
  endmethod.                    "load


method set_color.
    field-symbols: <color> type t_color.
    check iv_type is not initial.
    check iv_srgb is not initial or  iv_syscolorname is not initial.
    case iv_type.
      when c_dark1.
        assign dark1 to <color>.
      when c_dark2.
        assign dark2 to <color>.
      when c_light1.
        assign light1 to <color>.
      when c_light2.
        assign light2 to <color>.
      when c_accent1.
        assign accent1 to <color>.
      when c_accent2.
        assign accent2 to <color>.
      when c_accent3.
        assign accent3 to <color>.
      when c_accent4.
        assign accent4 to <color>.
      when c_accent5.
        assign accent5 to <color>.
      when c_accent6.
        assign accent6 to <color>.
      when c_hlink.
        assign hlink to <color>.
      when c_folhlink.
        assign folhlink to <color>.
    endcase.
    check <color> is assigned.
    clear <color>.
    if iv_srgb is not initial.
      <color>-srgb = iv_srgb.
    else.
      <color>-syscolor-val = iv_syscolorname.
      if iv_syscolorlast is not initial.
        <color>-syscolor-lastclr = iv_syscolorlast.
      else.
        <color>-syscolor-lastclr = '000000'.
      endif.
    endif.
  endmethod.                    "set_color


method set_defaults.
    name = 'Office'.
    dark1-syscolor-val = 'windowText'.
    dark1-syscolor-lastclr = '000000'.
    light1-syscolor-val = 'window'.
    light1-syscolor-lastclr = 'FFFFFF'.
    dark2-srgb = '44546A'.
    light2-srgb = 'E7E6E6'.
    accent1-srgb = '5B9BD5'.
    accent2-srgb = 'ED7D31'.
    accent3-srgb = 'A5A5A5'.
    accent4-srgb = 'FFC000'.
    accent5-srgb = '4472C4'.
    accent6-srgb = '70AD47'.
    hlink-srgb   = '0563C1'.
    folhlink-srgb = '954F72'.
  endmethod.                    "set_defaults


method set_name.
    if strlen( iv_name ) > 50.
      name = iv_name(50).
    else.
      name = iv_name.
    endif.
  endmethod.                    "set_name
ENDCLASS.
