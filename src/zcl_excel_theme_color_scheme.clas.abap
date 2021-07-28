CLASS zcl_excel_theme_color_scheme DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC

  GLOBAL FRIENDS zcl_excel_theme
                 zcl_excel_theme_elements .

  PUBLIC SECTION.

    TYPES t_srgb TYPE string .
    TYPES:
      BEGIN OF t_syscolor,
        val     TYPE string,
        lastclr TYPE t_srgb,
      END OF t_syscolor .
    TYPES:
      BEGIN OF t_color,
        srgb     TYPE t_srgb,
        syscolor TYPE t_syscolor,
      END OF t_color .

    CONSTANTS c_dark1 TYPE string VALUE 'dk1'.              "#EC NOTEXT
    CONSTANTS c_dark2 TYPE string VALUE 'dk2'.              "#EC NOTEXT
    CONSTANTS c_light1 TYPE string VALUE 'lt1'.             "#EC NOTEXT
    CONSTANTS c_light2 TYPE string VALUE 'lt2'.             "#EC NOTEXT
    CONSTANTS c_accent1 TYPE string VALUE 'accent1'.        "#EC NOTEXT
    CONSTANTS c_accent2 TYPE string VALUE 'accent2'.        "#EC NOTEXT
    CONSTANTS c_accent3 TYPE string VALUE 'accent3'.        "#EC NOTEXT
    CONSTANTS c_accent4 TYPE string VALUE 'accent4'.        "#EC NOTEXT
    CONSTANTS c_accent5 TYPE string VALUE 'accent5'.        "#EC NOTEXT
    CONSTANTS c_accent6 TYPE string VALUE 'accent6'.        "#EC NOTEXT
    CONSTANTS c_hlink TYPE string VALUE 'hlink'.            "#EC NOTEXT
    CONSTANTS c_folhlink TYPE string VALUE 'folHlink'.      "#EC NOTEXT
    CONSTANTS c_syscolor TYPE string VALUE 'sysClr'.        "#EC NOTEXT
    CONSTANTS c_srgbcolor TYPE string VALUE 'srgbClr'.      "#EC NOTEXT
    CONSTANTS c_val TYPE string VALUE 'val'.                "#EC NOTEXT
    CONSTANTS c_lastclr TYPE string VALUE 'lastClr'.        "#EC NOTEXT
    CONSTANTS c_name TYPE string VALUE 'name'.              "#EC NOTEXT
    CONSTANTS c_scheme TYPE string VALUE 'clrScheme'.       "#EC NOTEXT

    METHODS load
      IMPORTING
        !io_color_scheme TYPE REF TO if_ixml_element .
    METHODS set_color
      IMPORTING
        VALUE(iv_type)         TYPE string
        VALUE(iv_srgb)         TYPE t_srgb OPTIONAL
        VALUE(iv_syscolorname) TYPE string OPTIONAL
        VALUE(iv_syscolorlast) TYPE t_srgb .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
    METHODS constructor .
    METHODS set_name
      IMPORTING
        VALUE(iv_name) TYPE string .
  PROTECTED SECTION.

    DATA name TYPE string .
    DATA dark1 TYPE t_color .
    DATA dark2 TYPE t_color .
    DATA light1 TYPE t_color .
    DATA light2 TYPE t_color .
    DATA accent1 TYPE t_color .
    DATA accent2 TYPE t_color .
    DATA accent3 TYPE t_color .
    DATA accent4 TYPE t_color .
    DATA accent5 TYPE t_color .
    DATA accent6 TYPE t_color .
    DATA hlink TYPE t_color .
    DATA folhlink TYPE t_color .
  PRIVATE SECTION.

    METHODS get_color
      IMPORTING
        !io_object      TYPE REF TO if_ixml_element
      RETURNING
        VALUE(rv_color) TYPE t_color .
    METHODS set_defaults .
ENDCLASS.



CLASS zcl_excel_theme_color_scheme IMPLEMENTATION.


  METHOD build_xml.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    DATA: lo_color TYPE REF TO if_ixml_element.
    DATA: lo_syscolor TYPE REF TO if_ixml_element.
    DATA: lo_srgb TYPE REF TO if_ixml_element.
    DATA: lo_elements TYPE REF TO if_ixml_element.

    CHECK io_document IS BOUND.
    lo_elements ?= io_document->find_from_name_ns( name   = zcl_excel_theme=>c_theme_elements ).
    IF lo_elements IS BOUND.
      lo_scheme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                  name   = zcl_excel_theme_elements=>c_color_scheme
                                                               parent = lo_elements ).
      lo_scheme_element->set_attribute( name = c_name value = name ).

      "! Adding colors to scheme
      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                  name   = c_dark1
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF dark1-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = dark1-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = dark1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = dark1-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_light1
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF light1-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = light1-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = light1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = light1-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_dark2
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF dark2-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = dark2-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = dark2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = dark2-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_light2
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF light2-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = light2-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = light2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = light2-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent1
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent1-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent1-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent1-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent2
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent2-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent2-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent2-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent3
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent3-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent3-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent3-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent3-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent4
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent4-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent4-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent4-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent4-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent5
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent5-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent5-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent5-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent5-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent6
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent6-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent6-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent6-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent6-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_hlink
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF hlink-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = hlink-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = hlink-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = hlink-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_folhlink
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF folhlink-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = folhlink-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = folhlink-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = folhlink-syscolor-lastclr ).
        ENDIF.
      ENDIF.


    ENDIF.
  ENDMETHOD.                    "build_xml


  METHOD constructor.
    set_defaults( ).
  ENDMETHOD.                    "constructor


  METHOD get_color.
    DATA: lo_color_children TYPE REF TO if_ixml_node_list.
    DATA: lo_color_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_color_element TYPE REF TO if_ixml_element.
    CHECK io_object  IS NOT INITIAL.

    lo_color_children = io_object->get_children( ).
    lo_color_iterator = lo_color_children->create_iterator( ).
    lo_color_element ?= lo_color_iterator->get_next( ).
    IF lo_color_element IS BOUND.
      CASE lo_color_element->get_name( ).
        WHEN c_srgbcolor.
          rv_color-srgb = lo_color_element->get_attribute( name = c_val ).
        WHEN c_syscolor.
          rv_color-syscolor-val = lo_color_element->get_attribute( name = c_val ).
          rv_color-syscolor-lastclr = lo_color_element->get_attribute( name = c_lastclr ).
      ENDCASE.
    ENDIF.
  ENDMETHOD.                    "get_color


  METHOD load.
    DATA: lo_scheme_children TYPE REF TO if_ixml_node_list.
    DATA: lo_scheme_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    CHECK io_color_scheme  IS NOT INITIAL.

    name = io_color_scheme->get_attribute( name = c_name ).
    lo_scheme_children = io_color_scheme->get_children( ).
    lo_scheme_iterator = lo_scheme_children->create_iterator( ).
    lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    WHILE lo_scheme_element IS BOUND.
      CASE lo_scheme_element->get_name( ).
        WHEN c_dark1.
          dark1 = me->get_color( lo_scheme_element ).
        WHEN c_dark2.
          dark2 = me->get_color( lo_scheme_element ).
        WHEN c_light1.
          light1 = me->get_color( lo_scheme_element ).
        WHEN c_light2.
          light2 = me->get_color( lo_scheme_element ).
        WHEN c_accent1.
          accent1 = me->get_color( lo_scheme_element ).
        WHEN c_accent2.
          accent2 = me->get_color( lo_scheme_element ).
        WHEN c_accent3.
          accent3 = me->get_color( lo_scheme_element ).
        WHEN c_accent4.
          accent4 = me->get_color( lo_scheme_element ).
        WHEN c_accent5.
          accent5 = me->get_color( lo_scheme_element ).
        WHEN c_accent6.
          accent6 = me->get_color( lo_scheme_element ).
        WHEN c_hlink.
          hlink = me->get_color( lo_scheme_element ).
        WHEN c_folhlink.
          folhlink = me->get_color( lo_scheme_element ).
      ENDCASE.
      lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    ENDWHILE.
  ENDMETHOD.                    "load


  METHOD set_color.
    FIELD-SYMBOLS: <color> TYPE t_color.
    CHECK iv_type IS NOT INITIAL.
    CHECK iv_srgb IS NOT INITIAL OR  iv_syscolorname IS NOT INITIAL.
    CASE iv_type.
      WHEN c_dark1.
        ASSIGN dark1 TO <color>.
      WHEN c_dark2.
        ASSIGN dark2 TO <color>.
      WHEN c_light1.
        ASSIGN light1 TO <color>.
      WHEN c_light2.
        ASSIGN light2 TO <color>.
      WHEN c_accent1.
        ASSIGN accent1 TO <color>.
      WHEN c_accent2.
        ASSIGN accent2 TO <color>.
      WHEN c_accent3.
        ASSIGN accent3 TO <color>.
      WHEN c_accent4.
        ASSIGN accent4 TO <color>.
      WHEN c_accent5.
        ASSIGN accent5 TO <color>.
      WHEN c_accent6.
        ASSIGN accent6 TO <color>.
      WHEN c_hlink.
        ASSIGN hlink TO <color>.
      WHEN c_folhlink.
        ASSIGN folhlink TO <color>.
    ENDCASE.
    CHECK <color> IS ASSIGNED.
    CLEAR <color>.
    IF iv_srgb IS NOT INITIAL.
      <color>-srgb = iv_srgb.
    ELSE.
      <color>-syscolor-val = iv_syscolorname.
      IF iv_syscolorlast IS NOT INITIAL.
        <color>-syscolor-lastclr = iv_syscolorlast.
      ELSE.
        <color>-syscolor-lastclr = '000000'.
      ENDIF.
    ENDIF.
  ENDMETHOD.                    "set_color


  METHOD set_defaults.
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
  ENDMETHOD.                    "set_defaults


  METHOD set_name.
    IF strlen( iv_name ) > 50.
      name = iv_name(50).
    ELSE.
      name = iv_name.
    ENDIF.
  ENDMETHOD.                    "set_name
ENDCLASS.
