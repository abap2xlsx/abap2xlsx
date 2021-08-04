CLASS zcl_excel_theme_font_scheme DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF t_font,
        script   TYPE string,
        typeface TYPE string,
      END OF t_font .
    TYPES:
      tt_font TYPE SORTED TABLE OF t_font WITH UNIQUE KEY script .
    TYPES:
      BEGIN OF t_fonttype,
        typeface    TYPE string,
        panose      TYPE string,
        pitchfamily TYPE string,
        charset     TYPE string,
      END OF t_fonttype .
    TYPES:
      BEGIN OF t_fonts,
        latin TYPE t_fonttype,
        ea    TYPE t_fonttype,
        cs    TYPE t_fonttype,
        fonts TYPE tt_font,
      END OF t_fonts .
    TYPES:
      BEGIN OF t_scheme,
        name  TYPE string,
        major TYPE t_fonts,
        minor TYPE t_fonts,
      END OF t_scheme .

    CONSTANTS c_name TYPE string VALUE 'name'.              "#EC NOTEXT
    CONSTANTS c_scheme TYPE string VALUE 'fontScheme'.      "#EC NOTEXT
    CONSTANTS c_major TYPE string VALUE 'majorFont'.        "#EC NOTEXT
    CONSTANTS c_minor TYPE string VALUE 'minorFont'.        "#EC NOTEXT
    CONSTANTS c_font TYPE string VALUE 'font'.              "#EC NOTEXT
    CONSTANTS c_latin TYPE string VALUE 'latin'.            "#EC NOTEXT
    CONSTANTS c_ea TYPE string VALUE 'ea'.                  "#EC NOTEXT
    CONSTANTS c_cs TYPE string VALUE 'cs'.                  "#EC NOTEXT
    CONSTANTS c_typeface TYPE string VALUE 'typeface'.      "#EC NOTEXT
    CONSTANTS c_panose TYPE string VALUE 'panose'.          "#EC NOTEXT
    CONSTANTS c_pitchfamily TYPE string VALUE 'pitchFamily'. "#EC NOTEXT
    CONSTANTS c_charset TYPE string VALUE 'charset'.        "#EC NOTEXT
    CONSTANTS c_script TYPE string VALUE 'script'.          "#EC NOTEXT

    METHODS load
      IMPORTING
        !io_font_scheme TYPE REF TO if_ixml_element .
    METHODS set_name
      IMPORTING
        VALUE(iv_name) TYPE string .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
    METHODS modify_font
      IMPORTING
        VALUE(iv_type)     TYPE string
        VALUE(iv_script)   TYPE string
        VALUE(iv_typeface) TYPE string .
    METHODS modify_latin_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS modify_ea_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS modify_cs_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS constructor .
  PROTECTED SECTION.

    METHODS modify_lec_fonts
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_font_type)   TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
  PRIVATE SECTION.

    DATA font_scheme TYPE t_scheme .

    METHODS set_defaults .
ENDCLASS.



CLASS zcl_excel_theme_font_scheme IMPLEMENTATION.


  METHOD build_xml.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    DATA: lo_font TYPE REF TO if_ixml_element.
    DATA: lo_latin TYPE REF TO if_ixml_element.
    DATA: lo_ea TYPE REF TO if_ixml_element.
    DATA: lo_cs TYPE REF TO if_ixml_element.
    DATA: lo_major TYPE REF TO if_ixml_element.
    DATA: lo_minor TYPE REF TO if_ixml_element.
    DATA: lo_elements TYPE REF TO if_ixml_element.
    FIELD-SYMBOLS: <font> TYPE t_font.
    CHECK io_document IS BOUND.
    lo_elements ?= io_document->find_from_name_ns( name = zcl_excel_theme=>c_theme_elements ).
    IF lo_elements IS BOUND.
      lo_scheme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = zcl_excel_theme_elements=>c_font_scheme
                                                               parent = lo_elements ).
      lo_scheme_element->set_attribute( name = c_name value = font_scheme-name ).

      lo_major ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_major
                                                      parent = lo_scheme_element ).
      IF lo_major IS BOUND.
        lo_latin ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_latin
                                                         parent = lo_major ).
        lo_latin->set_attribute( name = c_typeface value = font_scheme-major-latin-typeface ).
        IF font_scheme-major-latin-panose IS NOT INITIAL.
          lo_latin->set_attribute( name = c_panose value = font_scheme-major-latin-panose ).
        ENDIF.
        IF font_scheme-major-latin-pitchfamily IS NOT INITIAL.
          lo_latin->set_attribute( name = c_pitchfamily value = font_scheme-major-latin-pitchfamily ).
        ENDIF.
        IF font_scheme-major-latin-charset IS NOT INITIAL.
          lo_latin->set_attribute( name = c_charset value = font_scheme-major-latin-charset ).
        ENDIF.

        lo_ea ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_ea
                                                         parent = lo_major ).
        lo_ea->set_attribute( name = c_typeface value = font_scheme-major-ea-typeface ).
        IF font_scheme-major-ea-panose IS NOT INITIAL.
          lo_ea->set_attribute( name = c_panose value = font_scheme-major-ea-panose ).
        ENDIF.
        IF font_scheme-major-ea-pitchfamily IS NOT INITIAL.
          lo_ea->set_attribute( name = c_pitchfamily value = font_scheme-major-ea-pitchfamily ).
        ENDIF.
        IF font_scheme-major-ea-charset IS NOT INITIAL.
          lo_ea->set_attribute( name = c_charset value = font_scheme-major-ea-charset ).
        ENDIF.

        lo_cs ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_cs
                                                      parent = lo_major ).
        lo_cs->set_attribute( name = c_typeface value = font_scheme-major-cs-typeface ).
        IF font_scheme-major-cs-panose IS NOT INITIAL.
          lo_cs->set_attribute( name = c_panose value = font_scheme-major-cs-panose ).
        ENDIF.
        IF font_scheme-major-cs-pitchfamily IS NOT INITIAL.
          lo_cs->set_attribute( name = c_pitchfamily value = font_scheme-major-cs-pitchfamily ).
        ENDIF.
        IF font_scheme-major-cs-charset IS NOT INITIAL.
          lo_cs->set_attribute( name = c_charset value = font_scheme-major-cs-charset ).
        ENDIF.

        LOOP AT font_scheme-major-fonts ASSIGNING <font>.
          IF <font>-script IS NOT INITIAL AND <font>-typeface IS NOT INITIAL.
            CLEAR lo_font.
            lo_font ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_font
                                                           parent = lo_major ).
            lo_font->set_attribute( name = c_script value = <font>-script ).
            lo_font->set_attribute( name = c_typeface value = <font>-typeface ).
          ENDIF.
        ENDLOOP.
        CLEAR: lo_latin, lo_ea, lo_cs, lo_font.
      ENDIF.

      lo_minor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_minor
                                                      parent = lo_scheme_element ).
      IF lo_minor IS BOUND.
        lo_latin ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_latin
                                                         parent = lo_minor ).
        lo_latin->set_attribute( name = c_typeface value = font_scheme-minor-latin-typeface ).
        IF font_scheme-minor-latin-panose IS NOT INITIAL.
          lo_latin->set_attribute( name = c_panose value = font_scheme-minor-latin-panose ).
        ENDIF.
        IF font_scheme-minor-latin-pitchfamily IS NOT INITIAL.
          lo_latin->set_attribute( name = c_pitchfamily value = font_scheme-minor-latin-pitchfamily ).
        ENDIF.
        IF font_scheme-minor-latin-charset IS NOT INITIAL.
          lo_latin->set_attribute( name = c_charset value = font_scheme-minor-latin-charset ).
        ENDIF.

        lo_ea ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_ea
                                                         parent = lo_minor ).
        lo_ea->set_attribute( name = c_typeface value = font_scheme-minor-ea-typeface ).
        IF font_scheme-minor-ea-panose IS NOT INITIAL.
          lo_ea->set_attribute( name = c_panose value = font_scheme-minor-ea-panose ).
        ENDIF.
        IF font_scheme-minor-ea-pitchfamily IS NOT INITIAL.
          lo_ea->set_attribute( name = c_pitchfamily value = font_scheme-minor-ea-pitchfamily ).
        ENDIF.
        IF font_scheme-minor-ea-charset IS NOT INITIAL.
          lo_ea->set_attribute( name = c_charset value = font_scheme-minor-ea-charset ).
        ENDIF.

        lo_cs ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_cs
                                                      parent = lo_minor ).
        lo_cs->set_attribute( name = c_typeface value = font_scheme-minor-cs-typeface ).
        IF font_scheme-minor-cs-panose IS NOT INITIAL.
          lo_cs->set_attribute( name = c_panose value = font_scheme-minor-cs-panose ).
        ENDIF.
        IF font_scheme-minor-cs-pitchfamily IS NOT INITIAL.
          lo_cs->set_attribute( name = c_pitchfamily value = font_scheme-minor-cs-pitchfamily ).
        ENDIF.
        IF font_scheme-minor-cs-charset IS NOT INITIAL.
          lo_cs->set_attribute( name = c_charset value = font_scheme-minor-cs-charset ).
        ENDIF.

        LOOP AT font_scheme-minor-fonts ASSIGNING <font>.
          IF <font>-script IS NOT INITIAL AND <font>-typeface IS NOT INITIAL.
            CLEAR lo_font.
            lo_font ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_font
                                                           parent = lo_minor ).
            lo_font->set_attribute( name = c_script value = <font>-script ).
            lo_font->set_attribute( name = c_typeface value = <font>-typeface ).
          ENDIF.
        ENDLOOP.
      ENDIF.


    ENDIF.
  ENDMETHOD.                    "build_xml


  METHOD constructor.
    set_defaults( ).
  ENDMETHOD.                    "constructor


  METHOD load.
    DATA: lo_scheme_children TYPE REF TO if_ixml_node_list.
    DATA: lo_scheme_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    DATA: lo_major_children TYPE REF TO if_ixml_node_list.
    DATA: lo_major_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_major_element TYPE REF TO if_ixml_element.
    DATA: lo_minor_children TYPE REF TO if_ixml_node_list.
    DATA: lo_minor_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_minor_element TYPE REF TO if_ixml_element.
    DATA: ls_font TYPE t_font.
    CHECK io_font_scheme IS NOT INITIAL.
    CLEAR font_scheme.
    font_scheme-name =  io_font_scheme->get_attribute( name = c_name ).
    lo_scheme_children = io_font_scheme->get_children( ).
    lo_scheme_iterator = lo_scheme_children->create_iterator( ).
    lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    WHILE lo_scheme_element IS BOUND.
      CASE lo_scheme_element->get_name( ).
        WHEN c_major.
          lo_major_children = lo_scheme_element->get_children( ).
          lo_major_iterator = lo_major_children->create_iterator( ).
          lo_major_element ?= lo_major_iterator->get_next( ).
          WHILE lo_major_element IS BOUND.
            CASE lo_major_element->get_name( ).
              WHEN c_latin.
                font_scheme-major-latin-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                font_scheme-major-latin-panose = lo_major_element->get_attribute(  name = c_panose ).
                font_scheme-major-latin-pitchfamily = lo_major_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-major-latin-charset = lo_major_element->get_attribute(  name = c_charset ).
              WHEN c_ea.
                font_scheme-major-ea-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                font_scheme-major-ea-panose = lo_major_element->get_attribute(  name = c_panose ).
                font_scheme-major-ea-pitchfamily = lo_major_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-major-ea-charset = lo_major_element->get_attribute(  name = c_charset ).
              WHEN c_cs.
                font_scheme-major-cs-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                font_scheme-major-cs-panose = lo_major_element->get_attribute(  name = c_panose ).
                font_scheme-major-cs-pitchfamily = lo_major_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-major-cs-charset = lo_major_element->get_attribute(  name = c_charset ).
              WHEN c_font.
                CLEAR ls_font.
                ls_font-script = lo_major_element->get_attribute(  name = c_script ).
                ls_font-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                TRY.
                    INSERT ls_font INTO TABLE font_scheme-major-fonts.
                  CATCH cx_root. "not the best but just to avoid duplicate lines dump

                ENDTRY.
            ENDCASE.
            lo_major_element ?= lo_major_iterator->get_next( ).
          ENDWHILE.
        WHEN c_minor.
          lo_minor_children = lo_scheme_element->get_children( ).
          lo_minor_iterator = lo_minor_children->create_iterator( ).
          lo_minor_element ?= lo_minor_iterator->get_next( ).
          WHILE lo_minor_element IS BOUND.
            CASE lo_minor_element->get_name( ).
              WHEN c_latin.
                font_scheme-minor-latin-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                font_scheme-minor-latin-panose = lo_minor_element->get_attribute(  name = c_panose ).
                font_scheme-minor-latin-pitchfamily = lo_minor_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-minor-latin-charset = lo_minor_element->get_attribute(  name = c_charset ).
              WHEN c_ea.
                font_scheme-minor-ea-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                font_scheme-minor-ea-panose = lo_minor_element->get_attribute(  name = c_panose ).
                font_scheme-minor-ea-pitchfamily = lo_minor_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-minor-ea-charset = lo_minor_element->get_attribute(  name = c_charset ).
              WHEN c_cs.
                font_scheme-minor-cs-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                font_scheme-minor-cs-panose = lo_minor_element->get_attribute(  name = c_panose ).
                font_scheme-minor-cs-pitchfamily = lo_minor_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-minor-cs-charset = lo_minor_element->get_attribute(  name = c_charset ).
              WHEN c_font.
                CLEAR ls_font.
                ls_font-script = lo_minor_element->get_attribute(  name = c_script ).
                ls_font-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                TRY.
                    INSERT ls_font INTO TABLE font_scheme-minor-fonts.
                  CATCH cx_root. "not the best but just to avoid duplicate lines dump

                ENDTRY.
            ENDCASE.
            lo_minor_element ?= lo_minor_iterator->get_next( ).
          ENDWHILE.
      ENDCASE.
      lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    ENDWHILE.
  ENDMETHOD.                    "load


  METHOD modify_cs_font.
    modify_lec_fonts(
      EXPORTING
        iv_type        = iv_type
        iv_font_type   = c_cs
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "modify_latin_font


  METHOD modify_ea_font.
    modify_lec_fonts(
      EXPORTING
        iv_type        = iv_type
        iv_font_type   = c_ea
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "modify_latin_font


  METHOD modify_font.
    DATA: ls_font TYPE t_font.
    FIELD-SYMBOLS: <font> TYPE t_font.
    ls_font-script = iv_script.
    ls_font-typeface = iv_typeface.
    TRY.
        CASE iv_type.
          WHEN c_major.
            READ TABLE font_scheme-major-fonts WITH KEY script = iv_script ASSIGNING <font>.
            IF sy-subrc EQ 0.
              <font> = ls_font.
            ELSE.
              INSERT ls_font INTO TABLE font_scheme-major-fonts.
            ENDIF.
          WHEN c_minor.
            READ TABLE font_scheme-minor-fonts WITH KEY script = iv_script ASSIGNING <font>.
            IF sy-subrc EQ 0.
              <font> = ls_font.
            ELSE.
              INSERT ls_font INTO TABLE font_scheme-minor-fonts.
            ENDIF.
        ENDCASE.
      CATCH cx_root. "not the best but just to avoid duplicate lines dump
    ENDTRY.
  ENDMETHOD.                    "add_font


  METHOD modify_latin_font.
    modify_lec_fonts(
      EXPORTING
        iv_type        = iv_type
        iv_font_type   = c_latin
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "modify_latin_font


  METHOD modify_lec_fonts.
    FIELD-SYMBOLS: <type> TYPE t_fonts,
                   <font> TYPE t_fonttype.
    CASE iv_type.
      WHEN c_minor.
        ASSIGN font_scheme-minor TO <type>.
      WHEN c_major.
        ASSIGN font_scheme-major TO <type>.
      WHEN OTHERS.
        RETURN.
    ENDCASE.
    CHECK <type> IS ASSIGNED.
    CASE iv_font_type.
      WHEN c_latin.
        ASSIGN <type>-latin TO <font>.
      WHEN c_ea.
        ASSIGN <type>-ea TO <font>.
      WHEN c_cs.
        ASSIGN <type>-cs TO <font>.
      WHEN OTHERS.
        RETURN.
    ENDCASE.
    CHECK <font> IS ASSIGNED.
    <font>-typeface = iv_typeface.
    <font>-panose = iv_panose.
    <font>-pitchfamily = iv_pitchfamily.
    <font>-charset = iv_charset.
  ENDMETHOD.                    "modify_lec_fonts


  METHOD set_defaults.
    CLEAR font_scheme.
    font_scheme-name = 'Office'.
    font_scheme-major-latin-typeface = 'Calibri Light'.
    font_scheme-major-latin-panose = '020F0302020204030204'.
    modify_font( iv_type = c_major iv_script = 'Jpan' iv_typeface = 'ＭＳ Ｐゴシック' ).
    modify_font( iv_type = c_major iv_script = 'Hang' iv_typeface = '맑은 고딕' ).
    modify_font( iv_type = c_major iv_script = 'Hans' iv_typeface = '宋体' ).
    modify_font( iv_type = c_major iv_script = 'Hant' iv_typeface = '新細明體' ).
    modify_font( iv_type = c_major iv_script = 'Arab' iv_typeface = 'Times New Roman' ).
    modify_font( iv_type = c_major iv_script = 'Hebr' iv_typeface = 'Times New Roman' ).
    modify_font( iv_type = c_major iv_script = 'Thai' iv_typeface = 'Tahoma' ).
    modify_font( iv_type = c_major iv_script = 'Ethi' iv_typeface = 'Nyala' ).
    modify_font( iv_type = c_major iv_script = 'Beng' iv_typeface = 'Vrinda' ).
    modify_font( iv_type = c_major iv_script = 'Gujr' iv_typeface = 'Shruti' ).
    modify_font( iv_type = c_major iv_script = 'Khmr' iv_typeface = 'MoolBoran' ).
    modify_font( iv_type = c_major iv_script = 'Knda' iv_typeface = 'Tunga' ).
    modify_font( iv_type = c_major iv_script = 'Guru' iv_typeface = 'Raavi' ).
    modify_font( iv_type = c_major iv_script = 'Cans' iv_typeface = 'Euphemia' ).
    modify_font( iv_type = c_major iv_script = 'Cher' iv_typeface = 'Plantagenet Cherokee' ).
    modify_font( iv_type = c_major iv_script = 'Yiii' iv_typeface = 'Microsoft Yi Baiti' ).
    modify_font( iv_type = c_major iv_script = 'Tibt' iv_typeface = 'Microsoft Himalaya' ).
    modify_font( iv_type = c_major iv_script = 'Thaa' iv_typeface = 'MV Boli' ).
    modify_font( iv_type = c_major iv_script = 'Deva' iv_typeface = 'Mangal' ).
    modify_font( iv_type = c_major iv_script = 'Telu' iv_typeface = 'Gautami' ).
    modify_font( iv_type = c_major iv_script = 'Taml' iv_typeface = 'Latha' ).
    modify_font( iv_type = c_major iv_script = 'Syrc' iv_typeface = 'Estrangelo Edessa' ).
    modify_font( iv_type = c_major iv_script = 'Orya' iv_typeface = 'Kalinga' ).
    modify_font( iv_type = c_major iv_script = 'Mlym' iv_typeface = 'Kartika' ).
    modify_font( iv_type = c_major iv_script = 'Laoo' iv_typeface = 'DokChampa' ).
    modify_font( iv_type = c_major iv_script = 'Sinh' iv_typeface = 'Iskoola Pota' ).
    modify_font( iv_type = c_major iv_script = 'Mong' iv_typeface = 'Mongolian Baiti' ).
    modify_font( iv_type = c_major iv_script = 'Viet' iv_typeface = 'Times New Roman' ).
    modify_font( iv_type = c_major iv_script = 'Uigh' iv_typeface = 'Microsoft Uighur' ).
    modify_font( iv_type = c_major iv_script = 'Geor' iv_typeface = 'Sylfaen' ).

    font_scheme-minor-latin-typeface = 'Calibri'.
    font_scheme-minor-latin-panose = '020F0502020204030204'.
    modify_font( iv_type = c_minor iv_script = 'Jpan' iv_typeface = 'ＭＳ Ｐゴシック' ).
    modify_font( iv_type = c_minor iv_script = 'Hang' iv_typeface = '맑은 고딕' ).
    modify_font( iv_type = c_minor iv_script = 'Hans' iv_typeface = '宋体' ).
    modify_font( iv_type = c_minor iv_script = 'Hant' iv_typeface = '新細明體' ).
    modify_font( iv_type = c_minor iv_script = 'Arab' iv_typeface = 'Arial' ).
    modify_font( iv_type = c_minor iv_script = 'Hebr' iv_typeface = 'Arial' ).
    modify_font( iv_type = c_minor iv_script = 'Thai' iv_typeface = 'Tahoma' ).
    modify_font( iv_type = c_minor iv_script = 'Ethi' iv_typeface = 'Nyala' ).
    modify_font( iv_type = c_minor iv_script = 'Beng' iv_typeface = 'Vrinda' ).
    modify_font( iv_type = c_minor iv_script = 'Gujr' iv_typeface = 'Shruti' ).
    modify_font( iv_type = c_minor iv_script = 'Khmr' iv_typeface = 'DaunPenh' ).
    modify_font( iv_type = c_minor iv_script = 'Knda' iv_typeface = 'Tunga' ).
    modify_font( iv_type = c_minor iv_script = 'Guru' iv_typeface = 'Raavi' ).
    modify_font( iv_type = c_minor iv_script = 'Cans' iv_typeface = 'Euphemia' ).
    modify_font( iv_type = c_minor iv_script = 'Cher' iv_typeface = 'Plantagenet Cherokee' ).
    modify_font( iv_type = c_minor iv_script = 'Yiii' iv_typeface = 'Microsoft Yi Baiti' ).
    modify_font( iv_type = c_minor iv_script = 'Tibt' iv_typeface = 'Microsoft Himalaya' ).
    modify_font( iv_type = c_minor iv_script = 'Thaa' iv_typeface = 'MV Boli' ).
    modify_font( iv_type = c_minor iv_script = 'Deva' iv_typeface = 'Mangal' ).
    modify_font( iv_type = c_minor iv_script = 'Telu' iv_typeface = 'Gautami' ).
    modify_font( iv_type = c_minor iv_script = 'Taml' iv_typeface = 'Latha' ).
    modify_font( iv_type = c_minor iv_script = 'Syrc' iv_typeface = 'Estrangelo Edessa' ).
    modify_font( iv_type = c_minor iv_script = 'Orya' iv_typeface = 'Kalinga' ).
    modify_font( iv_type = c_minor iv_script = 'Mlym' iv_typeface = 'Kartika' ).
    modify_font( iv_type = c_minor iv_script = 'Laoo' iv_typeface = 'DokChampa' ).
    modify_font( iv_type = c_minor iv_script = 'Sinh' iv_typeface = 'Iskoola Pota' ).
    modify_font( iv_type = c_minor iv_script = 'Mong' iv_typeface = 'Mongolian Baiti' ).
    modify_font( iv_type = c_minor iv_script = 'Viet' iv_typeface = 'Arial' ).
    modify_font( iv_type = c_minor iv_script = 'Uigh' iv_typeface = 'Microsoft Uighur' ).
    modify_font( iv_type = c_minor iv_script = 'Geor' iv_typeface = 'Sylfaen' ).

  ENDMETHOD.                    "set_defaults


  METHOD set_name.
    font_scheme-name = iv_name.
  ENDMETHOD.                    "set_name
ENDCLASS.
