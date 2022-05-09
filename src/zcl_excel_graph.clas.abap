CLASS zcl_excel_graph DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
*"* public components of class ZCL_EXCEL_GRAPH
*"* do not include other source files here!!!
      BEGIN OF s_style,
        c14style TYPE i,
        cstyle   TYPE i,
      END OF s_style .
    TYPES:
      BEGIN OF s_series,
        idx              TYPE i,
        order            TYPE i,
        invertifnegative TYPE string,
        symbol           TYPE string,
        smooth           TYPE string,
        lbl              TYPE string,
        ref              TYPE string,
        sername          TYPE string,
      END OF s_series .
    TYPES:
      t_series TYPE STANDARD TABLE OF s_series .
    TYPES:
      BEGIN OF s_pagemargins,
        b      TYPE string,
        l      TYPE string,
        r      TYPE string,
        t      TYPE string,
        header TYPE string,
        footer TYPE string,
      END OF s_pagemargins .

    DATA ns_1904val TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_langval TYPE string VALUE 'it-IT'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_roundedcornersval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA pagemargins TYPE s_pagemargins .
    DATA ns_autotitledeletedval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_plotvisonlyval TYPE string VALUE '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_dispblanksasval TYPE string VALUE 'gap'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showdlblsovermaxval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA title TYPE string .               .  . " .
    DATA series TYPE t_series .
    DATA ns_c14styleval TYPE string VALUE '102'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA print_label TYPE c VALUE 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_styleval TYPE string VALUE '2'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    CONSTANTS:
      BEGIN OF c_style_default,
        c14style TYPE i VALUE 102,
        cstyle   TYPE i VALUE 2,
      END OF c_style_default .
    CONSTANTS:
      BEGIN OF c_style_1,
        c14style TYPE i VALUE 101,
        cstyle   TYPE i VALUE 1,
      END OF c_style_1 .
    CONSTANTS:
      BEGIN OF c_style_3,
        c14style TYPE i VALUE 103,
        cstyle   TYPE i VALUE 3,
      END OF c_style_3 .
    CONSTANTS:
      BEGIN OF c_style_4,
        c14style TYPE i VALUE 104,
        cstyle   TYPE i VALUE 4,
      END OF c_style_4 .
    CONSTANTS:
      BEGIN OF c_style_5,
        c14style TYPE i VALUE 105,
        cstyle   TYPE i VALUE 5,
      END OF c_style_5 .
    CONSTANTS:
      BEGIN OF c_style_6,
        c14style TYPE i VALUE 106,
        cstyle   TYPE i VALUE 6,
      END OF c_style_6 .
    CONSTANTS:
      BEGIN OF c_style_7,
        c14style TYPE i VALUE 107,
        cstyle   TYPE i VALUE 7,
      END OF c_style_7 .
    CONSTANTS:
      BEGIN OF c_style_8,
        c14style TYPE i VALUE 108,
        cstyle   TYPE i VALUE 8,
      END OF c_style_8 .
    CONSTANTS:
      BEGIN OF c_style_9,
        c14style TYPE i VALUE 109,
        cstyle   TYPE i VALUE 9,
      END OF c_style_9 .
    CONSTANTS:
      BEGIN OF c_style_10,
        c14style TYPE i VALUE 110,
        cstyle   TYPE i VALUE 10,
      END OF c_style_10 .
    CONSTANTS:
      BEGIN OF c_style_11,
        c14style TYPE i VALUE 111,
        cstyle   TYPE i VALUE 11,
      END OF c_style_11 .
    CONSTANTS:
      BEGIN OF c_style_12,
        c14style TYPE i VALUE 112,
        cstyle   TYPE i VALUE 12,
      END OF c_style_12 .
    CONSTANTS:
      BEGIN OF c_style_13,
        c14style TYPE i VALUE 113,
        cstyle   TYPE i VALUE 13,
      END OF c_style_13 .
    CONSTANTS:
      BEGIN OF c_style_14,
        c14style TYPE i VALUE 114,
        cstyle   TYPE i VALUE 14,
      END OF c_style_14 .
    CONSTANTS:
      BEGIN OF c_style_15,
        c14style TYPE i VALUE 115,
        cstyle   TYPE i VALUE 15,
      END OF c_style_15 .
    CONSTANTS:
      BEGIN OF c_style_16,
        c14style TYPE i VALUE 116,
        cstyle   TYPE i VALUE 16,
      END OF c_style_16 .
    CONSTANTS:
      BEGIN OF c_style_17,
        c14style TYPE i VALUE 117,
        cstyle   TYPE i VALUE 17,
      END OF c_style_17 .
    CONSTANTS:
      BEGIN OF c_style_18,
        c14style TYPE i VALUE 118,
        cstyle   TYPE i VALUE 18,
      END OF c_style_18 .
    CONSTANTS:
      BEGIN OF c_style_19,
        c14style TYPE i VALUE 119,
        cstyle   TYPE i VALUE 19,
      END OF c_style_19 .
    CONSTANTS:
      BEGIN OF c_style_20,
        c14style TYPE i VALUE 120,
        cstyle   TYPE i VALUE 20,
      END OF c_style_20 .
    CONSTANTS:
      BEGIN OF c_style_21,
        c14style TYPE i VALUE 121,
        cstyle   TYPE i VALUE 21,
      END OF c_style_21 .
    CONSTANTS:
      BEGIN OF c_style_22,
        c14style TYPE i VALUE 122,
        cstyle   TYPE i VALUE 22,
      END OF c_style_22 .
    CONSTANTS:
      BEGIN OF c_style_23,
        c14style TYPE i VALUE 123,
        cstyle   TYPE i VALUE 23,
      END OF c_style_23 .
    CONSTANTS:
      BEGIN OF c_style_24,
        c14style TYPE i VALUE 124,
        cstyle   TYPE i VALUE 24,
      END OF c_style_24 .
    CONSTANTS:
      BEGIN OF c_style_25,
        c14style TYPE i VALUE 125,
        cstyle   TYPE i VALUE 25,
      END OF c_style_25 .
    CONSTANTS:
      BEGIN OF c_style_26,
        c14style TYPE i VALUE 126,
        cstyle   TYPE i VALUE 26,
      END OF c_style_26 .
    CONSTANTS:
      BEGIN OF c_style_27,
        c14style TYPE i VALUE 127,
        cstyle   TYPE i VALUE 27,
      END OF c_style_27 .
    CONSTANTS:
      BEGIN OF c_style_28,
        c14style TYPE i VALUE 128,
        cstyle   TYPE i VALUE 28,
      END OF c_style_28 .
    CONSTANTS:
      BEGIN OF c_style_29,
        c14style TYPE i VALUE 129,
        cstyle   TYPE i VALUE 29,
      END OF c_style_29 .
    CONSTANTS:
      BEGIN OF c_style_30,
        c14style TYPE i VALUE 130,
        cstyle   TYPE i VALUE 30,
      END OF c_style_30 .
    CONSTANTS:
      BEGIN OF c_style_31,
        c14style TYPE i VALUE 131,
        cstyle   TYPE i VALUE 31,
      END OF c_style_31 .
    CONSTANTS:
      BEGIN OF c_style_32,
        c14style TYPE i VALUE 132,
        cstyle   TYPE i VALUE 32,
      END OF c_style_32 .
    CONSTANTS:
      BEGIN OF c_style_33,
        c14style TYPE i VALUE 133,
        cstyle   TYPE i VALUE 33,
      END OF c_style_33 .
    CONSTANTS:
      BEGIN OF c_style_34,
        c14style TYPE i VALUE 134,
        cstyle   TYPE i VALUE 34,
      END OF c_style_34 .
    CONSTANTS:
      BEGIN OF c_style_35,
        c14style TYPE i VALUE 135,
        cstyle   TYPE i VALUE 35,
      END OF c_style_35 .
    CONSTANTS:
      BEGIN OF c_style_36,
        c14style TYPE i VALUE 136,
        cstyle   TYPE i VALUE 36,
      END OF c_style_36 .
    CONSTANTS:
      BEGIN OF c_style_37,
        c14style TYPE i VALUE 137,
        cstyle   TYPE i VALUE 37,
      END OF c_style_37 .
    CONSTANTS:
      BEGIN OF c_style_38,
        c14style TYPE i VALUE 138,
        cstyle   TYPE i VALUE 38,
      END OF c_style_38 .
    CONSTANTS:
      BEGIN OF c_style_39,
        c14style TYPE i VALUE 139,
        cstyle   TYPE i VALUE 39,
      END OF c_style_39 .
    CONSTANTS:
      BEGIN OF c_style_40,
        c14style TYPE i VALUE 140,
        cstyle   TYPE i VALUE 40,
      END OF c_style_40 .
    CONSTANTS:
      BEGIN OF c_style_41,
        c14style TYPE i VALUE 141,
        cstyle   TYPE i VALUE 41,
      END OF c_style_41 .
    CONSTANTS:
      BEGIN OF c_style_42,
        c14style TYPE i VALUE 142,
        cstyle   TYPE i VALUE 42,
      END OF c_style_42 .
    CONSTANTS:
      BEGIN OF c_style_43,
        c14style TYPE i VALUE 143,
        cstyle   TYPE i VALUE 43,
      END OF c_style_43 .
    CONSTANTS:
      BEGIN OF c_style_44,
        c14style TYPE i VALUE 144,
        cstyle   TYPE i VALUE 44,
      END OF c_style_44 .
    CONSTANTS:
      BEGIN OF c_style_45,
        c14style TYPE i VALUE 145,
        cstyle   TYPE i VALUE 45,
      END OF c_style_45 .
    CONSTANTS:
      BEGIN OF c_style_46,
        c14style TYPE i VALUE 146,
        cstyle   TYPE i VALUE 46,
      END OF c_style_46 .
    CONSTANTS:
      BEGIN OF c_style_47,
        c14style TYPE i VALUE 147,
        cstyle   TYPE i VALUE 47,
      END OF c_style_47 .
    CONSTANTS:
      BEGIN OF c_style_48,
        c14style TYPE i VALUE 148,
        cstyle   TYPE i VALUE 48,
      END OF c_style_48 .
    CONSTANTS c_show_true TYPE c VALUE '1'.                 "#EC NOTEXT
    CONSTANTS c_show_false TYPE c VALUE '0'.                "#EC NOTEXT
    CONSTANTS c_print_lbl_true TYPE c VALUE '1'.            "#EC NOTEXT
    CONSTANTS c_print_lbl_false TYPE c VALUE '0'.           "#EC NOTEXT

    METHODS constructor .
    METHODS create_serie
      IMPORTING
        !ip_idx              TYPE i OPTIONAL
        !ip_order            TYPE i
        !ip_invertifnegative TYPE string OPTIONAL
        !ip_symbol           TYPE string OPTIONAL
        !ip_smooth           TYPE c OPTIONAL
        !ip_lbl_from_col     TYPE zexcel_cell_column_alpha OPTIONAL
        !ip_lbl_from_row     TYPE zexcel_cell_row OPTIONAL
        !ip_lbl_to_col       TYPE zexcel_cell_column_alpha OPTIONAL
        !ip_lbl_to_row       TYPE zexcel_cell_row OPTIONAL
        !ip_lbl              TYPE string OPTIONAL
        !ip_ref_from_col     TYPE zexcel_cell_column_alpha OPTIONAL
        !ip_ref_from_row     TYPE zexcel_cell_row OPTIONAL
        !ip_ref_to_col       TYPE zexcel_cell_column_alpha OPTIONAL
        !ip_ref_to_row       TYPE zexcel_cell_row OPTIONAL
        !ip_ref              TYPE string OPTIONAL
        !ip_sername          TYPE string
        !ip_sheet            TYPE zexcel_sheet_title OPTIONAL .
    METHODS set_style
      IMPORTING
        !ip_style TYPE s_style .
    METHODS set_print_lbl
      IMPORTING
        !ip_value TYPE c .
    METHODS set_title
      IMPORTING
        ip_value TYPE string .
  PROTECTED SECTION.
*"* protected components of class ZCL_EXCEL_GRAPH
*"* do not include other source files here!!!
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_GRAPH
*"* do not include other source files here!!!
ENDCLASS.



CLASS zcl_excel_graph IMPLEMENTATION.


  METHOD constructor.
    "Load default values
    me->pagemargins-b = '0.75'.
    me->pagemargins-l = '0.7'.
    me->pagemargins-r = '0.7'.
    me->pagemargins-t = '0.75'.
    me->pagemargins-header = '0.3'.
    me->pagemargins-footer = '0.3'.
  ENDMETHOD.


  METHOD create_serie.
    DATA ls_serie TYPE s_series.

    DATA: lv_start_row_c TYPE c LENGTH 7,
          lv_stop_row_c  TYPE c LENGTH 7.


    IF ip_lbl IS NOT SUPPLIED.
      lv_stop_row_c = ip_lbl_to_row.
      SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_stop_row_c LEFT DELETING LEADING space.
      lv_start_row_c = ip_lbl_from_row.
      SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_start_row_c LEFT DELETING LEADING space.
      ls_serie-lbl = ip_sheet.
      ls_serie-lbl = zcl_excel_common=>escape_string( ip_value = ls_serie-lbl ).
      CONCATENATE ls_serie-lbl '!$' ip_lbl_from_col '$' lv_start_row_c ':$' ip_lbl_to_col '$' lv_stop_row_c INTO ls_serie-lbl.
      CLEAR: lv_start_row_c, lv_stop_row_c.
    ELSE.
      ls_serie-lbl = ip_lbl.
    ENDIF.
    IF ip_ref IS NOT SUPPLIED.
      lv_stop_row_c = ip_ref_to_row.
      SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_stop_row_c LEFT DELETING LEADING space.
      lv_start_row_c = ip_ref_from_row.
      SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_start_row_c LEFT DELETING LEADING space.
      ls_serie-ref = ip_sheet.
      ls_serie-ref = zcl_excel_common=>escape_string( ip_value = ls_serie-ref ).
      CONCATENATE ls_serie-ref '!$' ip_ref_from_col '$' lv_start_row_c ':$' ip_ref_to_col '$' lv_stop_row_c INTO ls_serie-ref.
      CLEAR: lv_start_row_c, lv_stop_row_c.
    ELSE.
      ls_serie-ref = ip_ref.
    ENDIF.
    ls_serie-idx = ip_idx.
    ls_serie-order = ip_order.
    ls_serie-invertifnegative = ip_invertifnegative.
    ls_serie-symbol = ip_symbol.
    ls_serie-smooth = ip_smooth.
    ls_serie-sername = ip_sername.
    APPEND ls_serie TO me->series.
    SORT me->series BY order ASCENDING.
  ENDMETHOD.


  METHOD set_print_lbl.
    me->print_label = ip_value.
  ENDMETHOD.


  METHOD set_style.
    me->ns_c14styleval = ip_style-c14style.
    CONDENSE me->ns_c14styleval NO-GAPS.
    me->ns_styleval = ip_style-cstyle.
    CONDENSE me->ns_styleval NO-GAPS.
  ENDMETHOD.


  METHOD set_title.
    me->title = ip_value.
  ENDMETHOD.
ENDCLASS.
