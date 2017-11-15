class ZCL_EXCEL_GRAPH definition
  public
  create public .

public section.

  types:
*"* public components of class ZCL_EXCEL_GRAPH
*"* do not include other source files here!!!
    BEGIN OF s_style,
              c14style type i,
              cstyle   type i,
         end of s_style .
  types:
    BEGIN OF s_series,
                  idx               TYPE i,
                  order             TYPE i,
                  invertifnegative  TYPE string,
                  symbol            TYPE string,
                  smooth            TYPE string,
                  lbl               TYPE string,
                  ref               TYPE string,
                  sername           TYPE string,
          END OF s_series .
  types:
    t_series TYPE STANDARD TABLE OF s_series .
  types:
    BEGIN OF s_pagemargins,
                  b TYPE string,
                  l TYPE string,
                  r TYPE string,
                  t TYPE string,
                  header TYPE string,
                  footer TYPE string,
          END OF s_pagemargins .

  data NS_1904VAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data NS_LANGVAL type STRING value 'it-IT'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data NS_ROUNDEDCORNERSVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data PAGEMARGINS type S_PAGEMARGINS .
  data NS_AUTOTITLEDELETEDVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data NS_PLOTVISONLYVAL type STRING value '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data NS_DISPBLANKSASVAL type STRING value 'gap'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWDLBLSOVERMAXVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data TITLE type STRING .               .  . " .
  data SERIES type T_SERIES .
  data NS_C14STYLEVAL type STRING value '102'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data PRINT_LABEL type C value 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  data NS_STYLEVAL type STRING value '2'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
  constants:
    BEGIN OF c_style_default,
              c14style type i value 102,
              cstyle   type i value 2,
             END OF c_style_default .
  constants:
    BEGIN OF c_style_1,
              c14style type i value 101,
              cstyle   type i value 1,
             END OF c_style_1 .
  constants:
    BEGIN OF c_style_3,
              c14style type i value 103,
              cstyle   type i value 3,
             END OF c_style_3 .
  constants:
    BEGIN OF c_style_4,
              c14style type i value 104,
              cstyle   type i value 4,
             END OF c_style_4 .
  constants:
    BEGIN OF c_style_5,
              c14style type i value 105,
              cstyle   type i value 5,
             END OF c_style_5 .
  constants:
    BEGIN OF c_style_6,
              c14style type i value 106,
              cstyle   type i value 6,
             END OF c_style_6 .
  constants:
    BEGIN OF c_style_7,
              c14style type i value 107,
              cstyle   type i value 7,
             END OF c_style_7 .
  constants:
    BEGIN OF c_style_8,
              c14style type i value 108,
              cstyle   type i value 8,
             END OF c_style_8 .
  constants:
    BEGIN OF c_style_9,
              c14style type i value 109,
              cstyle   type i value 9,
             END OF c_style_9 .
  constants:
    BEGIN OF c_style_10,
              c14style type i value 110,
              cstyle   type i value 10,
             END OF c_style_10 .
  constants:
    BEGIN OF c_style_11,
              c14style type i value 111,
              cstyle   type i value 11,
             END OF c_style_11 .
  constants:
    BEGIN OF c_style_12,
              c14style type i value 112,
              cstyle   type i value 12,
             END OF c_style_12 .
  constants:
    BEGIN OF c_style_13,
              c14style type i value 113,
              cstyle   type i value 13,
             END OF c_style_13 .
  constants:
    BEGIN OF c_style_14,
              c14style type i value 114,
              cstyle   type i value 14,
             END OF c_style_14 .
  constants:
    BEGIN OF c_style_15,
              c14style type i value 115,
              cstyle   type i value 15,
             END OF c_style_15 .
  constants:
    BEGIN OF c_style_16,
              c14style type i value 116,
              cstyle   type i value 16,
             END OF c_style_16 .
  constants:
    BEGIN OF c_style_17,
              c14style type i value 117,
              cstyle   type i value 17,
             END OF c_style_17 .
  constants:
    BEGIN OF c_style_18,
              c14style type i value 118,
              cstyle   type i value 18,
             END OF c_style_18 .
  constants:
    BEGIN OF c_style_19,
              c14style type i value 119,
              cstyle   type i value 19,
             END OF c_style_19 .
  constants:
    BEGIN OF c_style_20,
              c14style type i value 120,
              cstyle   type i value 20,
             END OF c_style_20 .
  constants:
    BEGIN OF c_style_21,
              c14style type i value 121,
              cstyle   type i value 21,
             END OF c_style_21 .
  constants:
    BEGIN OF c_style_22,
              c14style type i value 122,
              cstyle   type i value 22,
             END OF c_style_22 .
  constants:
    BEGIN OF c_style_23,
              c14style type i value 123,
              cstyle   type i value 23,
             END OF c_style_23 .
  constants:
    BEGIN OF c_style_24,
              c14style type i value 124,
              cstyle   type i value 24,
             END OF c_style_24 .
  constants:
    BEGIN OF c_style_25,
              c14style type i value 125,
              cstyle   type i value 25,
             END OF c_style_25 .
  constants:
    BEGIN OF c_style_26,
              c14style type i value 126,
              cstyle   type i value 26,
             END OF c_style_26 .
  constants:
    BEGIN OF c_style_27,
              c14style type i value 127,
              cstyle   type i value 27,
             END OF c_style_27 .
  constants:
    BEGIN OF c_style_28,
              c14style type i value 128,
              cstyle   type i value 28,
             END OF c_style_28 .
  constants:
    BEGIN OF c_style_29,
              c14style type i value 129,
              cstyle   type i value 29,
             END OF c_style_29 .
  constants:
    BEGIN OF c_style_30,
              c14style type i value 130,
              cstyle   type i value 30,
             END OF c_style_30 .
  constants:
    BEGIN OF c_style_31,
              c14style type i value 131,
              cstyle   type i value 31,
             END OF c_style_31 .
  constants:
    BEGIN OF c_style_32,
              c14style type i value 132,
              cstyle   type i value 32,
             END OF c_style_32 .
  constants:
    BEGIN OF c_style_33,
              c14style type i value 133,
              cstyle   type i value 33,
             END OF c_style_33 .
  constants:
    BEGIN OF c_style_34,
              c14style type i value 134,
              cstyle   type i value 34,
             END OF c_style_34 .
  constants:
    BEGIN OF c_style_35,
              c14style type i value 135,
              cstyle   type i value 35,
             END OF c_style_35 .
  constants:
    BEGIN OF c_style_36,
              c14style type i value 136,
              cstyle   type i value 36,
             END OF c_style_36 .
  constants:
    BEGIN OF c_style_37,
              c14style type i value 137,
              cstyle   type i value 37,
             END OF c_style_37 .
  constants:
    BEGIN OF c_style_38,
              c14style type i value 138,
              cstyle   type i value 38,
             END OF c_style_38 .
  constants:
    BEGIN OF c_style_39,
              c14style type i value 139,
              cstyle   type i value 39,
             END OF c_style_39 .
  constants:
    BEGIN OF c_style_40,
              c14style type i value 140,
              cstyle   type i value 40,
             END OF c_style_40 .
  constants:
    BEGIN OF c_style_41,
              c14style type i value 141,
              cstyle   type i value 41,
             END OF c_style_41 .
  constants:
    BEGIN OF c_style_42,
              c14style type i value 142,
              cstyle   type i value 42,
             END OF c_style_42 .
  constants:
    BEGIN OF c_style_43,
              c14style type i value 143,
              cstyle   type i value 43,
             END OF c_style_43 .
  constants:
    BEGIN OF c_style_44,
              c14style type i value 144,
              cstyle   type i value 44,
             END OF c_style_44 .
  constants:
    BEGIN OF c_style_45,
              c14style type i value 145,
              cstyle   type i value 45,
             END OF c_style_45 .
  constants:
    BEGIN OF c_style_46,
              c14style type i value 146,
              cstyle   type i value 46,
             END OF c_style_46 .
  constants:
    BEGIN OF c_style_47,
              c14style type i value 147,
              cstyle   type i value 47,
             END OF c_style_47 .
  constants:
    BEGIN OF c_style_48,
              c14style type i value 148,
              cstyle   type i value 48,
             END OF c_style_48 .
  constants C_SHOW_TRUE type C value '1'. "#EC NOTEXT
  constants C_SHOW_FALSE type C value '0'. "#EC NOTEXT
  constants C_PRINT_LBL_TRUE type C value '1'. "#EC NOTEXT
  constants C_PRINT_LBL_FALSE type C value '0'. "#EC NOTEXT

  methods CONSTRUCTOR .
  methods CREATE_SERIE
    importing
      !IP_IDX type I optional
      !IP_ORDER type I
      !IP_INVERTIFNEGATIVE type STRING optional
      !IP_SYMBOL type STRING optional
      !IP_SMOOTH type C optional
      !IP_LBL_FROM_COL type ZEXCEL_CELL_COLUMN_ALPHA optional
      !IP_LBL_FROM_ROW type ZEXCEL_CELL_ROW optional
      !IP_LBL_TO_COL type ZEXCEL_CELL_COLUMN_ALPHA optional
      !IP_LBL_TO_ROW type ZEXCEL_CELL_ROW optional
      !IP_LBL type STRING optional
      !IP_REF_FROM_COL type ZEXCEL_CELL_COLUMN_ALPHA optional
      !IP_REF_FROM_ROW type ZEXCEL_CELL_ROW optional
      !IP_REF_TO_COL type ZEXCEL_CELL_COLUMN_ALPHA optional
      !IP_REF_TO_ROW type ZEXCEL_CELL_ROW optional
      !IP_REF type STRING optional
      !IP_SERNAME type STRING
      !IP_SHEET type ZEXCEL_SHEET_TITLE optional .
  methods SET_STYLE
    importing
      !IP_STYLE type S_STYLE .
  methods SET_PRINT_LBL
    importing
      !IP_VALUE type C .
  methods SET_TITLE
    importing
      value(IP_VALUE) type STRING .
protected section.
*"* protected components of class ZCL_EXCEL_GRAPH
*"* do not include other source files here!!!
private section.
*"* private components of class ZCL_EXCEL_GRAPH
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_GRAPH IMPLEMENTATION.


method CONSTRUCTOR.
  "Load default values
  me->pagemargins-b = '0.75'.
  me->pagemargins-l = '0.7'.
  me->pagemargins-r = '0.7'.
  me->pagemargins-t = '0.75'.
  me->pagemargins-header = '0.3'.
  me->pagemargins-footer = '0.3'.
  endmethod.


method CREATE_SERIE.
  DATA ls_serie TYPE s_series.

  DATA: lv_start_row_c  TYPE char7,
        lv_stop_row_c   TYPE char7.


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
  endmethod.


method SET_PRINT_LBL.
  me->print_label = ip_value.
  endmethod.


method SET_STYLE.
  me->ns_c14styleval = ip_style-c14style.
  CONDENSE me->ns_c14styleval NO-GAPS.
  me->ns_styleval = ip_style-cstyle.
  CONDENSE me->ns_styleval NO-GAPS.
  endmethod.


method SET_TITLE.
  me->title = ip_value.
endmethod.
ENDCLASS.
