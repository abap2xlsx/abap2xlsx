*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL42
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  ZDEMO_EXCEL42.
type-POOLS: vrm.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_theme                TYPE REF TO zcl_excel_theme,
      lo_style                type ref to zcl_excel_style,
      lv_style_guid            type ZEXCEL_CELL_STYLE.
DATA: gc_save_file_name TYPE string VALUE '42 Theme Manipulation demo.&'.
include zdemo_excel_outputopt_incl.

initialization.


START-OF-SELECTION.


  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Create a bold / italic style with usage of major font
  lo_style               = lo_excel->add_new_style( ).
  lo_style->font->bold   = abap_true.
  lo_style->font->italic = abap_true.
  lo_style->font->scheme = zcl_excel_style_font=>c_scheme_major.
  lo_style->font->color-rgb  = zcl_excel_style_color=>c_red.
  lv_style_guid          = lo_style->get_guid( ).

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Styles' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Bold text'  ip_style = lv_style_guid ).

  "create theme
  create object lo_theme.
  lo_theme->set_theme_name( iv_name = 'Theme Demo 42 A2X' ).
  lo_theme->set_color_scheme_name( iv_name = 'Demo 42 A2X' ).

  "set theme colors
  lo_theme->set_color(
    exporting
      iv_type         = zcl_excel_theme_color_scheme=>c_dark1
      iv_srgb         = '5F9EA0'
*      iv_syscolorname =
*      iv_syscolorlast =
  ).
    lo_theme->set_color(
    exporting
      iv_type         = zcl_excel_theme_color_scheme=>c_dark2
      iv_srgb         = 'FFA500'
*      iv_syscolorname =
*      iv_syscolorlast =
  ).
     lo_theme->set_color(
    exporting
      iv_type         = zcl_excel_theme_color_scheme=>c_light1
      iv_srgb         = '778899'
*      iv_syscolorname =
*      iv_syscolorlast =
  ).

     lo_theme->set_color(
    exporting
      iv_type         = zcl_excel_theme_color_scheme=>c_light1
      iv_srgb         = '9932CC'
*      iv_syscolorname =
*      iv_syscolorlast =
  ).
  lo_theme->set_font_scheme_name( iv_name = 'Demo 42 A2X' ).


  "set theme latin fonts - major and minor
  lo_theme->set_latin_font(
    exporting
      iv_type        = zcl_excel_theme_font_scheme=>c_major
      iv_typeface    = 'Britannic Bold'
*      iv_panose      =
*      iv_pitchfamily =
*      iv_charset     =
  ).
   lo_theme->set_latin_font(
    exporting
      iv_type        = zcl_excel_theme_font_scheme=>c_minor
      iv_typeface    = 'Broadway'
*      iv_panose      =
*      iv_pitchfamily =
*      iv_charset     =
  ).
    "push theme to file
      lo_excel->set_theme( io_theme = lo_theme ).

   "output
      lcl_output=>output( cl_excel = lo_excel ).
