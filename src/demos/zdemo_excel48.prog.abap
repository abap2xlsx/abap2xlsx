*&---------------------------------------------------------------------*
*& Report zdemo_excel48
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel48.

DATA:
  lo_excel        TYPE REF TO zcl_excel,
  lo_worksheet    TYPE REF TO zcl_excel_worksheet,
  lo_style_1      TYPE REF TO zcl_excel_style,
  lo_style_2      TYPE REF TO zcl_excel_style,
  lv_style_1_guid TYPE zexcel_cell_style,
  lv_style_2_guid TYPE zexcel_cell_style,
  lv_value        TYPE string,
  ls_rtf          TYPE zexcel_s_rtf,
  lt_rtf          TYPE zexcel_t_rtf.


CONSTANTS:
  gc_save_file_name TYPE string VALUE '48_MultipleStylesInOneCell.xlsx'.

INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  lo_worksheet = lo_excel->get_active_worksheet( ).

  lo_style_1 = lo_excel->add_new_style( ).
  lo_style_1->font->color-rgb = 'FF000000'.

  lv_value = 'normal red underline normal red-underline bold italic bigger Times-New-Roman'.

  " red
  lo_style_2 = lo_excel->add_new_style( ).
  lo_style_2->font->color-rgb = 'FFFF0000'.
  ls_rtf-offset = 7.
  ls_rtf-length = 3.
  ls_rtf-font   = lo_style_2->font->get_structure( ).
  INSERT ls_rtf INTO TABLE lt_rtf.

  " underline
  lo_style_2 = lo_excel->add_new_style( ).
  lo_style_2->font->underline = abap_true.
  lo_style_2->font->underline_mode = lo_style_2->font->c_underline_single.
  ls_rtf-offset = 11.
  ls_rtf-length = 9.
  ls_rtf-font   = lo_style_2->font->get_structure( ).
  INSERT ls_rtf INTO TABLE lt_rtf.

  " red and underline
  lo_style_2 = lo_excel->add_new_style( ).
  lo_style_2->font->color-rgb = 'FFFF0000'.
  lo_style_2->font->underline = abap_true.
  lo_style_2->font->underline_mode = lo_style_2->font->c_underline_single.
  ls_rtf-offset = 28.
  ls_rtf-length = 13.
  ls_rtf-font   = lo_style_2->font->get_structure( ).
  INSERT ls_rtf INTO TABLE lt_rtf.

  " bold
  lo_style_2 = lo_excel->add_new_style( ).
  lo_style_2->font->bold = abap_true.
  ls_rtf-offset = 42.
  ls_rtf-length = 4.
  ls_rtf-font   = lo_style_2->font->get_structure( ).
  INSERT ls_rtf INTO TABLE lt_rtf.

  " italic
  lo_style_2 = lo_excel->add_new_style( ).
  lo_style_2->font->italic = abap_true.
  ls_rtf-offset = 47.
  ls_rtf-length = 6.
  ls_rtf-font   = lo_style_2->font->get_structure( ).
  INSERT ls_rtf INTO TABLE lt_rtf.

  " bigger
  lo_style_2 = lo_excel->add_new_style( ).
  lo_style_2->font->size = 28.
  ls_rtf-offset = 54.
  ls_rtf-length = 6.
  ls_rtf-font   = lo_style_2->font->get_structure( ).
  INSERT ls_rtf INTO TABLE lt_rtf.

  " Times-New-Roman
  lo_style_2 = lo_excel->add_new_style( ).
  lo_style_2->font->name = zcl_excel_style_font=>c_name_roman.
  lo_style_2->font->scheme = zcl_excel_style_font=>c_scheme_none.
  lo_style_2->font->family = zcl_excel_style_font=>c_family_roman.

  " Create an underline double style
  lo_style_2                        = lo_excel->add_new_style( ).
  lo_style_2->font->underline       = abap_true.
  lo_style_2->font->underline_mode  = zcl_excel_style_font=>c_underline_double.
  lo_style_2->font->name            = zcl_excel_style_font=>c_name_roman.
  lo_style_2->font->scheme          = zcl_excel_style_font=>c_scheme_none.
  lo_style_2->font->family          = zcl_excel_style_font=>c_family_roman.
  lv_style_2_guid = lo_style_2->get_guid( ).
  ls_rtf-offset = 61.
  ls_rtf-length = 15.
  ls_rtf-font   = lo_style_2->font->get_structure( ).
  INSERT ls_rtf INTO TABLE lt_rtf.

  lv_style_1_guid = lo_style_1->get_guid( ).
  lo_worksheet->set_cell(
    ip_column = 'B'
    ip_row    = 2
    ip_style  = lo_style_1->get_guid( )
    ip_value  = lv_value
    it_rtf    = lt_rtf ).

*** Create output
  lcl_output=>output( lo_excel ).
