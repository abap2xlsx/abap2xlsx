*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL13
*&
*&---------------------------------------------------------------------*
*& Example by: Alvaro "Blag" Tejada Galindo.
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel13.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lv_style_bold_border_guid TYPE zexcel_cell_style,
      lo_style_bold_border TYPE REF TO zcl_excel_style,
      lo_border_dark TYPE REF TO zcl_excel_style_border.


CONSTANTS: gc_save_file_name TYPE string VALUE '13_MergedCells.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'sheet1' ).

  CREATE OBJECT lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.

  lo_style_bold_border = lo_excel->add_new_style( ).
  lo_style_bold_border->font->bold = abap_true.
  lo_style_bold_border->font->italic = abap_false.
  lo_style_bold_border->font->color-rgb = zcl_excel_style_color=>c_black.
  lo_style_bold_border->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_bold_border->borders->allborders = lo_border_dark.
  lv_style_bold_border_guid = lo_style_bold_border->get_guid( ).

  lo_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = 'Test' ).

  lo_worksheet->set_cell( ip_row = 2 ip_column = 'B' ip_value = 'Banana' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'C' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'D' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'E' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'F' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'G' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'Apple' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'E' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'F' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'G' ip_value = '' ip_style = lv_style_bold_border_guid ).

  lo_worksheet->set_merge( ip_row = 4 ip_column_start = 'B' ip_column_end = 'G' ).

  " Test also if merge works when oher merged chells are empty
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = 'Tomato' ).
  lo_worksheet->set_merge( ip_row = 6 ip_column_start = 'B' ip_column_end = 'G' ).

  " Test the patch provided by Victor Alekhin to merge cells in one column
  lo_worksheet->set_cell(  ip_row = 8 ip_column       = 'B' ip_value = 'Merge cells also over multiple rows by Victor Alekhin' ).
  lo_worksheet->set_merge( ip_row = 8 ip_column_start = 'B' ip_column_end = 'G' ip_row_to = 10 ).

  " Test the patch provided by Alexander Budeyev with different column merges
  lo_worksheet->set_cell( ip_row = 12 ip_column = 'B' ip_value = 'Merge cells with different merges by Alexander Budeyev' ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'B' ip_value = 'Test' ).

  lo_worksheet->set_cell( ip_row = 13 ip_column = 'D' ip_value = 'Banana' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 14 ip_column = 'D' ip_value = '' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'E' ip_value = 'Apple' ip_style = lv_style_bold_border_guid ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'F' ip_value = '' ip_style = lv_style_bold_border_guid ).

  " Test merge (issue)
  lo_worksheet->set_merge( ip_row = 13 ip_column_start = 'B' ip_column_end = 'C' ip_row_to = 15 ).
  lo_worksheet->set_merge( ip_row = 13 ip_column_start = 'D' ip_column_end = 'D' ip_row_to = 14 ).
  lo_worksheet->set_merge( ip_row = 13 ip_column_start = 'E' ip_column_end = 'F' ).



*** Create output
  lcl_output=>output( lo_excel ).
