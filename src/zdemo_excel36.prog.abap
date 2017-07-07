*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL36
REPORT  zdemo_excel36.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_column               TYPE REF TO zcl_excel_column,
      col                     TYPE i.

DATA: lo_style_arial20      TYPE REF TO zcl_excel_style,
      lo_style_times11      TYPE REF TO zcl_excel_style,
      lo_style_cambria8red  TYPE REF TO zcl_excel_style.

DATA: lv_style_arial20_guid     TYPE zexcel_cell_style,
      lv_style_times11_guid     TYPE zexcel_cell_style,
      lv_style_cambria8red_guid TYPE zexcel_cell_style.


CONSTANTS: gc_save_file_name TYPE string VALUE '36_DefaultStyles.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Create a bold / italic style
  lo_style_arial20                  = lo_excel->add_new_style( ).
  lo_style_arial20->font->name      = zcl_excel_style_font=>c_name_arial.
  lo_style_arial20->font->scheme    = zcl_excel_style_font=>c_scheme_none.
  lo_style_arial20->font->size      = 20.
  lv_style_arial20_guid             = lo_style_arial20->get_guid( ).

  lo_style_times11                  = lo_excel->add_new_style( ).
  lo_style_times11->font->name      = zcl_excel_style_font=>c_name_roman.
  lo_style_times11->font->scheme    = zcl_excel_style_font=>c_scheme_none.
  lo_style_times11->font->size      = 11.
  lv_style_times11_guid             = lo_style_times11->get_guid( ).

  lo_style_cambria8red                  = lo_excel->add_new_style( ).
  lo_style_cambria8red->font->name      = zcl_excel_style_font=>c_name_cambria.
  lo_style_cambria8red->font->scheme    = zcl_excel_style_font=>c_scheme_none.
  lo_style_cambria8red->font->size      = 8.
  lo_style_cambria8red->font->color-rgb = zcl_excel_style_color=>c_red.
  lv_style_cambria8red_guid             = lo_style_cambria8red->get_guid( ).

  lo_excel->set_default_style( lv_style_arial20_guid ).  " Default for all new worksheets

* 1st sheet - do not change anything --> defaultstyle from lo_excel should apply
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Style for complete document' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 4 ip_value = 'All cells in this sheet are set to font Arial, fontsize 20' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 5 ip_value = 'because no separate style was passed for this sheet' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 6 ip_value = 'but a default style was set for the complete instance of zcl_excel' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 1 ip_value = space ). " Missing feature "set active cell - use this to simulate that


* 2nd sheet - defaultstyle for this sheet set explicitly ( set to Times New Roman 11 )
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Style for this sheet' ).
  lo_worksheet->zif_excel_sheet_properties~set_style( lv_style_times11_guid ).

  lo_worksheet->set_cell( ip_column = 2 ip_row = 4 ip_value = 'All  cells in this sheet are set to font Times New Roman, fontsize 11' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 5 ip_value = 'because this style was passed for this sheet' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 6 ip_value = 'thus the default style from zcl_excel does not apply to this sheet' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 1 ip_value = space ). " Missing feature "set active cell - use this to simulate that


* 3rd sheet - defaultstyle for columns  ( set to Times New Roman 11 )
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Style for 3 columns' ).
  lo_column = lo_worksheet->get_column( 'B' ).
  lo_column->set_column_style_by_guid( ip_style_guid = lv_style_times11_guid ).
  lo_column = lo_worksheet->get_column( 'C' ).
  lo_column->set_column_style_by_guid( ip_style_guid = lv_style_times11_guid ).
  lo_column = lo_worksheet->get_column( 'F' ).
  lo_column->set_column_style_by_guid( ip_style_guid = lv_style_times11_guid ).

  lo_worksheet->set_cell( ip_column = 2 ip_row = 4  ip_value = 'The columns B,C and F are set to Times New Roman' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 10 ip_value = 'All other cells in this sheet are set to font Arial, fontsize 20' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 11 ip_value = 'because no separate style was passed for this sheet' ).
  lo_worksheet->set_cell( ip_column = 2 ip_row = 12 ip_value = 'but a default style was set for the complete instance of zcl_excel' ).

  lo_worksheet->set_cell( ip_column = 8 ip_row = 1 ip_value = 'Of course' ip_style = lv_style_cambria8red_guid ).
  lo_worksheet->set_cell( ip_column = 8 ip_row = 2 ip_value = 'setting a specific style to a cell' ip_style = lv_style_cambria8red_guid ).
  lo_worksheet->set_cell( ip_column = 8 ip_row = 3 ip_value = 'takes precedence over all defaults' ip_style = lv_style_cambria8red_guid ).
  lo_worksheet->set_cell( ip_column = 8 ip_row = 4 ip_value = 'Here:  Cambria 8 in red' ip_style = lv_style_cambria8red_guid ).


* Set entry into each of the first 10 columns
  DO 20 TIMES.
    col = sy-index.
    CASE col.
      WHEN 2 " B
        OR 3 " C
        OR 6." F
        lo_worksheet->set_cell( ip_column = col ip_row = 6 ip_value = 'Times 11' ).
      WHEN OTHERS.
        lo_worksheet->set_cell( ip_column = col ip_row = 6 ip_value = 'Arial 20' ).
    ENDCASE.
  ENDDO.

  lo_worksheet->set_cell( ip_column = 2 ip_row = 1 ip_value = space ). " Missing feature "set active cell - use this to simulate that



  lo_excel->set_active_sheet_index( 1 ).


*** Create output
  lcl_output=>output( lo_excel ).
