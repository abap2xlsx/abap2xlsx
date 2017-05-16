*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL4
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel4.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,

      lo_hyperlink            TYPE REF TO zcl_excel_hyperlink,

      lv_tabcolor             TYPE zexcel_s_tabcolor,

      ls_header               TYPE zexcel_s_worksheet_head_foot,
      ls_footer               TYPE zexcel_s_worksheet_head_foot.

CONSTANTS: gc_save_file_name TYPE string VALUE '04_Sheets.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet1' ).
  lo_worksheet->zif_excel_sheet_properties~selected = zif_excel_sheet_properties=>c_selected.
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the first sheet' ).
* Set color to tab with sheetname   - Red
  lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = 'FF'
                                                            ip_green = '00'
                                                            ip_blu   = '00' ).
  lo_worksheet->set_tabcolor( lv_tabcolor ).

  lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet2!B2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is link to second sheet' ip_hyperlink = lo_hyperlink ).

  " Page printing settings
  lo_worksheet->sheet_setup->set_page_margins( ip_header = '1' ip_footer = '1' ip_unit = 'cm' ).
  lo_worksheet->sheet_setup->black_and_white   = 'X'.
  lo_worksheet->sheet_setup->fit_to_page       = 'X'.  " you should turn this on to activate fit_to_height and fit_to_width
  lo_worksheet->sheet_setup->fit_to_height     = 0.    " used only if ip_fit_to_page = 'X'
  lo_worksheet->sheet_setup->fit_to_width      = 2.    " used only if ip_fit_to_page = 'X'
  lo_worksheet->sheet_setup->orientation       = zcl_excel_sheet_setup=>c_orientation_landscape.
  lo_worksheet->sheet_setup->page_order        = zcl_excel_sheet_setup=>c_ord_downthenover.
  lo_worksheet->sheet_setup->paper_size        = zcl_excel_sheet_setup=>c_papersize_a4.
  lo_worksheet->sheet_setup->scale             = 80.   " used only if ip_fit_to_page = SPACE

  " Header and Footer
  ls_header-right_value = 'print date &D'.
  ls_header-right_font-size = 8.
  ls_header-right_font-name = zcl_excel_style_font=>c_name_arial.

  ls_footer-left_value = '&Z&F'. "Path / Filename
  ls_footer-left_font = ls_header-right_font.
  ls_footer-right_value = 'page &P of &N'. "page x of y
  ls_footer-right_font = ls_header-right_font.

  lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).


  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet2' ).
* Set color to tab with sheetname   - Green
  lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = '00'
                                                            ip_green = 'FF'
                                                            ip_blu   = '00' ).
  lo_worksheet->set_tabcolor( lv_tabcolor ).
  lo_worksheet->zif_excel_sheet_properties~selected = zif_excel_sheet_properties=>c_selected.
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the second sheet' ).
  lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet1!B2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is link to first sheet' ip_hyperlink = lo_hyperlink ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Sheet3 is hidden' ).

  lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet3' ).
* Set color to tab with sheetname   - Blue
  lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = '00'
                                                            ip_green = '00'
                                                            ip_blu   = 'FF' ).
  lo_worksheet->set_tabcolor( lv_tabcolor ).
  lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_hidden.

  lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet4' ).
* Set color to tab with sheetname   - other color
  lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = '00'
                                                            ip_green = 'FF'
                                                            ip_blu   = 'FF' ).
  lo_worksheet->set_tabcolor( lv_tabcolor ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Cell B3 has value 0' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 0 ).
  lo_worksheet->zif_excel_sheet_properties~show_zeros = zif_excel_sheet_properties=>c_hidezero.

  lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).

  lo_excel->set_active_sheet_index_by_name( 'Sheet1' ).


*** Create output
  lcl_output=>output( lo_excel ).
