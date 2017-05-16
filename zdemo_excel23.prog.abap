*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL23
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel23.

TYPE-POOLS: abap.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_hyperlink            TYPE REF TO zcl_excel_hyperlink.


CONSTANTS: gc_save_file_name TYPE string VALUE '23_Sheets_with_and_without_grid_lines.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.


  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the first sheet with grid lines and print centered horizontal & vertical' ).
  lo_worksheet->set_show_gridlines( i_show_gridlines = abap_true ).

  lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet2!B2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is a link to the second sheet' ip_hyperlink = lo_hyperlink ).

  lo_worksheet->zif_excel_sheet_protection~protected  = zif_excel_sheet_protection=>c_protected.
  lo_worksheet->zif_excel_sheet_properties~zoomscale        = 150.
  lo_worksheet->ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE_NORMAL = 150.

  lo_worksheet->sheet_setup->vertical_centered   = abap_true.
  lo_worksheet->sheet_setup->horizontal_centered = abap_true.

  " Second sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the second sheet with grid lines in display and print' ).
  lo_worksheet->set_show_gridlines(  i_show_gridlines  = abap_true ).
  lo_worksheet->set_print_gridlines( i_print_gridlines = abap_true ).

  lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet3!B2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is link to the third sheet' ip_hyperlink = lo_hyperlink ).

  lo_worksheet->zif_excel_sheet_protection~protected  = zif_excel_sheet_protection=>c_protected.
  lo_worksheet->zif_excel_sheet_properties~zoomscale                = 160.
  lo_worksheet->ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE_PAGELAYOUTVIEW = 200.

  " Third sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet3' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the third sheet without grid lines in display and print' ).
  lo_worksheet->set_show_gridlines(  i_show_gridlines  = abap_false ).
  lo_worksheet->set_print_gridlines( i_print_gridlines = abap_false ).

  lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet4!B2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is link to the fourth sheet' ip_hyperlink = lo_hyperlink ).

  lo_worksheet->zif_excel_sheet_protection~protected  = zif_excel_sheet_protection=>c_protected.
  lo_worksheet->zif_excel_sheet_properties~zoomscale                  = 170.
  lo_worksheet->ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE_SHEETLAYOUTVIEW  = 150.

  " Fourth sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet4' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the fourth sheet with grid lines and print centered ONLY horizontal' ).
  lo_worksheet->set_show_gridlines( i_show_gridlines = abap_true ).

  lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet1!B2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is link to the first sheet' ip_hyperlink = lo_hyperlink ).

  lo_worksheet->zif_excel_sheet_protection~protected  = zif_excel_sheet_protection=>c_protected.
  lo_worksheet->zif_excel_sheet_properties~zoomscale        = 150.
  lo_worksheet->ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE_NORMAL = 150.

"  lo_worksheet->sheet_setup->vertical_centered   = abap_true.
  lo_worksheet->sheet_setup->horizontal_centered = abap_true.



*** Create output
  lcl_output=>output( lo_excel ).
