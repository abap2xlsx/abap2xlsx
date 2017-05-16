*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL12
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel12.

DATA: lo_excel         TYPE REF TO zcl_excel,
      lo_worksheet     TYPE REF TO zcl_excel_worksheet,
      column_dimension TYPE REF TO zcl_excel_worksheet_columndime,
      row_dimension    TYPE REF TO zcl_excel_worksheet_rowdimensi.

DATA: lv_file      TYPE xstring,
      lv_bytecount TYPE i,
      lt_file_tab  TYPE solix_tab.

DATA: lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

CONSTANTS: gc_save_file_name TYPE string VALUE '12_HideSizeOutlineRowsAndColumns.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world in AutoSize column' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Hello world in a column width size 50' ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 4 ip_value = 'Hello world (hidden column)' ).
  lo_worksheet->set_cell( ip_column = 'F' ip_row = 2 ip_value = 'Outline column level 0' ).
  lo_worksheet->set_cell( ip_column = 'G' ip_row = 2 ip_value = 'Outline column level 1' ).
  lo_worksheet->set_cell( ip_column = 'H' ip_row = 2 ip_value = 'Outline column level 2' ).
  lo_worksheet->set_cell( ip_column = 'I' ip_row = 2 ip_value = 'Small' ).


  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello world (hidden row)' ).
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 5 ip_value = 'Hello world in a row height size 20' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 9 ip_value = 'Simple outline rows 10-16 ( collapsed )' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 19 ip_value = '3 Outlines - Outlinelevel 1 is collapsed' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 19 ip_value = 'One of the two inner outlines is expanded, one collapsed' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 20 ip_value = 'Inner outline level - expanded' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 24 ip_value = 'Inner outline level - lines 25-28 are collapsed' ).

  lo_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_off. " By default is on
  lo_worksheet->zif_excel_sheet_properties~summaryright = zif_excel_sheet_properties=>c_right_off. " By default is on

  " Column Settings
  " Auto size
  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'B' ).
  column_dimension->set_auto_size( ip_auto_size = abap_true ).
  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'I' ).
  column_dimension->set_auto_size( ip_auto_size = abap_true ).
  " Manual Width
  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'C' ).
  column_dimension->set_width( ip_width = 50 ).
  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'D' ).
  column_dimension->set_visible( ip_visible = abap_false ).
  " Implementation in the Writer is not working yet ===== TODO =====
  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'F' ).
  column_dimension->set_outline_level( ip_outline_level = 0 ).
  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'G' ).
  column_dimension->set_outline_level( ip_outline_level = 1 ).
  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'H' ).
  column_dimension->set_outline_level( ip_outline_level = 2 ).

  row_dimension = lo_worksheet->get_row_dimension( ip_row = 1 ).
  row_dimension->set_visible( ip_visible = abap_false ).
  row_dimension = lo_worksheet->get_row_dimension( ip_row = 5 ).
  row_dimension->set_row_height( ip_row_height = 20 ).
* obsolete, not intuitive.  Use new method shown below
*  " Implementation in the Writer is not working yet ===== TODO =====
*  row_dimension = lo_worksheet->get_row_dimension( ip_row = 6 ).
*  row_dimension->set_outline_level( ip_outline_level = 0 ).
*  row_dimension = lo_worksheet->get_row_dimension( ip_row = 7 ).
*  row_dimension->set_outline_level( ip_outline_level = 1 ).
*  row_dimension = lo_worksheet->get_row_dimension( ip_row = 8 ).
*  row_dimension->set_outline_level( ip_outline_level = 2 ).

* Define an outline rows 10-16, collapsed on startup
  lo_worksheet->set_row_outline( iv_row_from = 10
                                 iv_row_to   = 16
                                 iv_collapsed = abap_true ).  " collapsed

* Define an inner outline rows 21-22, expanded when outer outline becomes extended
  lo_worksheet->set_row_outline( iv_row_from = 21
                                 iv_row_to   = 22
                                 iv_collapsed = abap_false ). " expanded

* Define an inner outline rows 25-28, collapsed on startup
  lo_worksheet->set_row_outline( iv_row_from = 25
                                 iv_row_to   = 28
                                 iv_collapsed = abap_true ).  " collapsed

* Define an outer outline rows 20-30, collapsed on startup
  lo_worksheet->set_row_outline( iv_row_from = 20
                                 iv_row_to   = 30
                                 iv_collapsed = abap_true ).  " collapsed

* Hint:  the order you create the outlines can be arbitrary
*        You can start with inner outlines or with outer outlines

*--------------------------------------------------------------------*
* Hide columns right of column M
*--------------------------------------------------------------------*
  lo_worksheet->zif_excel_sheet_properties~hide_columns_from = 'M'.

*** Create output
  lcl_output=>output( lo_excel ).
