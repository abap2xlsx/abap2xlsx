*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL27
*& Test Styles for ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel27.

CONSTANTS: c_fish       TYPE string VALUE 'Fish'.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_range                TYPE REF TO zcl_excel_range,
      lo_data_validation      TYPE REF TO zcl_excel_data_validation,
      lo_style_cond           TYPE REF TO zcl_excel_style_cond,
      lo_style_1              TYPE REF TO zcl_excel_style,
      lo_style_2              TYPE REF TO zcl_excel_style,
      lv_style_1_guid         TYPE zexcel_cell_style,
      lv_style_2_guid         TYPE zexcel_cell_style,
      ls_cellis               TYPE zexcel_conditional_cellis.


DATA: lv_title          TYPE zexcel_sheet_title.

CONSTANTS: gc_save_file_name TYPE string VALUE '27_ConditionalFormatting.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.


  " Creates active sheet
  CREATE OBJECT lo_excel.

  lo_style_1                        = lo_excel->add_new_style( ).
  lo_style_1->fill->filltype        = zcl_excel_style_fill=>c_fill_solid.
  lo_style_1->fill->bgcolor-rgb     = zcl_excel_style_color=>c_green.
  lv_style_1_guid                   = lo_style_1->get_guid( ).

  lo_style_2                        = lo_excel->add_new_style( ).
  lo_style_2->fill->filltype        = zcl_excel_style_fill=>c_fill_solid.
  lo_style_2->fill->bgcolor-rgb     = zcl_excel_style_color=>c_red.
  lv_style_2_guid                   = lo_style_2->get_guid( ).

  " Get active sheet
  lo_worksheet        = lo_excel->get_active_worksheet( ).
  lv_title = 'Data Validation'.
  lo_worksheet->set_title( lv_title ).
  " Set values for dropdown
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = c_fish ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 'Anchovy' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'A' ip_value = 'Carp' ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 'Catfish' ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'A' ip_value = 'Cod' ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'A' ip_value = 'Eel' ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'A' ip_value = 'Haddock' ).

  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = c_fish.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'A'
                       ip_start_row     = 4
                       ip_stop_column   = 'A'
                       ip_stop_row      = 9 ).

  " 1st validation
  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1    = c_fish.
  lo_data_validation->cell_row    = 2.
  lo_data_validation->cell_column = 'C'.
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'C' ip_value = 'Select a value' ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule         = zcl_excel_style_cond=>c_rule_cellis.
  ls_cellis-formula           = '"Anchovy"'.
  ls_cellis-operator          = zcl_excel_style_cond=>c_operator_equal.
  ls_cellis-cell_style        = lv_style_1_guid.
  lo_style_cond->mode_cellis  = ls_cellis.
  lo_style_cond->priority     = 1.
  lo_style_cond->set_range( ip_start_column  = 'C'
                            ip_start_row     = 2
                            ip_stop_column   = 'C'
                            ip_stop_row      = 2 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule         = zcl_excel_style_cond=>c_rule_cellis.
  ls_cellis-formula           = '"Carp"'.
  ls_cellis-operator          = zcl_excel_style_cond=>c_operator_equal.
  ls_cellis-cell_style        = lv_style_2_guid.
  lo_style_cond->mode_cellis  = ls_cellis.
  lo_style_cond->priority     = 2.
  lo_style_cond->set_range( ip_start_column  = 'C'
                            ip_start_row     = 2
                            ip_stop_column   = 'C'
                            ip_stop_row      = 2 ).

*** Create output
  lcl_output=>output( lo_excel ).
