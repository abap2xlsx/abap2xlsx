*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel3.

TYPE-POOLS: abap.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_column               TYPE REF TO zcl_excel_column.

DATA: ls_table_settings       TYPE zexcel_s_table_settings.


DATA: lv_title TYPE zexcel_sheet_title,
      lt_carr  TYPE TABLE OF scarr,
      row TYPE zexcel_cell_row VALUE 2,
      lo_range TYPE REF TO zcl_excel_range.
DATA: lo_data_validation  TYPE REF TO zcl_excel_data_validation.
FIELD-SYMBOLS: <carr> LIKE LINE OF lt_carr.

CONSTANTS: c_airlines TYPE string VALUE 'Airlines'.


CONSTANTS: gc_save_file_name TYPE string VALUE '03_iTab.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_empty TYPE flag.

START-OF-SELECTION.
  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Internal table').

  DATA lt_test TYPE TABLE OF sflight.

  IF p_empty <> abap_true.
    SELECT * FROM sflight INTO TABLE lt_test. "#EC CI_NOWHERE
  ENDIF.

  ls_table_settings-table_style       = zcl_excel_table=>builtinstyle_medium2.
  ls_table_settings-show_row_stripes  = abap_true.
  ls_table_settings-nofilters         = abap_true.

  lo_worksheet->bind_table( ip_table          = lt_test
                            is_table_settings = ls_table_settings ).

  lo_worksheet->freeze_panes( ip_num_rows = 3 ). "freeze column headers when scrolling

  lo_column = lo_worksheet->get_column( ip_column = 'E' ). "make date field a bit wider
  lo_column->set_width( ip_width = 11 ).
  " Add another table for data validations
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lv_title = 'Data Validation'.
  lo_worksheet->set_title( lv_title ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = c_airlines ).
  SELECT * FROM scarr INTO TABLE lt_carr. "#EC CI_NOWHERE
  LOOP AT lt_carr ASSIGNING <carr>.
    lo_worksheet->set_cell( ip_row = row ip_column = 'A' ip_value = <carr>-carrid ).
    row = row + 1.
  ENDLOOP.
  row = row - 1.
  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = c_airlines.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'A'
                       ip_start_row     = 2
                       ip_stop_column   = 'A'
                       ip_stop_row      = row ).
  " Set Data Validation
  lo_excel->set_active_sheet_index( 1 ).
  lo_worksheet = lo_excel->get_active_worksheet( ).

  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1    = c_airlines.
  lo_data_validation->cell_row    = 4.
  lo_data_validation->cell_column = 'C'.

*** Create output
  lcl_output=>output( lo_excel ).
