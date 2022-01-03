*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel3.

TYPE-POOLS: abap.

TYPES: ty_sflight_lines TYPE TABLE OF sflight.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_column    TYPE REF TO zcl_excel_column.

DATA: ls_table_settings       TYPE zexcel_s_table_settings.


DATA: lv_title TYPE zexcel_sheet_title,
      lt_carr  TYPE TABLE OF scarr,
      row      TYPE zexcel_cell_row VALUE 2,
      ls_error TYPE zcl_excel_worksheet=>mty_s_ignored_errors,
      lt_error TYPE zcl_excel_worksheet=>mty_th_ignored_errors,
      lo_range TYPE REF TO zcl_excel_range.
DATA: lo_data_validation  TYPE REF TO zcl_excel_data_validation.
FIELD-SYMBOLS: <carr> LIKE LINE OF lt_carr.

CONSTANTS: c_airlines TYPE string VALUE 'Airlines'.


CONSTANTS: gc_save_file_name TYPE string VALUE '03_iTab.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_empty TYPE flag.
PARAMETERS: p_checkr NO-DISPLAY TYPE abap_bool.

START-OF-SELECTION.
  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Internal table' ).

  DATA lt_test TYPE ty_sflight_lines.

  IF p_empty <> abap_true.
    IF p_checkr = abap_true.
      PERFORM load_fixed_data_for_checker CHANGING lt_test.
    ELSE.
      SELECT * FROM sflight INTO TABLE lt_test.         "#EC CI_NOWHERE
    ENDIF.
  ENDIF.

  ls_table_settings-table_style       = zcl_excel_table=>builtinstyle_medium2.
  ls_table_settings-show_row_stripes  = abap_true.
  ls_table_settings-nofilters         = abap_true.

  lo_worksheet->bind_table( ip_table          = lt_test
                            is_table_settings = ls_table_settings ).

  lo_worksheet->freeze_panes( ip_num_rows = 3 ). "freeze column headers when scrolling
  IF lines( lt_test ) >= 1.
    ls_error-cell_coords = |B2:B{ lines( lt_test ) + 1 }|.
    ls_error-number_stored_as_text = abap_true.
    INSERT ls_error INTO TABLE lt_error.
    lo_worksheet->set_ignored_errors( lt_error ).
  ENDIF.

  lo_column = lo_worksheet->get_column( ip_column = 'E' ). "make date field a bit wider
  lo_column->set_width( ip_width = 11 ).
  " Add another table for data validations
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lv_title = 'Data Validation'.
  lo_worksheet->set_title( lv_title ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = c_airlines ).
  SELECT * FROM scarr INTO TABLE lt_carr.               "#EC CI_NOWHERE
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



FORM load_fixed_data_for_checker CHANGING ct_test TYPE ty_sflight_lines.
  DATA: lt_lines  TYPE TABLE OF string,
        lv_line   TYPE string,
        lt_fields TYPE TABLE OF string,
        lv_comp   TYPE i,
        lv_field  TYPE string,
        ls_test   TYPE sflight.
  FIELD-SYMBOLS: <lv_field> TYPE simple.

  APPEND 'AA 0017 20171219  422 USD 747-400  385 371 191334 31  28  21  21' TO lt_lines.
  APPEND 'AA 0017 20180309  422 USD 747-400  385 365 189984 31  29  21  20' TO lt_lines.
  APPEND 'AA 0017 20180528  422 USD 747-400  385 374 193482 31  30  21  20' TO lt_lines.
  APPEND 'AA 0017 20180816  422 USD 747-400  385 372 193127 31  30  21  20' TO lt_lines.
  APPEND 'AA 0017 20181104  422 USD 747-400  385  44  23908 31   4  21   3' TO lt_lines.
  APPEND 'AA 0017 20190123  422 USD 747-400  385  40  20347 31   3  21   2' TO lt_lines.
  APPEND 'AZ 0555 20171219  185 EUR 737-800  140 133  32143 12  12  10  10' TO lt_lines.
  APPEND 'AZ 0555 20180309  185 EUR 737-800  140 137  32595 12  12  10  10' TO lt_lines.
  APPEND 'AZ 0555 20180528  185 EUR 737-800  140 134  31899 12  11  10  10' TO lt_lines.
  APPEND 'AZ 0555 20180816  185 EUR 737-800  140 128  29775 12  10  10   9' TO lt_lines.
  APPEND 'AZ 0555 20181104  185 EUR 737-800  140   0      0 12   0  10   0' TO lt_lines.
  APPEND 'AZ 0555 20190123  185 EUR 737-800  140  23   5392 12   1  10   2' TO lt_lines.
  APPEND 'AZ 0789 20171219 1030 EUR 767-200  260 250 307176 21  20  11  11' TO lt_lines.
  APPEND 'AZ 0789 20180309 1030 EUR 767-200  260 252 306054 21  20  11  10' TO lt_lines.
  APPEND 'AZ 0789 20180528 1030 EUR 767-200  260 252 307063 21  20  11  10' TO lt_lines.
  APPEND 'AZ 0789 20180816 1030 EUR 767-200  260 249 300739 21  19  11  10' TO lt_lines.
  APPEND 'AZ 0789 20181104 1030 EUR 767-200  260 104 127647 21   8  11   5' TO lt_lines.
  APPEND 'AZ 0789 20190123 1030 EUR 767-200  260  18  22268 21   1  11   1' TO lt_lines.
  APPEND 'DL 0106 20171217  611 USD A380-800 475 458 324379 30  29  20  20' TO lt_lines.
  APPEND 'DL 0106 20180307  611 USD A380-800 475 458 324330 30  30  20  20' TO lt_lines.
  APPEND 'DL 0106 20180526  611 USD A380-800 475 459 328149 30  29  20  20' TO lt_lines.
  APPEND 'DL 0106 20180814  611 USD A380-800 475 462 326805 30  30  20  18' TO lt_lines.
  APPEND 'DL 0106 20181102  611 USD A380-800 475 167 115554 30  10  20   6' TO lt_lines.
  APPEND 'DL 0106 20190121  611 USD A380-800 475  11   9073 30   1  20   1' TO lt_lines.
  LOOP AT lt_lines INTO lv_line.
    CONDENSE lv_line.
    SPLIT lv_line AT space INTO TABLE lt_fields.
    lv_comp = 2.
    LOOP AT lt_fields INTO lv_field.
      ASSIGN COMPONENT lv_comp OF STRUCTURE ls_test TO <lv_field>.
      <lv_field> = lv_field.
      lv_comp = lv_comp + 1.
    ENDLOOP.
    APPEND ls_test TO ct_test.
  ENDLOOP.
ENDFORM.
