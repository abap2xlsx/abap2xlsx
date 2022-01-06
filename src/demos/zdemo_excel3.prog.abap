*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel3.

TYPES: ty_sflight_lines TYPE TABLE OF sflight,
       ty_scarr_lines   TYPE TABLE OF scarr.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_column    TYPE REF TO zcl_excel_column.

DATA: ls_table_settings       TYPE zexcel_s_table_settings.


DATA: lv_title TYPE zexcel_sheet_title,
      lt_carr  TYPE ty_scarr_lines,
      row      TYPE zexcel_cell_row VALUE 2,
      ls_error TYPE zcl_excel_worksheet=>mty_s_ignored_errors,
      lt_error TYPE zcl_excel_worksheet=>mty_th_ignored_errors,
      lo_range TYPE REF TO zcl_excel_range.
DATA: lo_data_validation TYPE REF TO zcl_excel_data_validation.
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

  IF p_checkr = abap_true.
    PERFORM set_column_headers USING lo_worksheet
      'Airline;Flight Number;Date;Airfare;Airline Currency;Plane Type;Max. capacity econ.;Occupied econ.;Total;Max. capacity bus.;Occupied bus.;Max. capacity 1st;Occupied 1st'.
  ENDIF.

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
  IF p_checkr = abap_true.
    PERFORM load_scarr_data_for_checker CHANGING lt_carr.
  ELSE.
    SELECT * FROM scarr INTO TABLE lt_carr.             "#EC CI_NOWHERE
  ENDIF.
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
  DATA: lt_lines  TYPE TABLE OF string.

  APPEND 'AA;0017;20171219; 422;USD;747-400 ;385;371;191334;31;28;21;21' TO lt_lines.
  APPEND 'AA;0017;20180309; 422;USD;747-400 ;385;365;189984;31;29;21;20' TO lt_lines.
  APPEND 'AA;0017;20180528; 422;USD;747-400 ;385;374;193482;31;30;21;20' TO lt_lines.
  APPEND 'AA;0017;20180816; 422;USD;747-400 ;385;372;193127;31;30;21;20' TO lt_lines.
  APPEND 'AA;0017;20181104; 422;USD;747-400 ;385; 44; 23908;31; 4;21; 3' TO lt_lines.
  APPEND 'AA;0017;20190123; 422;USD;747-400 ;385; 40; 20347;31; 3;21; 2' TO lt_lines.
  APPEND 'AZ;0555;20171219; 185;EUR;737-800 ;140;133; 32143;12;12;10;10' TO lt_lines.
  APPEND 'AZ;0555;20180309; 185;EUR;737-800 ;140;137; 32595;12;12;10;10' TO lt_lines.
  APPEND 'AZ;0555;20180528; 185;EUR;737-800 ;140;134; 31899;12;11;10;10' TO lt_lines.
  APPEND 'AZ;0555;20180816; 185;EUR;737-800 ;140;128; 29775;12;10;10; 9' TO lt_lines.
  APPEND 'AZ;0555;20181104; 185;EUR;737-800 ;140;  0;     0;12; 0;10; 0' TO lt_lines.
  APPEND 'AZ;0555;20190123; 185;EUR;737-800 ;140; 23;  5392;12; 1;10; 2' TO lt_lines.
  APPEND 'AZ;0789;20171219;1030;EUR;767-200 ;260;250;307176;21;20;11;11' TO lt_lines.
  APPEND 'AZ;0789;20180309;1030;EUR;767-200 ;260;252;306054;21;20;11;10' TO lt_lines.
  APPEND 'AZ;0789;20180528;1030;EUR;767-200 ;260;252;307063;21;20;11;10' TO lt_lines.
  APPEND 'AZ;0789;20180816;1030;EUR;767-200 ;260;249;300739;21;19;11;10' TO lt_lines.
  APPEND 'AZ;0789;20181104;1030;EUR;767-200 ;260;104;127647;21; 8;11; 5' TO lt_lines.
  APPEND 'AZ;0789;20190123;1030;EUR;767-200 ;260; 18; 22268;21; 1;11; 1' TO lt_lines.
  APPEND 'DL;0106;20171217; 611;USD;A380-800;475;458;324379;30;29;20;20' TO lt_lines.
  APPEND 'DL;0106;20180307; 611;USD;A380-800;475;458;324330;30;30;20;20' TO lt_lines.
  APPEND 'DL;0106;20180526; 611;USD;A380-800;475;459;328149;30;29;20;20' TO lt_lines.
  APPEND 'DL;0106;20180814; 611;USD;A380-800;475;462;326805;30;30;20;18' TO lt_lines.
  APPEND 'DL;0106;20181102; 611;USD;A380-800;475;167;115554;30;10;20; 6' TO lt_lines.
  APPEND 'DL;0106;20190121; 611;USD;A380-800;475; 11;  9073;30; 1;20; 1' TO lt_lines.

  PERFORM load_data USING lt_lines CHANGING ct_test.
ENDFORM.

FORM load_scarr_data_for_checker CHANGING ct_scarr TYPE ty_scarr_lines.
  DATA: lt_lines TYPE TABLE OF string.

  APPEND 'AA;American Airlines;USD;http://www.aa.com       ' TO lt_lines.
  APPEND 'AZ;Alitalia         ;EUR;http://www.alitalia.it  ' TO lt_lines.
  APPEND 'DL;Delta Airlines   ;USD;http://www.delta-air.com' TO lt_lines.

  PERFORM load_data USING lt_lines CHANGING ct_scarr.
ENDFORM.

FORM load_data USING it_data TYPE table CHANGING ct_data TYPE table.
  DATA: lv_line     TYPE string,
        lt_fields   TYPE TABLE OF string,
        lv_comp     TYPE i,
        lv_field    TYPE string,
        lv_ref_line TYPE REF TO data.
  FIELD-SYMBOLS:
    <lv_field> TYPE simple,
    <ls_line>  TYPE any.

  CREATE DATA lv_ref_line LIKE LINE OF ct_data.
  ASSIGN lv_ref_line->* TO <ls_line>.

  LOOP AT it_data INTO lv_line.
    CLEAR <ls_line>.
    SPLIT lv_line AT ';' INTO TABLE lt_fields.
    lv_comp = 2.
    LOOP AT lt_fields INTO lv_field.
      ASSIGN COMPONENT lv_comp OF STRUCTURE <ls_line> TO <lv_field>.
      <lv_field> = lv_field.
      lv_comp = lv_comp + 1.
    ENDLOOP.
    APPEND <ls_line> TO ct_data.
  ENDLOOP.
ENDFORM.

FORM set_column_headers
    USING io_worksheet TYPE REF TO zcl_excel_worksheet
          iv_headers   TYPE csequence
    RAISING zcx_excel.

  DATA: lt_headers TYPE TABLE OF string,
        lv_header  TYPE string,
        lv_tabix   TYPE i.

  SPLIT iv_headers AT ';' INTO TABLE lt_headers.
  LOOP AT lt_headers INTO lv_header.
    lv_tabix = sy-tabix.
    io_worksheet->set_cell( ip_row = 1 ip_column = lv_tabix ip_value = lv_header ).
  ENDLOOP.

ENDFORM.
