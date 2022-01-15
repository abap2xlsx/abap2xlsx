*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL22
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel22.

TYPES: ty_sflight_lines TYPE TABLE OF sflight.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_worksheet       TYPE REF TO zcl_excel_worksheet,
      lo_style           TYPE REF TO zcl_excel_style,
      lo_style_date      TYPE REF TO zcl_excel_style,
      lo_style_editable  TYPE REF TO zcl_excel_style,
      lo_data_validation TYPE REF TO zcl_excel_data_validation.

DATA: lt_field_catalog      TYPE zexcel_t_fieldcatalog,
      ls_table_settings     TYPE zexcel_s_table_settings,
      ls_table_settings_out TYPE zexcel_s_table_settings.

DATA: lv_style_guid           TYPE zexcel_cell_style.

DATA: lv_row            TYPE char10.

FIELD-SYMBOLS: <fs_field_catalog> TYPE zexcel_s_fieldcatalog.

CONSTANTS: gc_save_file_name TYPE string VALUE '22_itab_fieldcatalog.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'PN_MASSIVE' ).

  DATA lt_test TYPE ty_sflight_lines.
  PERFORM load_fixed_data CHANGING lt_test.

  " sheet style (white background)
  lo_style = lo_excel->add_new_style( ).
  lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_white.
  lv_style_guid = lo_style->get_guid( ).

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->zif_excel_sheet_properties~set_style( lv_style_guid ).
  lo_worksheet->zif_excel_sheet_protection~protected  = zif_excel_sheet_protection=>c_protected.
  lo_worksheet->zif_excel_sheet_protection~password   = zcl_excel_common=>encrypt_password( 'test' ).
  lo_worksheet->zif_excel_sheet_protection~sheet      = zif_excel_sheet_protection=>c_active.
  lo_worksheet->zif_excel_sheet_protection~objects    = zif_excel_sheet_protection=>c_active.
  lo_worksheet->zif_excel_sheet_protection~scenarios  = zif_excel_sheet_protection=>c_active.

  " Create cell style for display only fields
  lo_style = lo_excel->add_new_style( ).
  lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_gray.
  lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_text.

  " Create cell style for display only date field
  lo_style_date = lo_excel->add_new_style( ).
  lo_style_date->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style_date->fill->fgcolor-rgb  = zcl_excel_style_color=>c_gray.
  lo_style_date->number_format->format_code = zcl_excel_style_number_format=>c_format_date_ddmmyyyy.

  " Create cell style for editable fields
  lo_style_editable = lo_excel->add_new_style( ).
  lo_style_editable->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.

  lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_test ).

  LOOP AT lt_field_catalog ASSIGNING <fs_field_catalog>.
    CASE <fs_field_catalog>-fieldname.
      WHEN 'CARRID'.
        <fs_field_catalog>-position   = 3.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-style      = lo_style->get_guid( ).
      WHEN 'CONNID'.
        <fs_field_catalog>-position   = 1.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-style      = lo_style->get_guid( ).
        <fs_field_catalog>-scrtext_m  = |Flight\r\nNumber|. " Demonstrates header on 2 lines
      WHEN 'FLDATE'.
        <fs_field_catalog>-position   = 2.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-style      = lo_style_date->get_guid( ).
      WHEN 'PRICE'.
        <fs_field_catalog>-position   = 4.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-style      = lo_style_editable->get_guid( ).
        <fs_field_catalog>-totals_function = zcl_excel_table=>totals_function_sum.
      WHEN OTHERS.
        <fs_field_catalog>-dynpfld = abap_false.
    ENDCASE.
  ENDLOOP.

  ls_table_settings-table_style  = zcl_excel_table=>builtinstyle_medium2.
  ls_table_settings-show_row_stripes = abap_true.

  lo_worksheet->bind_table( EXPORTING
                              ip_table          = lt_test
                              it_field_catalog  = lt_field_catalog
                              is_table_settings = ls_table_settings
                            IMPORTING
                              es_table_settings = ls_table_settings_out ).

  lo_worksheet->freeze_panes( ip_num_rows = 3 ). "freeze column headers when scrolling

  lo_data_validation                  = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type            = zcl_excel_data_validation=>c_type_custom.
  lv_row = ls_table_settings_out-top_left_row.
  CONDENSE lv_row.
  CONCATENATE 'ISNUMBER(' ls_table_settings_out-top_left_column lv_row ')' INTO lo_data_validation->formula1.
  lo_data_validation->cell_row        = ls_table_settings_out-top_left_row.
  lo_data_validation->cell_column     = ls_table_settings_out-top_left_column.
  lo_data_validation->cell_row_to     = ls_table_settings_out-bottom_right_row.
  lo_data_validation->cell_column_to  = ls_table_settings_out-bottom_right_column.



*** Create output
  lcl_output=>output( lo_excel ).


FORM load_fixed_data CHANGING ct_test TYPE ty_sflight_lines.
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
