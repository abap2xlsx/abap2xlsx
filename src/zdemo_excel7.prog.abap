*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL7
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel7.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_style_cond           TYPE REF TO zcl_excel_style_cond.

DATA: ls_iconset3             TYPE zexcel_conditional_iconset,
      ls_iconset4             TYPE zexcel_conditional_iconset,
      ls_iconset5             TYPE zexcel_conditional_iconset,
      ls_databar              TYPE zexcel_conditional_databar,
      ls_colorscale2          TYPE zexcel_conditional_colorscale,
      ls_colorscale3          TYPE zexcel_conditional_colorscale.

CONSTANTS: gc_save_file_name TYPE string VALUE '07_ConditionalAll.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  ls_iconset3-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset3-cfvo1_value              = '0'.
  ls_iconset3-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset3-cfvo2_value              = '33'.
  ls_iconset3-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset3-cfvo3_value              = '66'.
  ls_iconset3-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

  ls_iconset4-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset4-cfvo1_value              = '0'.
  ls_iconset4-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset4-cfvo2_value              = '25'.
  ls_iconset4-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset4-cfvo3_value              = '50'.
  ls_iconset4-cfvo4_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset4-cfvo4_value              = '75'.
  ls_iconset4-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

  ls_iconset5-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset5-cfvo1_value              = '0'.
  ls_iconset5-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset5-cfvo2_value              = '20'.
  ls_iconset5-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset5-cfvo3_value              = '40'.
  ls_iconset5-cfvo4_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset5-cfvo4_value              = '60'.
  ls_iconset5-cfvo5_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset5-cfvo5_value              = '80'.
  ls_iconset5-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

  ls_databar-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_min.
  ls_databar-cfvo1_value              = '0'.
  ls_databar-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_max.
  ls_databar-cfvo2_value              = '0'.
  ls_databar-colorrgb                 = 'FF638EC6'.

  ls_colorscale2-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_min.
  ls_colorscale2-cfvo1_value              = '0'.
  ls_colorscale2-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percentile.
  ls_colorscale2-cfvo2_value              = '50'.
  ls_colorscale2-colorrgb1                = 'FFF8696B'.
  ls_colorscale2-colorrgb2                = 'FF63BE7B'.

  ls_colorscale3-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_min.
  ls_colorscale3-cfvo1_value              = '0'.
  ls_colorscale3-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percentile.
  ls_colorscale3-cfvo2_value              = '50'.
  ls_colorscale3-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_max.
  ls_colorscale3-cfvo3_value              = '0'.
  ls_colorscale3-colorrgb1                = 'FFF8696B'.
  ls_colorscale3-colorrgb2                = 'FFFFEB84'.
  ls_colorscale3-colorrgb3                = 'FF63BE7B'.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).

* ICONSET

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.

  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3arrows.

  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'B'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'B'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'C_ICONSET_3ARROWS' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'B' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'B' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'B' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3arrowsgray.
  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'C'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'C'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'C_ICONSET_3ARROWSGRAY' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'C' ip_value = 50 ).
  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3flags.
  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'D'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'D'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = 'C_ICONSET_3FLAGS' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'D' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'D' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'D' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'D' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'D' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3trafficlights.
  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'E'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'E'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'E' ip_value = 'C_ICONSET_3TRAFFICLIGHTS' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'E' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'E' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'E' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'E' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'E' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3trafficlights2.
  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'F'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'F'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'F' ip_value = 'C_ICONSET_3TRAFFICLIGHTS2' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'F' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'F' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'F' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'F' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'F' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3signs.
  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'G'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'G'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'G' ip_value = 'C_ICONSET_3SIGNS' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'G' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'G' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'G' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'G' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'G' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3symbols.
  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'H'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'H'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'H' ip_value = 'C_ICONSET_3SYMBOLS' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'H' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'H' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'H' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'H' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'H' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3symbols2.
  lo_style_cond->mode_iconset  = ls_iconset3.
  lo_style_cond->set_range( ip_start_column  = 'I'
                                   ip_start_row     = 5
                                   ip_stop_column   = 'I'
                                   ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'I' ip_value = 'C_ICONSET_3SYMBOLS2' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'I' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'I' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'I' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'I' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'I' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4arrows.
  lo_style_cond->mode_iconset  = ls_iconset4.
  lo_style_cond->set_range( ip_start_column  = 'B'
                                   ip_start_row     = 12
                                   ip_stop_column   = 'B'
                                   ip_stop_row      = 16 ).

  lo_worksheet->set_cell( ip_row = 11 ip_column = 'B' ip_value = 'C_ICONSET_4ARROWS' ).
  lo_worksheet->set_cell( ip_row = 12 ip_column = 'B' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'B' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 14 ip_column = 'B' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 15 ip_column = 'B' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 16 ip_column = 'B' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4arrowsgray.
  lo_style_cond->mode_iconset  = ls_iconset4.
  lo_style_cond->set_range( ip_start_column  = 'C'
                                   ip_start_row     = 12
                                   ip_stop_column   = 'C'
                                   ip_stop_row      = 16 ).

  lo_worksheet->set_cell( ip_row = 11 ip_column = 'C' ip_value = 'C_ICONSET_4ARROWSGRAY' ).
  lo_worksheet->set_cell( ip_row = 12 ip_column = 'C' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'C' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 14 ip_column = 'C' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 15 ip_column = 'C' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 16 ip_column = 'C' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4redtoblack.
  lo_style_cond->mode_iconset  = ls_iconset4.
  lo_style_cond->set_range( ip_start_column  = 'D'
                                   ip_start_row     = 12
                                   ip_stop_column   = 'D'
                                   ip_stop_row      = 16 ).

  lo_worksheet->set_cell( ip_row = 11 ip_column = 'D' ip_value = 'C_ICONSET_4REDTOBLACK' ).
  lo_worksheet->set_cell( ip_row = 12 ip_column = 'D' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'D' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 14 ip_column = 'D' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 15 ip_column = 'D' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 16 ip_column = 'D' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4rating.
  lo_style_cond->mode_iconset  = ls_iconset4.
  lo_style_cond->set_range( ip_start_column  = 'E'
                                   ip_start_row     = 12
                                   ip_stop_column   = 'E'
                                   ip_stop_row      = 16 ).

  lo_worksheet->set_cell( ip_row = 11 ip_column = 'E' ip_value = 'C_ICONSET_4RATING' ).
  lo_worksheet->set_cell( ip_row = 12 ip_column = 'E' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'E' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 14 ip_column = 'E' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 15 ip_column = 'E' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 16 ip_column = 'E' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4trafficlights.
  lo_style_cond->mode_iconset  = ls_iconset4.
  lo_style_cond->set_range( ip_start_column  = 'F'
                                   ip_start_row     = 12
                                   ip_stop_column   = 'F'
                                   ip_stop_row      = 16 ).

  lo_worksheet->set_cell( ip_row = 11 ip_column = 'F' ip_value = 'C_ICONSET_4TRAFFICLIGHTS' ).
  lo_worksheet->set_cell( ip_row = 12 ip_column = 'F' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 13 ip_column = 'F' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 14 ip_column = 'F' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 15 ip_column = 'F' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 16 ip_column = 'F' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5arrows.
  lo_style_cond->mode_iconset  = ls_iconset5.
  lo_style_cond->set_range( ip_start_column  = 'B'
                                   ip_start_row     = 19
                                   ip_stop_column   = 'B'
                                   ip_stop_row      = 23 ).

  lo_worksheet->set_cell( ip_row = 18 ip_column = 'B' ip_value = 'C_ICONSET_5ARROWS' ).
  lo_worksheet->set_cell( ip_row = 19 ip_column = 'B' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 20 ip_column = 'B' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 21 ip_column = 'B' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 22 ip_column = 'B' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 23 ip_column = 'B' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5arrowsgray.
  lo_style_cond->mode_iconset  = ls_iconset5.
  lo_style_cond->set_range( ip_start_column  = 'C'
                                   ip_start_row     = 19
                                   ip_stop_column   = 'C'
                                   ip_stop_row      = 23 ).

  lo_worksheet->set_cell( ip_row = 18 ip_column = 'C' ip_value = 'C_ICONSET_5ARROWSGRAY' ).
  lo_worksheet->set_cell( ip_row = 19 ip_column = 'C' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 20 ip_column = 'C' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 21 ip_column = 'C' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 22 ip_column = 'C' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 23 ip_column = 'C' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5rating.
  lo_style_cond->mode_iconset  = ls_iconset5.
  lo_style_cond->set_range( ip_start_column  = 'D'
                                   ip_start_row     = 19
                                   ip_stop_column   = 'D'
                                   ip_stop_row      = 23 ).

  lo_worksheet->set_cell( ip_row = 18 ip_column = 'D' ip_value = 'C_ICONSET_5RATING' ).
  lo_worksheet->set_cell( ip_row = 19 ip_column = 'D' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 20 ip_column = 'D' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 21 ip_column = 'D' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 22 ip_column = 'D' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 23 ip_column = 'D' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5quarters.
  lo_style_cond->mode_iconset  = ls_iconset5.
  lo_style_cond->set_range( ip_start_column  = 'E'
                                   ip_start_row     = 19
                                   ip_stop_column   = 'E'
                                   ip_stop_row      = 23 ).

* DATABAR

  lo_worksheet->set_cell( ip_row = 25 ip_column = 'B' ip_value = 'DATABAR' ).
  lo_worksheet->set_cell( ip_row = 26 ip_column = 'B' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 27 ip_column = 'B' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 28 ip_column = 'B' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 29 ip_column = 'B' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 30 ip_column = 'B' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule            = zcl_excel_style_cond=>c_rule_databar.
  lo_style_cond->priority        = 1.
  lo_style_cond->mode_databar = ls_databar.
  lo_style_cond->set_range( ip_start_column  = 'B'
                                   ip_start_row     = 26
                                   ip_stop_column   = 'B'
                                   ip_stop_row      = 30 ).

* COLORSCALE

  lo_worksheet->set_cell( ip_row = 25 ip_column = 'C' ip_value = 'COLORSCALE 2 COLORS' ).
  lo_worksheet->set_cell( ip_row = 26 ip_column = 'C' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 27 ip_column = 'C' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 28 ip_column = 'C' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 29 ip_column = 'C' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 30 ip_column = 'C' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule            = zcl_excel_style_cond=>c_rule_colorscale.
  lo_style_cond->priority        = 1.
  lo_style_cond->mode_colorscale = ls_colorscale2.
  lo_style_cond->set_range( ip_start_column  = 'C'
                                   ip_start_row     = 26
                                   ip_stop_column   = 'C'
                                   ip_stop_row      = 30 ).


  lo_worksheet->set_cell( ip_row = 25 ip_column = 'D' ip_value = 'COLORSCALE 3 COLORS' ).
  lo_worksheet->set_cell( ip_row = 26 ip_column = 'D' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 27 ip_column = 'D' ip_value = 20 ).
  lo_worksheet->set_cell( ip_row = 28 ip_column = 'D' ip_value = 30 ).
  lo_worksheet->set_cell( ip_row = 29 ip_column = 'D' ip_value = 40 ).
  lo_worksheet->set_cell( ip_row = 30 ip_column = 'D' ip_value = 50 ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule            = zcl_excel_style_cond=>c_rule_colorscale.
  lo_style_cond->priority        = 1.
  lo_style_cond->mode_colorscale = ls_colorscale3.
  lo_style_cond->set_range( ip_start_column  = 'D'
                                   ip_start_row     = 26
                                   ip_stop_column   = 'D'
                                   ip_stop_row      = 30 ).

*** Create output
  lcl_output=>output( lo_excel ).
