REPORT zdemo_excel40.


DATA: lo_excel         TYPE REF TO zcl_excel,
      lo_worksheet     TYPE REF TO zcl_excel_worksheet,
      lo_style_changer TYPE REF TO zif_excel_style_changer.

DATA: lv_row       TYPE zexcel_cell_row,
      lv_col       TYPE i,
      lv_row_char  TYPE char10,
      lv_value     TYPE string,
      ls_fontcolor TYPE zexcel_style_color_argb.

CONSTANTS: gc_save_file_name TYPE string VALUE '40_Printsettings.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.



START-OF-SELECTION.
  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Demo Printsettings' ).

*--------------------------------------------------------------------*
*  Prepare sheet with trivial data
*    - first 4 columns will have fontocolor set
*    - first 3 rows will  have fontcolor set
*    These marked cells will be used for repeatable rows/columns on printpages
*--------------------------------------------------------------------*
    lo_worksheet->set_area(
          ip_range        = 'A1:T100'
          ip_formula      = 'CHAR(64+COLUMN())&TEXT(ROW(),"????????0")'
          ip_area         = lo_worksheet->c_area-whole ).

    lo_style_changer = zcl_excel_style_changer=>create( lo_excel ).
    lo_style_changer->set_fill_filltype( zcl_excel_style_fill=>c_fill_solid ).
    lo_style_changer->set_fill_fgcolor_rgb( zcl_excel_style_color=>c_yellow ).
    lo_worksheet->change_area_style(
          ip_range        = 'A1:T3'
          ip_style_changer = lo_style_changer ).

    lo_style_changer = zcl_excel_style_changer=>create( lo_excel ).
    lo_style_changer->set_font_color_rgb( zcl_excel_style_color=>c_red ).
    lo_worksheet->change_area_style(
          ip_range        = 'A1:D100'
          ip_style_changer = lo_style_changer ).

*--------------------------------------------------------------------*
*  Printsettings
*--------------------------------------------------------------------*
  TRY.
      lo_worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = 'A'
                                                                            iv_columns_to   = 'D' ).
      lo_worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows(    iv_rows_from    = 1
                                                                            iv_rows_to      = 3 ).
    CATCH zcx_excel .
  ENDTRY.

*** Create output
  lcl_output=>output( lo_excel ).
