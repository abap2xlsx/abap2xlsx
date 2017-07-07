*&---------------------------------------------------------------------*
*& Report  ZANGRY_BIRDS
*& Just for fun
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zangry_birds.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_border_light         TYPE REF TO zcl_excel_style_border,
      lo_style_color0         TYPE REF TO zcl_excel_style,
      lo_style_color1         TYPE REF TO zcl_excel_style,
      lo_style_color2         TYPE REF TO zcl_excel_style,
      lo_style_color3         TYPE REF TO zcl_excel_style,
      lo_style_color4         TYPE REF TO zcl_excel_style,
      lo_style_color5         TYPE REF TO zcl_excel_style,
      lo_style_color6         TYPE REF TO zcl_excel_style,
      lo_style_color7         TYPE REF TO zcl_excel_style,
      lo_style_credit         TYPE REF TO zcl_excel_style,
      lo_style_link           TYPE REF TO zcl_excel_style,
      lo_column               TYPE REF TO zcl_excel_column,
      lo_row                  TYPE REF TO zcl_excel_row,
      lo_hyperlink            TYPE REF TO zcl_excel_hyperlink.

DATA: lv_style_color0_guid    TYPE zexcel_cell_style,
      lv_style_color1_guid    TYPE zexcel_cell_style,
      lv_style_color2_guid    TYPE zexcel_cell_style,
      lv_style_color3_guid    TYPE zexcel_cell_style,
      lv_style_color4_guid    TYPE zexcel_cell_style,
      lv_style_color5_guid    TYPE zexcel_cell_style,
      lv_style_color6_guid    TYPE zexcel_cell_style,
      lv_style_color7_guid    TYPE zexcel_cell_style,
      lv_style_credit_guid    TYPE zexcel_cell_style,
      lv_style_link_guid      TYPE zexcel_cell_style,
      lv_style                TYPE zexcel_cell_style.

DATA: lv_col_str        TYPE zexcel_cell_column_alpha,
      lv_row            TYPE i,
      lv_col            TYPE i,
      lt_mapper         TYPE TABLE OF zexcel_cell_style,
      ls_mapper         TYPE zexcel_cell_style.

DATA: lv_file           TYPE xstring,
      lv_bytecount      TYPE i,
      lt_file_tab       TYPE solix_tab.

DATA: lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

CONSTANTS: lv_default_file_name TYPE string VALUE 'angry_birds.xlsx'.

PARAMETERS: p_path TYPE zexcel_export_dir.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( EXPORTING initial_folder  = lv_workdir
                                              CHANGING  selected_folder = lv_workdir ).
  p_path = lv_workdir.

INITIALIZATION.
  cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

START-OF-SELECTION.

  IF p_path IS INITIAL.
    p_path = lv_workdir.
  ENDIF.
  cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_file_separator ).
  CONCATENATE p_path lv_file_separator lv_default_file_name INTO lv_full_path.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  CREATE OBJECT lo_border_light.
  lo_border_light->border_color-rgb = zcl_excel_style_color=>c_white.
  lo_border_light->border_style = zcl_excel_style_border=>c_border_thin.

  " Create color white
  lo_style_color0                      = lo_excel->add_new_style( ).
  lo_style_color0->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color0->fill->fgcolor-rgb   = 'FFFFFFFF'.
  lo_style_color0->borders->allborders = lo_border_light.
  lv_style_color0_guid                 = lo_style_color0->get_guid( ).

  " Create color black
  lo_style_color1                      = lo_excel->add_new_style( ).
  lo_style_color1->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color1->fill->fgcolor-rgb   = 'FF252525'.
  lo_style_color1->borders->allborders = lo_border_light.
  lv_style_color1_guid                 = lo_style_color1->get_guid( ).

  " Create color dark green
  lo_style_color2                      = lo_excel->add_new_style( ).
  lo_style_color2->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color2->fill->fgcolor-rgb   = 'FF75913A'.
  lo_style_color2->borders->allborders = lo_border_light.
  lv_style_color2_guid                 = lo_style_color2->get_guid( ).

  " Create color light green
  lo_style_color3                      = lo_excel->add_new_style( ).
  lo_style_color3->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color3->fill->fgcolor-rgb   = 'FF9DFB73'.
  lo_style_color3->borders->allborders = lo_border_light.
  lv_style_color3_guid                 = lo_style_color3->get_guid( ).

  " Create color green
  lo_style_color4                      = lo_excel->add_new_style( ).
  lo_style_color4->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color4->fill->fgcolor-rgb   = 'FF92CF56'.
  lo_style_color4->borders->allborders = lo_border_light.
  lv_style_color4_guid                 = lo_style_color4->get_guid( ).

  " Create color 2dark green
  lo_style_color5                      = lo_excel->add_new_style( ).
  lo_style_color5->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color5->fill->fgcolor-rgb   = 'FF506228'.
  lo_style_color5->borders->allborders = lo_border_light.
  lv_style_color5_guid                 = lo_style_color5->get_guid( ).

  " Create color yellow
  lo_style_color6                      = lo_excel->add_new_style( ).
  lo_style_color6->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color6->fill->fgcolor-rgb   = 'FFC3E224'.
  lo_style_color6->borders->allborders = lo_border_light.
  lv_style_color6_guid                 = lo_style_color6->get_guid( ).

  " Create color yellow
  lo_style_color7                      = lo_excel->add_new_style( ).
  lo_style_color7->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color7->fill->fgcolor-rgb   = 'FFB3C14F'.
  lo_style_color7->borders->allborders = lo_border_light.
  lv_style_color7_guid                 = lo_style_color7->get_guid( ).

  " Credits
  lo_style_credit = lo_excel->add_new_style( ).
  lo_style_credit->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_credit->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
  lo_style_credit->font->size   = 20.
  lv_style_credit_guid = lo_style_credit->get_guid( ).

  " Link
  lo_style_link = lo_excel->add_new_style( ).
  lo_style_link->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_link->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
*  lo_style_link->font->size   = 20.
  lv_style_link_guid = lo_style_link->get_guid( ).

  " Create image map                                                          " line 2
  DO 30 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 3
  DO 28 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 5 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 4
  DO 27 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 5
  DO 9 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 15 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 6
  DO 7 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 13 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 7
  DO 6 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 5 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 11 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 5 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 8
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 9
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 10
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 11
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 7 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 12
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 13
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 8 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 14
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 12 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 15
  DO 6 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 8 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 16
  DO 7 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 7 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 5 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 17
  DO 8 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 13 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 18
  DO 6 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 23 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 19
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 27 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 20
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 23 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 21
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 19 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 22
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 17 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 23
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 17 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 8 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 24
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 10 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 25
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 8 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 26
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color6_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 27
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color6_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 28
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color6_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 29
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 30
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 31
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color4_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 32
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 8 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color5_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 33
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 34
  DO 3 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 10 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 35
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 11 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 36
  DO 4 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 10 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 11 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 37
  DO 5 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 10 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color7_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 11 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 38
  DO 6 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 10 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 11 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 39
  DO 7 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 22 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 1 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 40
  DO 7 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 17 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 41
  DO 8 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 3 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 15 TIMES. APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 42
  DO 9 TIMES.  APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 5 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 9 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 2 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 43
  DO 11 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 5 TIMES.  APPEND lv_style_color3_guid TO lt_mapper. ENDDO.
  DO 7 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 44
  DO 13 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 6 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  DO 4 TIMES.  APPEND lv_style_color2_guid TO lt_mapper. ENDDO.
  DO 8 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 45
  DO 16 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 13 TIMES. APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape
                                                            " line 46
  DO 18 TIMES. APPEND lv_style_color0_guid TO lt_mapper. ENDDO.
  DO 8 TIMES.  APPEND lv_style_color1_guid TO lt_mapper. ENDDO.
  APPEND INITIAL LINE TO lt_mapper. " escape

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Angry Birds' ).

  lv_row = 1.
  lv_col = 1.

  LOOP AT lt_mapper INTO ls_mapper.
    lv_col_str = zcl_excel_common=>convert_column2alpha( lv_col ).
    IF ls_mapper IS INITIAL.
      lo_row = lo_worksheet->get_row( ip_row = lv_row ).
      lo_row->set_row_height( ip_row_height = 8 ).
      lv_col = 1.
      lv_row = lv_row + 1.
      CONTINUE.
    ENDIF.
    lo_worksheet->set_cell( ip_column = lv_col_str
                            ip_row    = lv_row
                            ip_value  = space
                            ip_style  = ls_mapper ).
    lv_col = lv_col + 1.

    lo_column = lo_worksheet->get_column( ip_column = lv_col_str ).
    lo_column->set_width( ip_width = 2 ).
  ENDLOOP.

  lo_worksheet->set_show_gridlines( i_show_gridlines = abap_false ).

  lo_worksheet->set_cell( ip_column = 'AP'
                          ip_row    = 15
                          ip_value  = 'Created with abap2xlsx'
                          ip_style  = lv_style_credit_guid ).

  lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = 'http://www.plinky.it/abap/abap2xlsx.php' ).
  lo_worksheet->set_cell( ip_column    = 'AP'
                          ip_row       = 24
                          ip_value     = 'http://www.plinky.it/abap/abap2xlsx.php'
                          ip_style     = lv_style_link_guid
                          ip_hyperlink = lo_hyperlink ).

  lo_column = lo_worksheet->get_column( ip_column = 'AP' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_worksheet->set_merge( ip_row = 15 ip_column_start = 'AP' ip_row_to = 22 ip_column_end = 'AR' ).
  lo_worksheet->set_merge( ip_row = 24 ip_column_start = 'AP' ip_row_to = 26 ip_column_end = 'AR' ).

  CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
  lv_file = lo_excel_writer->write_file( lo_excel ).

  " Convert to binary
  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
    EXPORTING
      buffer        = lv_file
    IMPORTING
      output_length = lv_bytecount
    TABLES
      binary_tab    = lt_file_tab.
*  " This method is only available on AS ABAP > 6.40
*  lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
*  lv_bytecount = xstrlen( lv_file ).

  " Save the file
  cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           CHANGING data_tab     = lt_file_tab ).
