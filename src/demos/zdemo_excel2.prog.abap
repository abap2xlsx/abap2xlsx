*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL2
*& Test Styles for ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel2.

DATA: lo_excel                  TYPE REF TO zcl_excel,
      lo_worksheet              TYPE REF TO zcl_excel_worksheet,
      lo_style_bold             TYPE REF TO zcl_excel_style,
      lo_style_underline        TYPE REF TO zcl_excel_style,
      lo_style_filled           TYPE REF TO zcl_excel_style,
      lo_style_filled_green     TYPE REF TO zcl_excel_style,
      lo_style_filled_turquoise TYPE REF TO zcl_excel_style,
      lo_style_border           TYPE REF TO zcl_excel_style,
      lo_style_button           TYPE REF TO zcl_excel_style,
      lo_border_dark            TYPE REF TO zcl_excel_style_border,
      lo_border_light           TYPE REF TO zcl_excel_style_border,
      lo_style_gr_cornerlb      TYPE REF TO zcl_excel_style,
      lo_style_gr_cornerlt      TYPE REF TO zcl_excel_style,
      lo_style_gr_cornerrb      TYPE REF TO zcl_excel_style,
      lo_style_gr_cornerrt      TYPE REF TO zcl_excel_style,
      lo_style_gr_horizontal90  TYPE REF TO zcl_excel_style,
      lo_style_gr_horizontal270 TYPE REF TO zcl_excel_style,
      lo_style_gr_horizontalb   TYPE REF TO zcl_excel_style,
      lo_style_gr_vertical      TYPE REF TO zcl_excel_style,
      lo_style_gr_vertical2     TYPE REF TO zcl_excel_style,
      lo_style_gr_fromcenter    TYPE REF TO zcl_excel_style,
      lo_style_gr_diagonal45    TYPE REF TO zcl_excel_style,
      lo_style_gr_diagonal45b   TYPE REF TO zcl_excel_style,
      lo_style_gr_diagonal135   TYPE REF TO zcl_excel_style,
      lo_style_gr_diagonal135b  TYPE REF TO zcl_excel_style.

DATA: lv_file      TYPE xstring,
      lv_bytecount TYPE i,
      lt_file_tab  TYPE solix_tab.

DATA: lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.
DATA: lo_row TYPE REF TO zcl_excel_row.

CONSTANTS: gc_save_file_name TYPE string VALUE '02_Styles.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.



START-OF-SELECTION.


  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Create border object
  CREATE OBJECT lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.
  CREATE OBJECT lo_border_light.
  lo_border_light->border_color-rgb = zcl_excel_style_color=>c_gray.
  lo_border_light->border_style = zcl_excel_style_border=>c_border_thin.
  " Create a bold / italic style
  lo_style_bold               = lo_excel->add_new_style( ).
  lo_style_bold->font->bold   = abap_true.
  lo_style_bold->font->italic = abap_true.
  lo_style_bold->font->name   = zcl_excel_style_font=>c_name_arial.
  lo_style_bold->font->scheme = zcl_excel_style_font=>c_scheme_none.
  lo_style_bold->font->color-rgb  = zcl_excel_style_color=>c_red.
  " Create an underline double style
  lo_style_underline                        = lo_excel->add_new_style( ).
  lo_style_underline->font->underline       = abap_true.
  lo_style_underline->font->underline_mode  = zcl_excel_style_font=>c_underline_double.
  lo_style_underline->font->name            = zcl_excel_style_font=>c_name_roman.
  lo_style_underline->font->scheme          = zcl_excel_style_font=>c_scheme_none.
  lo_style_underline->font->family          = zcl_excel_style_font=>c_family_roman.
  " Create filled style yellow
  lo_style_filled                 = lo_excel->add_new_style( ).
  lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style_filled->fill->fgcolor-theme  = zcl_excel_style_color=>c_theme_accent6.
  " Create border with button effects
  lo_style_button                   = lo_excel->add_new_style( ).
  lo_style_button->borders->right   = lo_border_dark.
  lo_style_button->borders->down    = lo_border_dark.
  lo_style_button->borders->left    = lo_border_light.
  lo_style_button->borders->top     = lo_border_light.
  "Create style with border
  lo_style_border                         = lo_excel->add_new_style( ).
  lo_style_border->borders->allborders    = lo_border_dark.
  lo_style_border->borders->diagonal      = lo_border_dark.
  lo_style_border->borders->diagonal_mode = zcl_excel_style_borders=>c_diagonal_both.
  " Create filled style green
  lo_style_filled_green                     = lo_excel->add_new_style( ).
  lo_style_filled_green->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_filled_green->fill->fgcolor-rgb  = zcl_excel_style_color=>c_green.
  lo_style_filled_green->font->name         = zcl_excel_style_font=>c_name_cambria.
  lo_style_filled_green->font->scheme       = zcl_excel_style_font=>c_scheme_major.

  " Create filled with gradients
  lo_style_gr_cornerlb                     = lo_excel->add_new_style(  ).
  lo_style_gr_cornerlb->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerlb.
  lo_style_gr_cornerlb->fill->fgcolor-rgb  = zcl_excel_style_color=>c_blue.
  lo_style_gr_cornerlb->fill->bgcolor-rgb  = zcl_excel_style_color=>c_white.
  lo_style_gr_cornerlb->font->name         = zcl_excel_style_font=>c_name_cambria.
  lo_style_gr_cornerlb->font->scheme       = zcl_excel_style_font=>c_scheme_major.

  lo_style_gr_cornerlt                     = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_cornerlt->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerlt.

  lo_style_gr_cornerrb                     = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_cornerrb->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerrb.

  lo_style_gr_cornerrt                     = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_cornerrt->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerrt.

  lo_style_gr_horizontal90                 = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_horizontal90->fill->filltype = zcl_excel_style_fill=>c_fill_gradient_horizontal90.

  lo_style_gr_horizontal270                = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_horizontal270->fill->filltype = zcl_excel_style_fill=>c_fill_gradient_horizontal270.

  lo_style_gr_horizontalb                  = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_horizontalb->fill->filltype  = zcl_excel_style_fill=>c_fill_gradient_horizontalb.

  lo_style_gr_vertical                     = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_vertical->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_vertical.

  lo_style_gr_vertical2                    = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_vertical2->fill->filltype    = zcl_excel_style_fill=>c_fill_gradient_vertical.
  lo_style_gr_vertical2->fill->fgcolor-rgb = zcl_excel_style_color=>c_white.
  lo_style_gr_vertical2->fill->bgcolor-rgb = zcl_excel_style_color=>c_blue.

  lo_style_gr_fromcenter                   = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_fromcenter->fill->filltype   = zcl_excel_style_fill=>c_fill_gradient_fromcenter.

  lo_style_gr_diagonal45                   = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_diagonal45->fill->filltype   = zcl_excel_style_fill=>c_fill_gradient_diagonal45.

  lo_style_gr_diagonal45b                  = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_diagonal45b->fill->filltype  = zcl_excel_style_fill=>c_fill_gradient_diagonal45b.

  lo_style_gr_diagonal135                  = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_diagonal135->fill->filltype  = zcl_excel_style_fill=>c_fill_gradient_diagonal135.

  lo_style_gr_diagonal135b                 = lo_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
  lo_style_gr_diagonal135b->fill->filltype = zcl_excel_style_fill=>c_fill_gradient_diagonal135b.



  " Create filled style turquoise using legacy excel ver <= 2003 palette. (https://github.com/abap2xlsx/abap2xlsx/issues/92)
  lo_style_filled_turquoise                 = lo_excel->add_new_style( ).
  lo_excel->legacy_palette->set_color( "replace built-in color from palette with out custom RGB turquoise
      ip_index =     16
      ip_color =     '0040E0D0' ).

  lo_style_filled_turquoise->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style_filled_turquoise->fill->fgcolor-indexed  = 16.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Styles' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Bold text'            ip_style = lo_style_bold ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 4 ip_value = 'Underlined text'      ip_style = lo_style_underline ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 5 ip_value = 'Filled text'          ip_style = lo_style_filled ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 6 ip_value = 'Borders'              ip_style = lo_style_border ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 7 ip_value = 'I''m not a button :)' ip_style = lo_style_button ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 9 ip_value = 'Modified color for Excel 2003' ip_style = lo_style_filled_turquoise ).
  " Fill the cell and apply one style
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 6 ip_value = 'Filled text'          ip_style = lo_style_filled ).
  " Change the style
  lo_worksheet->set_cell_style( ip_column = 'B' ip_row = 6 ip_style = lo_style_filled_green ).
  " Add Style to an empty cell to test Fix for Issue
  "#44 Exception ZCX_EXCEL thrown when style is set for an empty cell
  " https://github.com/abap2xlsx/abap2xlsx/issues/44
  lo_worksheet->set_cell_style( ip_column = 'E' ip_row = 6 ip_style = lo_style_filled_green ).


  lo_worksheet->set_cell( ip_column = 'B' ip_row = 10  ip_style = lo_style_gr_cornerlb ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerlb ).
  lo_row = lo_worksheet->get_row( ip_row = 10 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 11  ip_style = lo_style_gr_cornerlt ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerlt ).
  lo_row = lo_worksheet->get_row( ip_row = 11 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 12  ip_style = lo_style_gr_cornerrb ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerrb ).
  lo_row = lo_worksheet->get_row( ip_row = 12 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 13  ip_style = lo_style_gr_cornerrt ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerrt ).
  lo_row = lo_worksheet->get_row( ip_row = 13 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 14  ip_style = lo_style_gr_horizontal90 ip_value = zcl_excel_style_fill=>c_fill_gradient_horizontal90 ).
  lo_row = lo_worksheet->get_row( ip_row = 14 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 15  ip_style = lo_style_gr_horizontal270 ip_value = zcl_excel_style_fill=>c_fill_gradient_horizontal270 ).
  lo_row = lo_worksheet->get_row( ip_row = 15 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 16  ip_style = lo_style_gr_horizontalb ip_value = zcl_excel_style_fill=>c_fill_gradient_horizontalb ).
  lo_row = lo_worksheet->get_row( ip_row = 16 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 17  ip_style = lo_style_gr_vertical ip_value = zcl_excel_style_fill=>c_fill_gradient_vertical ).
  lo_row = lo_worksheet->get_row( ip_row = 17 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 18  ip_style = lo_style_gr_vertical2 ip_value = zcl_excel_style_fill=>c_fill_gradient_vertical ).
  lo_row = lo_worksheet->get_row( ip_row = 18 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 19  ip_style = lo_style_gr_fromcenter ip_value = zcl_excel_style_fill=>c_fill_gradient_fromcenter ).
  lo_row = lo_worksheet->get_row( ip_row = 19 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 20  ip_style = lo_style_gr_diagonal45 ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal45 ).
  lo_row = lo_worksheet->get_row( ip_row = 20 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 21  ip_style = lo_style_gr_diagonal45b ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal45b ).
  lo_row = lo_worksheet->get_row( ip_row = 21 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 22  ip_style = lo_style_gr_diagonal135 ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal135 ).
  lo_row = lo_worksheet->get_row( ip_row = 22 ).
  lo_row->set_row_height( ip_row_height = 30 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 23  ip_style = lo_style_gr_diagonal135b ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal135b ).
  lo_row = lo_worksheet->get_row( ip_row = 23 ).
  lo_row->set_row_height( ip_row_height = 30 ).


  lcl_output=>output( lo_excel ).
