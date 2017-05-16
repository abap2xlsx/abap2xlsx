*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL2
*& Test Styles for ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel35.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_style_bold           TYPE REF TO zcl_excel_style,
      lo_style_underline      TYPE REF TO zcl_excel_style,
      lo_style_filled         TYPE REF TO zcl_excel_style,
      lo_style_border         TYPE REF TO zcl_excel_style,
      lo_style_button         TYPE REF TO zcl_excel_style,
      lo_border_dark          TYPE REF TO zcl_excel_style_border,
      lo_border_light         TYPE REF TO zcl_excel_style_border.

DATA: lv_style_bold_guid         TYPE zexcel_cell_style,
      lv_style_underline_guid    TYPE zexcel_cell_style,
      lv_style_filled_guid       TYPE zexcel_cell_style,
      lv_style_filled_green_guid TYPE zexcel_cell_style,
      lv_style_border_guid       TYPE zexcel_cell_style,
      lv_style_button_guid       TYPE zexcel_cell_style,
      lv_style_filled_turquoise_guid TYPE zexcel_cell_style.

DATA: lv_file                 TYPE xstring,
      lv_bytecount            TYPE i,
      lt_file_tab             TYPE solix_tab.

DATA: lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

CONSTANTS: lv_default_file_name TYPE string VALUE '35_Static_Styles.xlsx'.

PARAMETERS: p_path TYPE zexcel_export_dir.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( EXPORTING initial_folder  = lv_workdir
                                              CHANGING  selected_folder = lv_workdir ).
  p_path = lv_workdir.

INITIALIZATION.
  cl_gui_frontend_services=>GET_DESKTOP_DIRECTORY( CHANGING DESKTOP_DIRECTORY = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

  sy-title = 'ZDEMO_EXCEL2;Issue 139: Change cellstyle retroactivly'.

START-OF-SELECTION.

  IF p_path IS INITIAL.
    p_path = lv_workdir.
  ENDIF.
  cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_file_separator ).
  CONCATENATE p_path lv_file_separator lv_default_file_name INTO lv_full_path.

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
  lv_style_bold_guid          = lo_style_bold->get_guid( ).
  " Create an underline double style
  lo_style_underline                        = lo_excel->add_new_style( ).
  lo_style_underline->font->underline       = abap_true.
  lo_style_underline->font->underline_mode  = zcl_excel_style_font=>c_underline_double.
  lo_style_underline->font->name            = zcl_excel_style_font=>c_name_roman.
  lo_style_underline->font->scheme          = zcl_excel_style_font=>c_scheme_none.
  lo_style_underline->font->family          = zcl_excel_style_font=>c_family_roman.
  lv_style_underline_guid                   = lo_style_underline->get_guid( ).
  " Create filled style yellow
  lo_style_filled                 = lo_excel->add_new_style( ).
  lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style_filled->fill->fgcolor-theme  = zcl_excel_style_color=>c_theme_accent6.
  lv_style_filled_guid            = lo_style_filled->get_guid( ).
  " Create border with button effects
  lo_style_button                   = lo_excel->add_new_style( ).
  lo_style_button->borders->right   = lo_border_dark.
  lo_style_button->borders->down    = lo_border_dark.
  lo_style_button->borders->left    = lo_border_light.
  lo_style_button->borders->top     = lo_border_light.
  lv_style_button_guid              = lo_style_button->get_guid( ).
  "Create style with border
  lo_style_border                         = lo_excel->add_new_style( ).
  lo_style_border->borders->allborders    = lo_border_dark.
  lo_style_border->borders->diagonal      = lo_border_dark.
  lo_style_border->borders->diagonal_mode = zcl_excel_style_borders=>c_diagonal_both.
  lv_style_border_guid                    = lo_style_border->get_guid( ).
  " Create filled style green
  lo_style_filled                     = lo_excel->add_new_style( ).
  lo_style_filled->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_filled->fill->fgcolor-rgb  = zcl_excel_style_color=>c_green.
  lo_style_filled->font->name         = zcl_excel_style_font=>c_name_cambria.
  lo_style_filled->font->scheme       = zcl_excel_style_font=>c_scheme_major.
  lv_style_filled_green_guid          = lo_style_filled->get_guid( ).

  " Create filled style turquoise using legacy excel ver <= 2003 palette. (https://code.sdn.sap.com/spaces/abap2xlsx/tickets/92)
  lo_style_filled                 = lo_excel->add_new_style( ).
  lo_excel->legacy_palette->set_color( "replace built-in color from palette with out custom RGB turquoise
      ip_index =     16
      ip_color =     '0040E0D0' ).

  lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style_filled->fill->fgcolor-indexed  = 16.
  lv_style_filled_turquoise_guid            = lo_style_filled->get_guid( ).

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Styles' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Bold text'            ip_style = lv_style_bold_guid ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 4 ip_value = 'Underlined text'      ip_style = lv_style_underline_guid ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 5 ip_value = 'Filled text'          ip_style = lv_style_filled_guid ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 6 ip_value = 'Borders'              ip_style = lv_style_border_guid ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 7 ip_value = 'I''m not a button :)' ip_style = lv_style_button_guid ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 9 ip_value = 'Modified color for Excel 2003' ip_style = lv_style_filled_turquoise_guid ).
  " Fill the cell and apply one style
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 6 ip_value = 'Filled text'          ip_style = lv_style_filled_guid ).
  " Change the style
  lo_worksheet->set_cell_style( ip_column = 'B' ip_row = 6 ip_style = lv_style_filled_green_guid ).
  " Add Style to an empty cell to test Fix for Issue
  "#44 Exception ZCX_EXCEL thrown when style is set for an empty cell
  " https://code.sdn.sap.com/spaces/abap2xlsx/tickets/44-exception-zcx_excel-thrown-when-style-is-set-for-an-empty-cell
  lo_worksheet->set_cell_style( ip_column = 'E' ip_row = 6 ip_style = lv_style_filled_green_guid ).


* Demonstrate how to retroactivly change the cellstyle
*Filled text and underlinded text
  lo_worksheet->change_cell_style(  ip_column                     = 'B'
                                    ip_row                        = 5
                                    ip_font_bold                 = abap_true
                                    ip_font_italic               = abap_true ).

  lo_worksheet->change_cell_style(  ip_column                     = 'D'
                                    ip_row                        = 4
                                    ip_font_bold                 = abap_true
                                    ip_font_italic               = abap_true ).

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
