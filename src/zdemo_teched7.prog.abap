*&---------------------------------------------------------------------*
*& Report  ZDEMO_TECHED3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_teched3.

*******************************
*   Data Object declaration   *
*******************************

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet.

DATA: lo_style_title           TYPE REF TO zcl_excel_style,
      lo_drawing               TYPE REF TO zcl_excel_drawing,
      lo_range                 TYPE REF TO zcl_excel_range,
      lo_data_validation       TYPE REF TO zcl_excel_data_validation,
      lo_column                 TYPE REF TO zcl_excel_column,
      lv_style_title_guid      TYPE zexcel_cell_style,
      ls_key                   TYPE wwwdatatab.

DATA: lv_file                 TYPE xstring,
      lv_bytecount            TYPE i,
      lt_file_tab             TYPE solix_tab.

DATA: lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

CONSTANTS: lv_default_file_name TYPE string VALUE 'TechEd01.xlsx'.

*******************************
* Selection screen management *
*******************************

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

*******************************
*    abap2xlsx create XLSX    *
*******************************

  " Create excel instance
  CREATE OBJECT lo_excel.

  " Styles
  lo_style_title                   = lo_excel->add_new_style( ).
  lo_style_title->font->bold       = abap_true.
  lo_style_title->font->color-rgb  = zcl_excel_style_color=>c_blue.
  lv_style_title_guid              = lo_style_title->get_guid( ).

  " Get active sheet
  lo_worksheet        = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Demo TechEd' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 5 ip_value = 'TechEd demo' ip_style = lv_style_title_guid ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 7 ip_value = 'Is abap2xlsx simple' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 8 ip_value = 'Is abap2xlsx CooL' ).

  lo_worksheet->set_cell( ip_column = 'B' ip_row = 10 ip_value = 'Total score' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 10 ip_formula = 'SUM(C7:C8)' ).

  " add logo from SMWO
  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_position( ip_from_row = 2
                            ip_from_col = 'B' ).

  ls_key-relid = 'MI'.
  ls_key-objid = 'WBLOGO'.
  lo_drawing->set_media_www( ip_key = ls_key
                             ip_width = 140
                             ip_height = 64 ).

  " assign drawing to the worksheet
  lo_worksheet->add_drawing( lo_drawing ).

  " Add new sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Values' ).

  " Set values for range
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 1 ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'A' ip_value = 2 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 3 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'A' ip_value = 4 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'A' ip_value = 5 ).

  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = 'Values'.
  lo_range->set_value( ip_sheet_name    = 'Values'
                       ip_start_column  = 'A'
                       ip_start_row     = 4
                       ip_stop_column   = 'A'
                       ip_stop_row      = 8 ).

  lo_excel->set_active_sheet_index( 1 ).

  " add data validation
  lo_worksheet        = lo_excel->get_active_worksheet( ).

  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1    = 'Values'.
  lo_data_validation->cell_row    = 7.
  lo_data_validation->cell_column = 'C'.
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 'Select a value' ).


  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1    = 'Values'.
  lo_data_validation->cell_row    = 8.
  lo_data_validation->cell_column = 'C'.
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 'Select a value' ).

  " add autosize (column width)
  lo_column = lo_worksheet->get_column( ip_column = 'B' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( ip_column = 'C' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).

  " Create xlsx stream
  CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
  lv_file = lo_excel_writer->write_file( lo_excel ).

*******************************
*            Output           *
*******************************

  " Convert to binary
  lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
  lv_bytecount = xstrlen( lv_file ).

  " Save the file
  cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           CHANGING data_tab     = lt_file_tab ).
