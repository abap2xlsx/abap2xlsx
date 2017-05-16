*&---------------------------------------------------------------------*
*& Report  ZDEMO_TECHED2
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_teched2.

*******************************
*   Data Object declaration   *
*******************************

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet.

DATA: lo_style_title           TYPE REF TO zcl_excel_style,
      lv_style_title_guid      TYPE zexcel_cell_style.

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
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Demo TechEd' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 5 ip_value = 'TechEd demo' ip_style = lv_style_title_guid ).

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
