*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL28
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel28.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_hyperlink            TYPE REF TO zcl_excel_hyperlink,
      lo_column               TYPE REF TO zcl_excel_column.

DATA: lv_file                 TYPE xstring,
      lv_bytecount            TYPE i,
      lt_file_tab             TYPE solix_tab.

DATA: lv_file_name      TYPE string,
      lv_file_path      TYPE string,
      lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

CONSTANTS: lv_default_file_name TYPE string VALUE '28_HelloWorld.csv'.

PARAMETERS: p_path TYPE string.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.

  cl_gui_frontend_services=>directory_browse( EXPORTING initial_folder  = p_path
                                               CHANGING selected_folder = p_path ).

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

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = sy-datum ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = sy-uzeit ).

  lo_column = lo_worksheet->get_column( 'B' ).
  lo_column->set_width( 11 ).

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the second sheet' ).

  CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_csv.
  zcl_excel_writer_csv=>set_delimiter( ip_value = cl_abap_char_utilities=>horizontal_tab ).
  zcl_excel_writer_csv=>set_enclosure( ip_value = '''' ).
  zcl_excel_writer_csv=>set_endofline( ip_value = cl_abap_char_utilities=>cr_lf ).

  zcl_excel_writer_csv=>set_active_sheet_index( i_active_worksheet = 2 ).
*  zcl_excel_writer_csv=>set_active_sheet_index_by_name(  I_WORKSHEET_NAME = 'Sheet2' ).

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
  REPLACE FIRST OCCURRENCE OF '.csv'  IN lv_full_path WITH '_Sheet2.csv'.
  cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           CHANGING data_tab     = lt_file_tab ).

*  zcl_excel_writer_csv=>set_active_sheet_index( i_active_worksheet = 2 ).
  zcl_excel_writer_csv=>set_active_sheet_index_by_name(  i_worksheet_name = 'Sheet1' ).
  lv_file = lo_excel_writer->write_file( lo_excel ).
  REPLACE FIRST OCCURRENCE OF '_Sheet2.csv'  IN lv_full_path WITH '_Sheet1.csv'.

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
