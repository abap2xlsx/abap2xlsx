*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL26
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel29.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      lo_excel_reader         TYPE REF TO zif_excel_reader.

DATA: lv_file                 TYPE xstring,
      lv_bytecount            TYPE i,
      lt_file_tab             TYPE solix_tab.

DATA: lv_full_path      TYPE string,
      lv_filename       TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

PARAMETERS: p_path TYPE zexcel_export_dir OBLIGATORY.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.

  DATA: lt_filetable TYPE filetable,
        lv_rc TYPE i.

  cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = 'Select Macro-Enabled Workbook template'
      default_extension       = '*.xlsm'
      file_filter             = 'Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm'
      initial_directory       = lv_workdir
    CHANGING
      file_table              = lt_filetable
      rc                      = lv_rc
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
  READ TABLE lt_filetable INTO lv_filename INDEX 1.
  p_path = lv_filename.

START-OF-SELECTION.

  lv_full_path = p_path.

  CREATE OBJECT lo_excel_reader TYPE zcl_excel_reader_xlsm.
  CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_xlsm.
  lo_excel = lo_excel_reader->load_file( lv_full_path ).
  lv_file = lo_excel_writer->write_file( lo_excel ).
  REPLACE '.xlsm' IN lv_full_path WITH 'FromReader.xlsm'.

  " Convert to binary
  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
    EXPORTING
      buffer        = lv_file
    IMPORTING
      output_length = lv_bytecount
    TABLES
      binary_tab    = lt_file_tab.

  " Save the file
  cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           CHANGING data_tab     = lt_file_tab ).
