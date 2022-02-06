*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL29
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel29.

DATA: lo_excel        TYPE REF TO zcl_excel,
      lo_excel_writer TYPE REF TO zif_excel_writer,
      lo_excel_reader TYPE REF TO zif_excel_reader.

DATA: lv_file      TYPE xstring,
      lv_bytecount TYPE i,
      lt_file_tab  TYPE solix_tab.

DATA: lv_full_path TYPE string,
      lv_filename  TYPE string,
      lv_workdir   TYPE string.
DATA: lv_separator TYPE c LENGTH 1.

SELECTION-SCREEN COMMENT /1(83) p_text1.
SELECTION-SCREEN COMMENT /1(83) p_text2.
SELECTION-SCREEN SKIP 1.

PARAMETERS: p_smw0 RADIOBUTTON GROUP rad1 DEFAULT 'X'.
PARAMETERS: p_objid TYPE w3objid OBLIGATORY DEFAULT 'ZDEMO_EXCEL29_INPUT'.

PARAMETERS: p_file RADIOBUTTON GROUP rad1.
PARAMETERS: p_path TYPE zexcel_export_dir.

LOAD-OF-PROGRAM.
  p_text1 = 'abap2xlsx works with VBA macro by using an existing VBA binary.'.
  p_text2 = '(we do not want to create a VBA editor).'.

AT SELECTION-SCREEN OUTPUT.
  IF p_path IS INITIAL.
    cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
    cl_gui_cfw=>flush( ).
    cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_separator ).
    p_path = lv_workdir && lv_separator && 'TestMacro.xlsm'.
  ENDIF.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.

  DATA: lt_filetable TYPE filetable,
        lv_rc        TYPE i.

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
* load template
  IF p_file IS NOT INITIAL.
    lo_excel = lo_excel_reader->load_file( lv_full_path ).
  ELSE.
    PERFORM load_smw0 USING lo_excel_reader p_objid CHANGING lo_excel.
  ENDIF.
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

FORM load_smw0
    USING
        io_reader  TYPE REF TO zif_excel_reader
        iv_w3objid TYPE w3objid
    CHANGING
        ro_excel   TYPE REF TO zcl_excel
    RAISING
        zcx_excel.

  DATA: lv_excel_data   TYPE xstring,
        lt_mime         TYPE TABLE OF w3mime,
        ls_key          TYPE wwwdatatab,
        lv_errormessage TYPE string,
        lv_filesize     TYPE i,
        lv_filesizec    TYPE c LENGTH 10.

*--------------------------------------------------------------------*
* Read file into binary string
*--------------------------------------------------------------------*

  ls_key-relid = 'MI'.
  ls_key-objid = iv_w3objid .

  CALL FUNCTION 'WWWDATA_IMPORT'
    EXPORTING
      key    = ls_key
    TABLES
      mime   = lt_mime
    EXCEPTIONS
      OTHERS = 1.
  IF sy-subrc <> 0.
    lv_errormessage = 'A problem occured when reading the MIME object'(004).
    zcx_excel=>raise_text( lv_errormessage ).
  ENDIF.

  CALL FUNCTION 'WWWPARAMS_READ'
    EXPORTING
      relid = ls_key-relid
      objid = ls_key-objid
      name  = 'filesize'
    IMPORTING
      value = lv_filesizec.

  lv_filesize = lv_filesizec.
  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lv_filesize
    IMPORTING
      buffer       = lv_excel_data
    TABLES
      binary_tab   = lt_mime
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.

*--------------------------------------------------------------------*
* Parse Excel data into ZCL_EXCEL object from binary string
*--------------------------------------------------------------------*
  ro_excel = io_reader->load( i_excel2007 = lv_excel_data ).

ENDFORM.
