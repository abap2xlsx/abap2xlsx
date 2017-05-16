*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL15
*&
*&---------------------------------------------------------------------*
*& 2010-10-30, Gregor Wolf:
*& Added the functionality to ouput the read table content
*& 2011-12-19, Shahrin Shahrulzaman:
*& Added the functionality to have multiple input and output files
*&---------------------------------------------------------------------*

REPORT  zdemo_excel15.

TYPE-POOLS: abap.

TYPES:
  BEGIN OF t_demo_excel15,
    input TYPE string,
  END OF t_demo_excel15.

DATA: excel           TYPE REF TO zcl_excel,
      lo_excel_writer TYPE REF TO zif_excel_writer,
      reader          TYPE REF TO zif_excel_reader.

DATA: ex  TYPE REF TO zcx_excel,
      msg TYPE string.

DATA: lv_file                 TYPE xstring,
      lv_bytecount            TYPE i,
      lt_file_tab             TYPE solix_tab.

DATA: lv_workdir        TYPE string,
      output_file_path  TYPE string,
      input_file_path   TYPE string,
      lv_file_separator TYPE c.

DATA: worksheet      TYPE REF TO zcl_excel_worksheet,
      highest_column TYPE zexcel_cell_column,
      highest_row    TYPE int4,
      column         TYPE zexcel_cell_column VALUE 1,
      col_str        TYPE zexcel_cell_column_alpha,
      row            TYPE int4               VALUE 1,
      value          TYPE zexcel_cell_value.

DATA:
      lt_files       TYPE TABLE OF t_demo_excel15.
FIELD-SYMBOLS: <wa_files> TYPE t_demo_excel15.

PARAMETERS: p_path  TYPE zexcel_export_dir,
            p_noout TYPE xfeld DEFAULT abap_true.


AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( EXPORTING initial_folder  = lv_workdir
                                              CHANGING  selected_folder = lv_workdir ).
  p_path = lv_workdir.

INITIALIZATION.
  cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

  APPEND INITIAL LINE TO lt_files ASSIGNING <wa_files>.
  <wa_files>-input  = '01_HelloWorld.xlsx'.
  APPEND INITIAL LINE TO lt_files ASSIGNING <wa_files>.
  <wa_files>-input  = '02_Styles.xlsx'.
  APPEND INITIAL LINE TO lt_files ASSIGNING <wa_files>.
  <wa_files>-input  = '03_iTab.xlsx'.
  APPEND INITIAL LINE TO lt_files ASSIGNING <wa_files>.
  <wa_files>-input  = '04_Sheets.xlsx'.
  APPEND INITIAL LINE TO lt_files ASSIGNING <wa_files>.
  <wa_files>-input  = '08_Range.xlsx'.
  APPEND INITIAL LINE TO lt_files ASSIGNING <wa_files>.
  <wa_files>-input  = '13_MergedCells.xlsx'.
  APPEND INITIAL LINE TO lt_files ASSIGNING <wa_files>.
  <wa_files>-input  = '31_AutosizeWithDifferentFontSizes.xlsx'.

START-OF-SELECTION.

  IF p_path IS INITIAL.
    p_path = lv_workdir.
  ENDIF.
  cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_file_separator ).

  LOOP AT lt_files ASSIGNING <wa_files>.
    CONCATENATE p_path lv_file_separator <wa_files>-input INTO input_file_path.
    CONCATENATE p_path lv_file_separator '15_' <wa_files>-input INTO output_file_path.
    REPLACE '.xlsx' IN output_file_path WITH 'FromReader.xlsx'.

    TRY.
        CREATE OBJECT reader TYPE zcl_excel_reader_2007.
        excel = reader->load_file( input_file_path ).

        IF p_noout EQ abap_false.
          worksheet = excel->get_active_worksheet( ).
          highest_column = worksheet->get_highest_column( ).
          highest_row    = worksheet->get_highest_row( ).

          WRITE: 'Highest column: ', highest_column, 'Highest row: ', highest_row.
          WRITE: /.

          WHILE row <= highest_row.
            WHILE column <= highest_column.
              col_str = zcl_excel_common=>convert_column2alpha( column ).
              worksheet->get_cell(
                EXPORTING
                  ip_column = col_str
                  ip_row    = row
                IMPORTING
                  ep_value = value
              ).
              WRITE: value.
              column = column + 1.
            ENDWHILE.
            WRITE: /.
            column = 1.
            row = row + 1.
          ENDWHILE.
        ENDIF.
        CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
        lv_file = lo_excel_writer->write_file( excel ).

        " Convert to binary
        CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
          EXPORTING
            buffer        = lv_file
          IMPORTING
            output_length = lv_bytecount
          TABLES
            binary_tab    = lt_file_tab.
*    " This method is only available on AS ABAP > 6.40
*    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
*    lv_bytecount = xstrlen( lv_file ).

        " Save the file
        cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                          filename     = output_file_path
                                                          filetype     = 'BIN'
                                                 CHANGING data_tab     = lt_file_tab ).


      CATCH zcx_excel INTO ex.    " Exceptions for ABAP2XLSX
        msg = ex->get_text( ).
        WRITE: / msg.
    ENDTRY.
  ENDLOOP.
