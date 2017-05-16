*--------------------------------------------------------------------*
* REPORT  ZDEMO_EXCEL32
* Demo for export options from ALV GRID:
* export data from ALV (CL_GUI_ALV_GRID) object or cl_salv_table object
* to Excel.
*--------------------------------------------------------------------*
REPORT  zdemo_excel32.

*----------------------------------------------------------------------*
*       CLASS lcl_handle_events DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_handle_events DEFINITION.
  PUBLIC SECTION.
    METHODS:
      on_user_command FOR EVENT added_function OF cl_salv_events
        IMPORTING e_salv_function.
ENDCLASS.                    "lcl_handle_events DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_handle_events IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_handle_events IMPLEMENTATION.
  METHOD on_user_command.
    PERFORM user_command." using e_salv_function text-i08.
  ENDMETHOD.                    "on_user_command
ENDCLASS.                     "lcl_handle_events IMPLEMENTATION

*--------------------------------------------------------------------*
* DATA DECLARATION
*--------------------------------------------------------------------*

DATA: lo_excel          TYPE REF TO zcl_excel,
      lo_worksheet      TYPE REF TO zcl_excel_worksheet,
      lo_salv           TYPE REF TO cl_salv_table,
      gr_events         TYPE REF TO lcl_handle_events,
      lr_events         TYPE REF TO cl_salv_events_table,
      gt_sbook          TYPE TABLE OF sbook.

DATA: l_path            TYPE string,  " local dir
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

CONSTANTS:
      lv_default_file_name TYPE string VALUE '32_Export_ALV.xlsx',
      lv_default_file_name2 TYPE string VALUE '32_Export_Convert.xlsx'.
*--------------------------------------------------------------------*
*START-OF-SELECTION
*--------------------------------------------------------------------*

START-OF-SELECTION.

* get data
* ------------------------------------------

  SELECT *
      INTO TABLE gt_sbook[]
      FROM sbook                                        "#EC CI_NOWHERE
      UP TO 100 ROWS.

* Display ALV
* ------------------------------------------

  TRY.
      cl_salv_table=>factory(
        EXPORTING
          list_display = abap_false
        IMPORTING
          r_salv_table = lo_salv
        CHANGING
          t_table      = gt_sbook[] ).
    CATCH cx_salv_msg .
  ENDTRY.

  TRY.
      lo_salv->set_screen_status(
        EXPORTING
          report        = sy-repid
          pfstatus      = 'ALV_STATUS'
          set_functions = lo_salv->c_functions_all ).
    CATCH cx_salv_msg .
  ENDTRY.

  lr_events = lo_salv->get_event( ).
  CREATE OBJECT gr_events.
  SET HANDLER gr_events->on_user_command FOR lr_events.

  lo_salv->display( ).


*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND
*&---------------------------------------------------------------------*
*      ALV user command
*--------------------------------------------------------------------*
FORM user_command .

* get save file path
      cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = l_path ).
      cl_gui_cfw=>flush( ).
      cl_gui_frontend_services=>directory_browse(
        EXPORTING initial_folder = l_path
        CHANGING selected_folder = l_path ).

      IF l_path IS INITIAL.
        cl_gui_frontend_services=>get_sapgui_workdir(
          CHANGING sapworkdir = lv_workdir ).
        l_path = lv_workdir.
      ENDIF.

      cl_gui_frontend_services=>get_file_separator(
        CHANGING file_separator = lv_file_separator ).



* export file to save file path
  CASE sy-ucomm.
    WHEN 'EXCELBIND'.
      CONCATENATE l_path lv_file_separator lv_default_file_name
                  INTO l_path.
      PERFORM export_to_excel_bind.

    WHEN 'EXCELCONV'.

      CONCATENATE l_path lv_file_separator lv_default_file_name2
                  INTO l_path.
      PERFORM export_to_excel_conv.

  ENDCASE.
ENDFORM.                    " USER_COMMAND
*--------------------------------------------------------------------*
* FORM EXPORT_TO_EXCEL_CONV
*--------------------------------------------------------------------*
* This subroutine is principal demo session
*--------------------------------------------------------------------*
FORM export_to_excel_conv.
  DATA: lo_converter TYPE REF TO zcl_excel_converter.

  CREATE OBJECT lo_converter.
*TRY.
  lo_converter->convert(
    EXPORTING
      io_alv        = lo_salv
      it_table      = gt_sbook
      i_row_int     = 2
      i_column_int  = 2
*    i_table       =
*    i_style_table =
*    io_worksheet  =
*  CHANGING
*    co_excel      =
         ).
* CATCH zcx_excel .
*ENDTRY.
  lo_converter->write_file( i_path = l_path ).

ENDFORM.                    "EXPORT_TO_EXCEL_CONV

*--------------------------------------------------------------------*
* FORM EXPORT_TO_EXCEL_BIND
*--------------------------------------------------------------------*
* This subroutine is principal demo session
*--------------------------------------------------------------------*
FORM export_to_excel_bind.
* create zcl_excel_worksheet object
  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet1' ).

* write to excel using method Bin_object
*try.
  lo_worksheet->bind_alv(
      io_alv      = lo_salv
      it_table    = gt_sbook
      i_top       = 2
      i_left      = 1
         ).
* catch zcx_excel .
*endtry.


  PERFORM write_file.

ENDFORM.                    "EXPORT_TO_EXCEL_BIND
*&---------------------------------------------------------------------*
*&      Form  WRITE_FILE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM write_file .
  DATA: lt_file	    TYPE solix_tab,
        l_bytecount	TYPE i,
        l_file      TYPE xstring.

  DATA: lo_excel_writer         TYPE REF TO zif_excel_writer.

  DATA: ls_seoclass TYPE seoclass.

  CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
  l_file = lo_excel_writer->write_file( lo_excel ).

  SELECT SINGLE * INTO ls_seoclass
    FROM seoclass
    WHERE clsname = 'CL_BCS_CONVERT'.

  IF sy-subrc = 0.
    CALL METHOD (ls_seoclass-clsname)=>xstring_to_solix
      EXPORTING
        iv_xstring = l_file
      RECEIVING
        et_solix   = lt_file.

    l_bytecount = XSTRLEN( l_file ).
  ELSE.
    " Convert to binary
    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer        = l_file
      IMPORTING
        output_length = l_bytecount
      TABLES
        binary_tab    = lt_file.
  ENDIF.

  cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = l_bytecount
                                                    filename     = l_path
                                                    filetype     = 'BIN'
                                           CHANGING data_tab     = lt_file ).

ENDFORM.                    " WRITE_FILE
