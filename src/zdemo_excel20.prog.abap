*--------------------------------------------------------------------*
* REPORT  ZDEMO_EXCEL20
* Demo for method zcl_excel_worksheet-bind_alv:
* export data from ALV (CL_GUI_ALV_GRID) object to excel
*--------------------------------------------------------------------*
REPORT  zdemo_excel20.

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
      lo_alv            TYPE REF TO cl_gui_alv_grid,
      lo_salv           TYPE REF TO cl_salv_table,
      gr_events         TYPE REF TO lcl_handle_events,
      lr_events         TYPE REF TO cl_salv_events_table,
      gt_sbook          TYPE TABLE OF sbook,
      gt_listheader     TYPE slis_t_listheader,
      wa_listheader     LIKE LINE OF gt_listheader.

DATA: l_path            TYPE string,  " local dir
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

CONSTANTS:
      lv_default_file_name TYPE string VALUE '20_BindAlv.xlsx'.
*--------------------------------------------------------------------*
*START-OF-SELECTION
*--------------------------------------------------------------------*

START-OF-SELECTION.

* get data
* ------------------------------------------

  SELECT *
      INTO TABLE gt_sbook[]
      FROM sbook                                        "#EC CI_NOWHERE
      UP TO 10 ROWS.

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
  IF sy-ucomm = 'EXCEL'.

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

    CONCATENATE l_path lv_file_separator lv_default_file_name
                INTO l_path.

* export file to save file path

    PERFORM export_to_excel.

  ENDIF.
ENDFORM.                    " USER_COMMAND

*--------------------------------------------------------------------*
* FORM EXPORT_TO_EXCEL
*--------------------------------------------------------------------*
* This subroutine is principal demo session
*--------------------------------------------------------------------*
FORM export_to_excel.

* create zcl_excel_worksheet object

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->get_active_worksheet( ).

* get ALV object from screen

  CALL FUNCTION 'GET_GLOBALS_FROM_SLVC_FULLSCR'
    IMPORTING
      e_grid = lo_alv.

* build list header

  wa_listheader-typ = 'H'.
  wa_listheader-info = sy-title.
  APPEND wa_listheader TO gt_listheader.

  wa_listheader-typ = 'S'.
  wa_listheader-info =  'Created by: ABAP2XLSX Group'.
  APPEND wa_listheader TO gt_listheader.

  wa_listheader-typ = 'A'.
  wa_listheader-info =
      'Project hosting at https://cw.sdn.sap.com/cw/groups/abap2xlsx'.
  APPEND wa_listheader TO gt_listheader.

* write to excel using method Bin_ALV

  lo_worksheet->bind_alv_ole2(
    EXPORTING
*      I_DOCUMENT_URL          = SPACE " excel template
*      I_XLS                   = 'X' " create in xls format?
      i_save_path             = l_path
      io_alv                  = lo_alv
    it_listheader           = gt_listheader
      i_top                   = 2
      i_left                  = 1
*      I_COLUMNS_HEADER        = 'X'
*      I_COLUMNS_AUTOFIT       = 'X'
*      I_FORMAT_COL_HEADER     =
*      I_FORMAT_SUBTOTAL       =
*      I_FORMAT_TOTAL          =
    EXCEPTIONS
      miss_guide              = 1
      ex_transfer_kkblo_error = 2
      fatal_error             = 3
      inv_data_range          = 4
      dim_mismatch_vkey       = 5
      dim_mismatch_sema       = 6
      error_in_sema           = 7
      OTHERS                  = 8
          ).
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

ENDFORM.                    "EXPORT_TO_EXCEL
