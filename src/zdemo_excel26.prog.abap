*--------------------------------------------------------------------*
* REPORT  ZDEMO_EXCEL26
* Demo for method zcl_excel_worksheet-bind_object:
* export data from ALV (CL_GUI_ALV_GRID) object or cl_salv_table object
* to Excel.
*--------------------------------------------------------------------*
report  zdemo_excel26.

*----------------------------------------------------------------------*
*       CLASS lcl_handle_events DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class lcl_handle_events definition.
  public section.
    methods:
      on_user_command for event added_function of cl_salv_events
        importing e_salv_function.
endclass.                    "lcl_handle_events DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_handle_events IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class lcl_handle_events implementation.
  method on_user_command.
    perform user_command." using e_salv_function text-i08.
  endmethod.                    "on_user_command
endclass.                     "lcl_handle_events IMPLEMENTATION

*--------------------------------------------------------------------*
* DATA DECLARATION
*--------------------------------------------------------------------*

data: lo_excel          type ref to zcl_excel,
      lo_worksheet      type ref to zcl_excel_worksheet,
      lo_salv           type ref to cl_salv_table,
      gr_events         type ref to lcl_handle_events,
      lr_events         type ref to cl_salv_events_table,
      gt_sbook          type table of sbook.

data: l_path            type string,  " local dir
      lv_workdir        type string,
      lv_file_separator type c.

constants:
      lv_default_file_name type string value '26_Bind_ALV.xlsx'.
*--------------------------------------------------------------------*
*START-OF-SELECTION
*--------------------------------------------------------------------*

start-of-selection.

* get data
* ------------------------------------------

  select *
      into table gt_sbook[]
      from sbook                                        "#EC CI_NOWHERE
      up to 10 rows.

* Display ALV
* ------------------------------------------

  try.
      cl_salv_table=>factory(
        exporting
          list_display = abap_false
        importing
          r_salv_table = lo_salv
        changing
          t_table      = gt_sbook[] ).
    catch cx_salv_msg .
  endtry.

  try.
      lo_salv->set_screen_status(
        exporting
          report        = sy-repid
          pfstatus      = 'ALV_STATUS'
          set_functions = lo_salv->c_functions_all ).
    catch cx_salv_msg .
  endtry.

  lr_events = lo_salv->get_event( ).
  create object gr_events.
  set handler gr_events->on_user_command for lr_events.

  lo_salv->display( ).


*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND
*&---------------------------------------------------------------------*
*      ALV user command
*--------------------------------------------------------------------*
form user_command .
  if sy-ucomm = 'EXCEL'.

* get save file path
    cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = l_path ).
    cl_gui_cfw=>flush( ).
    cl_gui_frontend_services=>directory_browse(
      exporting initial_folder = l_path
      changing selected_folder = l_path ).

    if l_path is initial.
      cl_gui_frontend_services=>get_sapgui_workdir(
        changing sapworkdir = lv_workdir ).
      l_path = lv_workdir.
    endif.

    cl_gui_frontend_services=>get_file_separator(
      changing file_separator = lv_file_separator ).

    concatenate l_path lv_file_separator lv_default_file_name
                into l_path.

* export file to save file path
    perform export_to_excel.

  endif.
endform.                    " USER_COMMAND

*--------------------------------------------------------------------*
* FORM EXPORT_TO_EXCEL
*--------------------------------------------------------------------*
* This subroutine is principal demo session
*--------------------------------------------------------------------*
form export_to_excel.
* create zcl_excel_worksheet object

  create object lo_excel.
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet1' ).

* write to excel using method Bin_object
  try.
      lo_worksheet->bind_alv(
          io_alv      = lo_salv
          it_table    = gt_sbook
          i_top       = 2
          i_left      = 1
             ).
    catch zcx_excel .
  endtry.

  perform write_file.

endform.                    "EXPORT_TO_EXCEL
*&---------------------------------------------------------------------*
*&      Form  WRITE_FILE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
form write_file .
  data: lt_file	    type solix_tab,
        l_bytecount	type i,
        l_file      type xstring.

  data: lo_excel_writer         type ref to zif_excel_writer.

  data: ls_seoclass type seoclass.

  create object lo_excel_writer type zcl_excel_writer_2007.
  l_file = lo_excel_writer->write_file( lo_excel ).

  select single * into ls_seoclass
    from seoclass
    where clsname = 'CL_BCS_CONVERT'.

  if sy-subrc = 0.
    call method (ls_seoclass-clsname)=>xstring_to_solix
      exporting
        iv_xstring = l_file
      receiving
        et_solix   = lt_file.

    l_bytecount = xstrlen( l_file ).
  else.
    " Convert to binary
    call function 'SCMS_XSTRING_TO_BINARY'
      exporting
        buffer        = l_file
      importing
        output_length = l_bytecount
      tables
        binary_tab    = lt_file.
  endif.

  cl_gui_frontend_services=>gui_download( exporting bin_filesize = l_bytecount
                                                    filename     = l_path
                                                    filetype     = 'BIN'
                                           changing data_tab     = lt_file ).

endform.                    " WRITE_FILE
