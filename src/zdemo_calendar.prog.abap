*&---------------------------------------------------------------------*
*& Report  ZDEMO_CALENDAR
*& abap2xlsx Demo: Create Calendar with Pictures
*&---------------------------------------------------------------------*
*& This report creates a monthly calendar in the specified date range.
*& Each month is put on a seperate worksheet. The pictures for each
*& month can be specified in a tab delimited file called "Calendar.txt"
*& which is saved in the Export Directory. By default this is the SAP
*& Workdir. The file contains 3 fields:
*&
*& Month (with leading 0)
*& Image Filename
*& Image Description
*& URL for the Description
*&
*& The Images should be landscape JPEG's with a 3:2 ratio and min.
*& 450 pixel height. They must also be saved in the Export Directory.
*& In my tests I've discovered a limit of 20 MB in the
*& cl_gui_frontend_services=>gui_download method. So keep your images
*& smaller or change to a server export using OPEN DATASET.
*&---------------------------------------------------------------------*

REPORT  zdemo_calendar.

TYPE-POOLS: abap.
CONSTANTS: gc_save_file_name TYPE string VALUE 'Calendar.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.
INCLUDE zdemo_calendar_classes.

DATA: lv_workdir        TYPE string.

PARAMETERS: p_from    TYPE dfrom DEFAULT '20130101',
            p_to      TYPE dto   DEFAULT '20131231'.

SELECTION-SCREEN BEGIN OF BLOCK orientation WITH FRAME TITLE orient.
PARAMETERS: p_portr   TYPE flag RADIOBUTTON GROUP orie,
            p_lands   TYPE flag RADIOBUTTON GROUP orie DEFAULT 'X'.
SELECTION-SCREEN END OF BLOCK orientation.

INITIALIZATION.
  cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.
  orient = 'Orientation'(000).

START-OF-SELECTION.

  DATA: lo_excel                TYPE REF TO zcl_excel,
        lo_excel_writer         TYPE REF TO zif_excel_writer,
        lo_worksheet            TYPE REF TO zcl_excel_worksheet,
        lo_column               TYPE REF TO zcl_excel_column,
        lo_row                  TYPE REF TO zcl_excel_row,
        lo_hyperlink            TYPE REF TO zcl_excel_hyperlink,
        lo_drawing              TYPE REF TO zcl_excel_drawing.

  DATA: lo_style_month          TYPE REF TO zcl_excel_style,
        lv_style_month_guid     TYPE zexcel_cell_style.
  DATA: lo_style_border         TYPE REF TO zcl_excel_style,
        lo_border_dark          TYPE REF TO zcl_excel_style_border,
        lv_style_border_guid    TYPE zexcel_cell_style.
  DATA: lo_style_center         TYPE REF TO zcl_excel_style,
        lv_style_center_guid    TYPE zexcel_cell_style.

  DATA: lv_file                 TYPE xstring,
        lv_bytecount            TYPE i,
        lt_file_tab             TYPE solix_tab.

  DATA: lv_full_path      TYPE string,
        image_descr_path  TYPE string,
        lv_file_separator TYPE c.
  DATA: lv_content TYPE xstring,
        width      TYPE i,
        lv_height  TYPE i,
        lv_from_row TYPE zexcel_cell_row.

  DATA: month      TYPE i,
        month_nr   TYPE fcmnr,
        count      TYPE i VALUE 1,
        title      TYPE zexcel_sheet_title,
        value      TYPE string,
        image_path TYPE string,
        date_from  TYPE datum,
        date_to    TYPE datum,
        row        TYPE zexcel_cell_row,
        to_row     TYPE zexcel_cell_row,
        to_col     TYPE zexcel_cell_column_alpha,
        to_col_end TYPE zexcel_cell_column_alpha,
        to_col_int TYPE i.

  DATA: month_names TYPE TABLE OF t247.
  FIELD-SYMBOLS: <month_name> LIKE LINE OF month_names.

  TYPES: BEGIN OF tt_datatab,
          month_nr   TYPE fcmnr,
          filename   TYPE string,
          descr      TYPE string,
          url        TYPE string,
        END OF tt_datatab.

  DATA: image_descriptions TYPE TABLE OF tt_datatab.
  FIELD-SYMBOLS: <img_descr> LIKE LINE OF image_descriptions.

  CONSTANTS: lv_default_file_name TYPE string       VALUE 'Calendar', "#EC NOTEXT
             c_from_row_portrait  TYPE zexcel_cell_row VALUE 28,
             c_from_row_landscape TYPE zexcel_cell_row VALUE 38,
             from_col TYPE zexcel_cell_column_alpha VALUE 'C',
             c_height_portrait  TYPE i              VALUE 450,   " Image Height in Portrait Mode
             c_height_landscape TYPE i              VALUE 670,   " Image Height in Landscape Mode
             c_factor TYPE f                        VALUE '1.5'. " Image Ratio, default 3:2

  IF p_path IS INITIAL.
    p_path = lv_workdir.
  ENDIF.
  cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_file_separator ).
  CONCATENATE p_path lv_file_separator lv_default_file_name '.xlsx' INTO lv_full_path. "#EC NOTEXT

  " Read Image Names for Month and Description
  CONCATENATE p_path lv_file_separator lv_default_file_name '.txt' INTO image_descr_path. "#EC NOTEXT
  cl_gui_frontend_services=>gui_upload(
    EXPORTING
      filename                = image_descr_path   " Name of file
      filetype                = 'ASC'    " File Type (ASCII, Binary)
      has_field_separator     = 'X'
      read_by_line            = 'X'    " File Written Line-By-Line to the Internal Table
    CHANGING
      data_tab                = image_descriptions    " Transfer table for file contents
    EXCEPTIONS
      file_open_error         = 1
      file_read_error         = 2
      no_batch                = 3
      gui_refuse_filetransfer = 4
      invalid_type            = 5
      no_authority            = 6
      unknown_error           = 7
      bad_data_format         = 8
      header_not_allowed      = 9
      separator_not_allowed   = 10
      header_too_long         = 11
      unknown_dp_error        = 12
      access_denied           = 13
      dp_out_of_memory        = 14
      disk_full               = 15
      dp_timeout              = 16
      not_supported_by_gui    = 17
      error_no_gui            = 18
      OTHERS                  = 19
  ).
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Create Styles
  " Create an underline double style
  lo_style_month                        = lo_excel->add_new_style( ).
  " lo_style_month->font->underline       = abap_true.
  " lo_style_month->font->underline_mode  = zcl_excel_style_font=>c_underline_single.
  lo_style_month->font->name            = zcl_excel_style_font=>c_name_roman.
  lo_style_month->font->scheme          = zcl_excel_style_font=>c_scheme_none.
  lo_style_month->font->family          = zcl_excel_style_font=>c_family_roman.
  lo_style_month->font->bold            = abap_true.
  lo_style_month->font->size            = 36.
  lv_style_month_guid                   = lo_style_month->get_guid( ).
  " Create border object
  CREATE OBJECT lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style     = zcl_excel_style_border=>c_border_thin.
  "Create style with border
  lo_style_border                         = lo_excel->add_new_style( ).
  lo_style_border->borders->allborders    = lo_border_dark.
  lo_style_border->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_right.
  lo_style_border->alignment->vertical    = zcl_excel_style_alignment=>c_vertical_top.
  lv_style_border_guid                    = lo_style_border->get_guid( ).
  "Create style alignment center
  lo_style_center                         = lo_excel->add_new_style( ).
  lo_style_center->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_center->alignment->vertical    = zcl_excel_style_alignment=>c_vertical_top.
  lv_style_center_guid                    = lo_style_center->get_guid( ).

  " Get Month Names
  CALL FUNCTION 'MONTH_NAMES_GET'
    TABLES
      month_names = month_names.

  zcl_date_calculation=>months_between_two_dates(
    EXPORTING
      i_date_from = p_from
      i_date_to   = p_to
      i_incl_to   = abap_true
    IMPORTING
      e_month     = month
  ).

  date_from = p_from.

  WHILE count <= month.
    IF count = 1.
      " Get active sheet
      lo_worksheet = lo_excel->get_active_worksheet( ).
    ELSE.
      lo_worksheet = lo_excel->add_new_worksheet(  ).
    ENDIF.

    lo_worksheet->zif_excel_sheet_properties~selected = zif_excel_sheet_properties=>c_selected.

    title = count.
    value = count.
    CONDENSE title.
    CONDENSE value.
    lo_worksheet->set_title( title ).
    lo_worksheet->set_print_gridlines( abap_false ).
    lo_worksheet->sheet_setup->paper_size = zcl_excel_sheet_setup=>c_papersize_a4.
    lo_worksheet->sheet_setup->horizontal_centered = abap_true.
    lo_worksheet->sheet_setup->vertical_centered   = abap_true.
    lo_column = lo_worksheet->get_column( 'A' ).
    lo_column->set_width( '1.0' ).
    lo_column = lo_worksheet->get_column( 'B' ).
    lo_column->set_width( '2.0' ).
    IF p_lands = abap_true.
      lo_worksheet->sheet_setup->orientation = zcl_excel_sheet_setup=>c_orientation_landscape.
      lv_height = c_height_landscape.
      lv_from_row = c_from_row_landscape.
      lo_worksheet->sheet_setup->margin_top    = '0.10'.
      lo_worksheet->sheet_setup->margin_left   = '0.10'.
      lo_worksheet->sheet_setup->margin_right  = '0.10'.
      lo_worksheet->sheet_setup->margin_bottom = '0.10'.
    ELSE.
      lo_column = lo_worksheet->get_column( 'K' ).
      lo_column->set_width( '3.0' ).
      lo_worksheet->sheet_setup->margin_top    = '0.80'.
      lo_worksheet->sheet_setup->margin_left   = '0.55'.
      lo_worksheet->sheet_setup->margin_right  = '0.05'.
      lo_worksheet->sheet_setup->margin_bottom = '0.30'.
      lv_height = c_height_portrait.
      lv_from_row = c_from_row_portrait.
    ENDIF.

    " Add Month Name
    month_nr = date_from+4(2).
    IF p_portr = abap_true.
      READ TABLE month_names WITH KEY mnr = month_nr ASSIGNING <month_name>.
      CONCATENATE <month_name>-ltx ` ` date_from(4) INTO value.
      row = lv_from_row - 2.
      to_col = from_col.
    ELSE.
      row = lv_from_row - 1.
      to_col_int = zcl_excel_common=>convert_column2int( from_col ) + 32.
      to_col = zcl_excel_common=>convert_column2alpha( to_col_int ).
      to_col_int = to_col_int + 1.
      to_col_end = zcl_excel_common=>convert_column2alpha( to_col_int ).
      CONCATENATE month_nr '/' date_from+2(2) INTO value.
      to_row = row + 2.
      lo_worksheet->set_merge(
        EXPORTING
          ip_column_start = to_col  " Cell Column Start
          ip_column_end   = to_col_end    " Cell Column End
          ip_row          = row       " Cell Row
          ip_row_to       = to_row    " Cell Row
      ).
    ENDIF.
    lo_worksheet->set_cell(
      EXPORTING
        ip_column    = to_col " Cell Column
        ip_row       = row      " Cell Row
        ip_value     = value    " Cell Value
        ip_style     = lv_style_month_guid
    ).

*    to_col_int = zcl_excel_common=>convert_column2int( from_col ) + 7.
*    to_col = zcl_excel_common=>convert_column2alpha( to_col_int ).
*
*    lo_worksheet->set_merge(
*      EXPORTING
*        ip_column_start = from_col  " Cell Column Start
*        ip_column_end   = to_col    " Cell Column End
*        ip_row          = row       " Cell Row
*        ip_row_to       = row       " Cell Row
*    ).

    " Add drawing from a XSTRING read from a file
    UNASSIGN <img_descr>.
    READ TABLE image_descriptions WITH KEY month_nr = month_nr ASSIGNING <img_descr>.
    IF <img_descr> IS ASSIGNED.
      value = <img_descr>-descr.
      IF p_portr = abap_true.
        row = lv_from_row - 3.
      ELSE.
        row = lv_from_row - 2.
      ENDIF.
      IF NOT <img_descr>-url IS INITIAL.
        lo_hyperlink = zcl_excel_hyperlink=>create_external_link( <img_descr>-url ).
        lo_worksheet->set_cell(
          EXPORTING
            ip_column    = from_col " Cell Column
            ip_row       = row      " Cell Row
            ip_value     = value    " Cell Value
            ip_hyperlink = lo_hyperlink
        ).
      ELSE.
        lo_worksheet->set_cell(
          EXPORTING
            ip_column    = from_col " Cell Column
            ip_row       = row      " Cell Row
            ip_value     = value    " Cell Value
        ).
      ENDIF.
      lo_row = lo_worksheet->get_row( row ).
      lo_row->set_row_height( '22.0' ).

      " In Landscape mode the row between the description and the
      " dates should be not so high
      IF p_lands = abap_true.
        row = lv_from_row - 3.
        lo_worksheet->set_cell(
          EXPORTING
            ip_column    = from_col " Cell Column
            ip_row       = row      " Cell Row
            ip_value     = ' '      " Cell Value
        ).
        lo_row = lo_worksheet->get_row( row ).
        lo_row->set_row_height( '7.0' ).
        row = lv_from_row - 1.
        lo_row = lo_worksheet->get_row( row ).
        lo_row->set_row_height( '5.0' ).
      ENDIF.

      CONCATENATE p_path lv_file_separator <img_descr>-filename INTO image_path.
      lo_drawing = lo_excel->add_new_drawing( ).
      lo_drawing->set_position( ip_from_row = 1
                                ip_from_col = 'B' ).

      lv_content = zcl_helper=>load_image( image_path ).
      width = lv_height * c_factor.
      lo_drawing->set_media( ip_media      = lv_content
                             ip_media_type = zcl_excel_drawing=>c_media_type_jpg
                             ip_width      = width
                             ip_height     = lv_height  ).
      lo_worksheet->add_drawing( lo_drawing ).
    ENDIF.

    " Add Calendar
*    CALL FUNCTION 'SLS_MISC_GET_LAST_DAY_OF_MONTH'
*      EXPORTING
*        day_in            = date_from
*      IMPORTING
*        last_day_of_month = date_to.
    date_to = date_from.
    date_to+6(2) = '01'.              " First of month
    add 31 to date_to.                " Somewhere in following month
    date_to = date_to - date_to+6(2). " Last of month
    IF p_portr = abap_true.
      zcl_helper=>add_calendar(
        EXPORTING
          i_date_from = date_from
          i_date_to   = date_to
          i_from_row  = lv_from_row
          i_from_col  = from_col
          i_day_style = lv_style_border_guid
          i_cw_style  = lv_style_center_guid
        CHANGING
          c_worksheet = lo_worksheet
      ).
    ELSE.
      zcl_helper=>add_calendar_landscape(
        EXPORTING
          i_date_from = date_from
          i_date_to   = date_to
          i_from_row  = lv_from_row
          i_from_col  = from_col
          i_day_style = lv_style_border_guid
          i_cw_style  = lv_style_center_guid
        CHANGING
          c_worksheet = lo_worksheet
      ).
    ENDIF.
    count = count + 1.
    date_from = date_to + 1.
  ENDWHILE.

  lo_excel->set_active_sheet_index_by_name( '1' ).
*** Create output
  lcl_output=>output( lo_excel ).
