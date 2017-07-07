*&---------------------------------------------------------------------*
*&  Include           ZDEMO_CALENDAR_CLASSES
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*&       Class ZCL_DATE_CALCULATION
*&---------------------------------------------------------------------*
*        Text
*----------------------------------------------------------------------*
CLASS zcl_date_calculation DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS: months_between_two_dates
    IMPORTING
      i_date_from TYPE datum
      i_date_to   TYPE datum
      i_incl_to   TYPE flag
    EXPORTING
      e_month     TYPE i.
ENDCLASS.               "ZCL_DATE_CALCULATION


*----------------------------------------------------------------------*
*       CLASS ZCL_DATE_CALCULATION IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_date_calculation IMPLEMENTATION.
  METHOD months_between_two_dates.
    DATA: date_to TYPE datum.
    DATA: BEGIN OF datum_von,
            jjjj(4) TYPE n,
              mm(2) TYPE n,
              tt(2) TYPE n,
          END OF datum_von.

    DATA: BEGIN OF datum_bis,
            jjjj(4) TYPE n,
              mm(2) TYPE n,
              tt(2) TYPE n,
          END OF datum_bis.

    e_month = 0.

    CHECK NOT ( i_date_from IS INITIAL )
      AND NOT ( i_date_to IS INITIAL ).

    date_to = i_date_to.
    IF i_incl_to = abap_true.
      date_to = date_to + 1.
    ENDIF.

    datum_von = i_date_from.
    datum_bis = date_to.

    e_month   =   ( datum_bis-jjjj - datum_von-jjjj ) * 12
                + ( datum_bis-mm   - datum_von-mm   ).
  ENDMETHOD.                    "MONTHS_BETWEEN_TWO_DATES
ENDCLASS.                    "ZCL_DATE_CALCULATION IMPLEMENTATION

*----------------------------------------------------------------------*
*       CLASS zcl_date_calculation_test DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_date_calculation_test DEFINITION FOR TESTING
  " DURATION SHORT
  " RISK LEVEL HARMLESS
  "#AU Duration Medium
  "#AU Risk_Level Harmless
  .
  PUBLIC SECTION.
    METHODS:
      months_between_two_dates FOR TESTING.
ENDCLASS.                    "zcl_date_calculation_test DEFINITION
*----------------------------------------------------------------------*
*       CLASS zcl_date_calculation_test IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_date_calculation_test IMPLEMENTATION.
  METHOD months_between_two_dates.

    DATA: date_from TYPE datum VALUE '20120101',
          date_to   TYPE datum VALUE '20121231'.
    DATA: month TYPE i.

    zcl_date_calculation=>months_between_two_dates(
      EXPORTING
        i_date_from = date_from
        i_date_to   = date_to
        i_incl_to   = abap_true
      IMPORTING
        e_month     = month
    ).

    cl_aunit_assert=>assert_equals(
        exp                  = 12    " Data Object with Expected Type
        act                  = month    " Data Object with Current Value
        msg                  = 'Calculated date is wrong'    " Message in Case of Error
    ).

  ENDMETHOD.                    "months_between_two_dates
ENDCLASS.                    "zcl_date_calculation_test IMPLEMENTATION
*----------------------------------------------------------------------*
*       CLASS zcl_helper DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_helper DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS:
    load_image
      IMPORTING
        filename TYPE string
      RETURNING value(r_image) TYPE xstring,
    add_calendar
      IMPORTING
        i_date_from TYPE datum
        i_date_to   TYPE datum
        i_from_row  TYPE zexcel_cell_row
        i_from_col  TYPE zexcel_cell_column_alpha
        i_day_style TYPE zexcel_cell_style
        i_cw_style  TYPE zexcel_cell_style
      CHANGING
        c_worksheet TYPE REF TO zcl_excel_worksheet,
    add_calendar_landscape
      IMPORTING
        i_date_from TYPE datum
        i_date_to   TYPE datum
        i_from_row  TYPE zexcel_cell_row
        i_from_col  TYPE zexcel_cell_column_alpha
        i_day_style TYPE zexcel_cell_style
        i_cw_style  TYPE zexcel_cell_style
      CHANGING
        c_worksheet TYPE REF TO zcl_excel_worksheet,
    add_a2x_footer
      IMPORTING
        i_from_row  TYPE zexcel_cell_row
        i_from_col  TYPE zexcel_cell_column_alpha
      CHANGING
        c_worksheet TYPE REF TO zcl_excel_worksheet,
    add_calender_week
      IMPORTING
        i_date  TYPE datum
        i_row   TYPE zexcel_cell_row
        i_col   TYPE zexcel_cell_column_alpha
        i_style TYPE zexcel_cell_style
      CHANGING
        c_worksheet TYPE REF TO zcl_excel_worksheet.
ENDCLASS.                    "zcl_helper DEFINITION

*----------------------------------------------------------------------*
*       CLASS zcl_helper IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_helper IMPLEMENTATION.
  METHOD load_image.
    "Load samle image
    DATA: lt_bin TYPE solix_tab,
          lv_len TYPE i.

    CALL METHOD cl_gui_frontend_services=>gui_upload
      EXPORTING
        filename                = filename
        filetype                = 'BIN'
      IMPORTING
        filelength              = lv_len
      CHANGING
        data_tab                = lt_bin
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
        OTHERS                  = 19.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_len
      IMPORTING
        buffer       = r_image
      TABLES
        binary_tab   = lt_bin
      EXCEPTIONS
        failed       = 1
        OTHERS       = 2.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDMETHOD.                    "load_image
  METHOD add_calendar.
    DATA: day_names     TYPE TABLE OF t246.
    DATA: row           TYPE zexcel_cell_row,
          row_max       TYPE i,
          col_int       TYPE zexcel_cell_column,
          col_max       TYPE i,
          from_col_int  TYPE zexcel_cell_column,
          col           TYPE zexcel_cell_column_alpha,
          lo_column     TYPE REF TO zcl_excel_column,
          lo_row        TYPE REF TO zcl_excel_row.
    DATA: lv_date          TYPE datum,
          value            TYPE string,
          weekday          TYPE wotnr,
          weekrow          TYPE wotnr VALUE 1,
          day              TYPE i,
          width            TYPE f,
          height           TYPE f.
    DATA: hyperlink TYPE REF TO zcl_excel_hyperlink.

    FIELD-SYMBOLS: <day_name> LIKE LINE OF day_names.

    lv_date = i_date_from.
    from_col_int = zcl_excel_common=>convert_column2int( i_from_col ).
    " Add description for Calendar Week
    c_worksheet->set_cell(
      EXPORTING
        ip_column    = i_from_col    " Cell Column
        ip_row       = i_from_row    " Cell Row
        ip_value     = 'CW'(001)    " Cell Value
        ip_style     = i_cw_style
    ).

    " Add Days
    CALL FUNCTION 'DAY_NAMES_GET'
      TABLES
        day_names = day_names.

    LOOP AT day_names ASSIGNING <day_name>.
      row     = i_from_row.
      col_int = from_col_int + <day_name>-wotnr.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      value = <day_name>-langt.
      c_worksheet->set_cell(
        EXPORTING
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_cw_style
      ).
    ENDLOOP.

    WHILE lv_date <= i_date_to.
      day = lv_date+6(2).
      CALL FUNCTION 'FIMA_X_DAY_IN_MONTH_COMPUTE'
        EXPORTING
          i_datum        = lv_date
        IMPORTING
          e_wochentag_nr = weekday.

      row     = i_from_row + weekrow.
      col_int = from_col_int + weekday.
      col = zcl_excel_common=>convert_column2alpha( col_int ).

      value = day.
      CONDENSE value.

      c_worksheet->set_cell(
        EXPORTING
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_day_style    " Single-Character Indicator
      ).

      IF weekday = 7.
        " Add Calender Week
        zcl_helper=>add_calender_week(
          EXPORTING
            i_date      = lv_date
            i_row       = row
            i_col       = i_from_col
            i_style     = i_cw_style
          CHANGING
            c_worksheet = c_worksheet
        ).
        weekrow = weekrow + 1.
      ENDIF.
      lv_date = lv_date + 1.
    ENDWHILE.
    " Add Calender Week
    zcl_helper=>add_calender_week(
      EXPORTING
        i_date      = lv_date
        i_row       = row
        i_col       = i_from_col
        i_style     = i_cw_style
      CHANGING
        c_worksheet = c_worksheet
    ).
    " Add Created with abap2xlsx
    row = row + 2.
    zcl_helper=>add_a2x_footer(
      EXPORTING
        i_from_row  = row
        i_from_col  = i_from_col
      CHANGING
        c_worksheet = c_worksheet
    ).
    col_int = from_col_int.
    col_max = from_col_int + 7.
    WHILE col_int <= col_max.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      IF sy-index = 1.
        width = '5.0'.
      ELSE.
        width = '11.4'.
      ENDIF.
      lo_column = c_worksheet->get_column( col ).
      lo_column->set_width( width ).
      col_int = col_int + 1.
    ENDWHILE.
    row = i_from_row + 1.
    row_max = i_from_row + 6.
    WHILE row <= row_max.
      height = 50.
      lo_row = c_worksheet->get_row( row ).
      lo_row->set_row_height( height ).
      row = row + 1.
    ENDWHILE.
  ENDMETHOD.                    "add_calendar
  METHOD add_a2x_footer.
    DATA: value TYPE string,
          hyperlink TYPE REF TO zcl_excel_hyperlink.

    value = 'Created with abap2xlsx. Find more information at http://www.plinky.it/abap/abap2xlsx.php.'(002).
    hyperlink = zcl_excel_hyperlink=>create_external_link( 'http://www.plinky.it/abap/abap2xlsx.php' ). "#EC NOTEXT
    c_worksheet->set_cell(
      EXPORTING
        ip_column    = i_from_col    " Cell Column
        ip_row       = i_from_row    " Cell Row
        ip_value     = value    " Cell Value
        ip_hyperlink = hyperlink
    ).

  ENDMETHOD.                    "add_a2x_footer
  METHOD add_calendar_landscape.
    DATA: day_names TYPE TABLE OF t246.

    DATA: lv_date          TYPE datum,
          day              TYPE i,
          value            TYPE string,
          weekday          TYPE wotnr.
    DATA: row               TYPE zexcel_cell_row,
          from_col_int      TYPE zexcel_cell_column,
          col_int           TYPE zexcel_cell_column,
          col               TYPE zexcel_cell_column_alpha.
    DATA: lo_column         TYPE REF TO zcl_excel_column,
          lo_row            TYPE REF TO zcl_excel_row.

    FIELD-SYMBOLS: <day_name> LIKE LINE OF day_names.

    lv_date = i_date_from.
    " Add Days
    CALL FUNCTION 'DAY_NAMES_GET'
      TABLES
        day_names = day_names.

    WHILE lv_date <= i_date_to.
      day = lv_date+6(2).
      CALL FUNCTION 'FIMA_X_DAY_IN_MONTH_COMPUTE'
        EXPORTING
          i_datum        = lv_date
        IMPORTING
          e_wochentag_nr = weekday.
      " Day name row
      row     = i_from_row.
      col_int = from_col_int + day + 2.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      READ TABLE day_names ASSIGNING <day_name>
        WITH KEY wotnr = weekday.
      value = <day_name>-kurzt.
      c_worksheet->set_cell(
        EXPORTING
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_cw_style
      ).

      " Day row
      row     = i_from_row + 1.
      value = day.
      CONDENSE value.

      c_worksheet->set_cell(
        EXPORTING
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_day_style    " Single-Character Indicator
      ).
      " width
      lo_column = c_worksheet->get_column( col ).
      lo_column->set_width( '3.6' ).


      lv_date = lv_date + 1.
    ENDWHILE.
    " Add ABAP2XLSX Footer
    row = i_from_row + 2.
    c_worksheet->set_cell(
      EXPORTING
        ip_column    = col    " Cell Column
        ip_row       = row    " Cell Row
        ip_value     = ' '    " Cell Value
    ).
    lo_row = c_worksheet->get_row( row ).
    lo_row->set_row_height( '5.0' ).
    row = i_from_row + 3.
    zcl_helper=>add_a2x_footer(
      EXPORTING
        i_from_row  = row
        i_from_col  = i_from_col
      CHANGING
        c_worksheet = c_worksheet
    ).

    " Set with for all 31 coulumns
    WHILE day < 32.
      day = day + 1.
      col_int = from_col_int + day + 2.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      " width
      lo_column = c_worksheet->get_column( col ).
      lo_column->set_width( '3.6' ).
    ENDWHILE.
  ENDMETHOD.                    "ADD_CALENDAR_LANDSCAPE

  METHOD add_calender_week.
    DATA: week     TYPE kweek,
          week_int TYPE i,
          value    TYPE string.
    " Add Calender Week
    CALL FUNCTION 'DATE_GET_WEEK'
      EXPORTING
        date = i_date  " Date for which the week should be calculated
      IMPORTING
        week = week.    " Week for date (format:YYYYWW)
    value = week+4(2).
    week_int = value.
    value = week_int.
    CONDENSE value.
    c_worksheet->set_cell(
      EXPORTING
        ip_column    = i_col    " Cell Column
        ip_row       = i_row    " Cell Row
        ip_value     = value    " Cell Value
        ip_style     = i_style
    ).
  ENDMETHOD.                    "add_calender_week
ENDCLASS.                    "zcl_helper IMPLEMENTATION
