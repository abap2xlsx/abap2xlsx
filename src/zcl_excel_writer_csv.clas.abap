CLASS zcl_excel_writer_csv DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  PUBLIC SECTION.

    INTERFACES zif_excel_writer .

    CLASS-METHODS set_delimiter
      IMPORTING
        VALUE(ip_value) TYPE c DEFAULT ';' .
    CLASS-METHODS set_enclosure
      IMPORTING
        VALUE(ip_value) TYPE c DEFAULT '"' .
    CLASS-METHODS set_endofline
      IMPORTING
        VALUE(ip_value) TYPE any DEFAULT cl_abap_char_utilities=>cr_lf .
    CLASS-METHODS set_active_sheet_index
      IMPORTING
        !i_active_worksheet TYPE zexcel_active_worksheet .
    CLASS-METHODS set_active_sheet_index_by_name
      IMPORTING
        !i_worksheet_name TYPE zexcel_worksheets_name .
*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA excel TYPE REF TO zcl_excel .
    CLASS-DATA delimiter TYPE c VALUE ';' ##NO_TEXT.
    CLASS-DATA enclosure TYPE c VALUE '"' ##NO_TEXT.
    CLASS-DATA:
      eol TYPE c LENGTH 2 VALUE cl_abap_char_utilities=>cr_lf ##NO_TEXT.
    CLASS-DATA worksheet_name TYPE zexcel_worksheets_name .
    CLASS-DATA worksheet_index TYPE zexcel_active_worksheet .

    METHODS create
      RETURNING
        VALUE(ep_excel) TYPE xstring
      RAISING
        zcx_excel.
    METHODS create_csv
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
ENDCLASS.



CLASS zcl_excel_writer_csv IMPLEMENTATION.


  METHOD create.

* .csv format with ; delimiter

    ep_excel = me->create_csv( ).

  ENDMETHOD.


  METHOD create_csv.

    TYPES: BEGIN OF lty_format,
             cmpname  TYPE seocmpname,
             attvalue TYPE seovalue,
           END OF lty_format.
    DATA: lt_format TYPE STANDARD TABLE OF lty_format,
          ls_format LIKE LINE OF lt_format,
          lv_date   TYPE d,
          lv_tmp    TYPE string,
          lv_time   TYPE c LENGTH 8.

    DATA: lo_iterator  TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet TYPE REF TO zcl_excel_worksheet.

    DATA: lt_cell_data TYPE zexcel_t_cell_data_unsorted,
          lv_row       TYPE i,
          lv_col       TYPE i,
          lv_string    TYPE string,
          lc_value     TYPE string,
          lv_attrname  TYPE seocmpname.

    DATA: ls_numfmt TYPE zexcel_s_style_numfmt,
          lo_style  TYPE REF TO zcl_excel_style.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data.

* --- Retrieve supported cell format
    SELECT * INTO CORRESPONDING FIELDS OF TABLE lt_format
      FROM seocompodf
     WHERE clsname  = 'ZCL_EXCEL_STYLE_NUMBER_FORMAT'
       AND typtype  = 1
       AND type     = 'ZEXCEL_NUMBER_FORMAT'.

* --- Retrieve SAP date format
    CLEAR ls_format.
    SELECT ddtext INTO ls_format-attvalue FROM dd07t WHERE domname    = 'XUDATFM'
                                                       AND ddlanguage = sy-langu.
      ls_format-cmpname = 'DATE'.
      CONDENSE ls_format-attvalue.
      CONCATENATE '''' ls_format-attvalue '''' INTO ls_format-attvalue.
      APPEND ls_format TO lt_format.
    ENDSELECT.


    LOOP AT lt_format INTO ls_format.
      TRANSLATE ls_format-attvalue TO UPPER CASE.
      MODIFY lt_format FROM ls_format.
    ENDLOOP.


* STEP 1: Collect strings from the first worksheet
    lo_iterator = excel->get_worksheets_iterator( ).
    DATA: current_worksheet_title TYPE zexcel_sheet_title.

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      IF worksheet_name IS NOT INITIAL.
        current_worksheet_title = lo_worksheet->get_title( ).
        CHECK current_worksheet_title = worksheet_name.
      ELSE.
        IF worksheet_index IS INITIAL.
          worksheet_index = 1.
        ENDIF.
        CHECK worksheet_index = sy-index.
      ENDIF.
      APPEND LINES OF lo_worksheet->sheet_content TO lt_cell_data.
      EXIT. " Take first worksheet only
    ENDWHILE.

    DELETE lt_cell_data WHERE cell_formula IS NOT INITIAL. " delete formula content

    SORT lt_cell_data BY cell_row
                         cell_column.
    lv_row = 1.
    lv_col = 1.
    CLEAR lv_string.
    LOOP AT lt_cell_data ASSIGNING <fs_sheet_content>.

*   --- Retrieve Cell Style format and data type
      CLEAR ls_numfmt.
      IF <fs_sheet_content>-data_type IS INITIAL AND <fs_sheet_content>-cell_style IS NOT INITIAL.
        lo_iterator = excel->get_styles_iterator( ).
        WHILE lo_iterator->has_next( ) EQ abap_true.
          lo_style ?= lo_iterator->get_next( ).
          CHECK lo_style->get_guid( ) = <fs_sheet_content>-cell_style.
          ls_numfmt     = lo_style->number_format->get_structure( ).
          EXIT.
        ENDWHILE.
      ENDIF.
      IF <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.
        " determine data-type
        CLEAR lv_attrname.
        CONCATENATE '''' ls_numfmt-numfmt '''' INTO ls_numfmt-numfmt.
        TRANSLATE ls_numfmt-numfmt TO UPPER CASE.
        READ TABLE lt_format INTO ls_format WITH KEY attvalue = ls_numfmt-numfmt.
        IF sy-subrc = 0.
          lv_attrname = ls_format-cmpname.
        ENDIF.

        IF lv_attrname IS NOT INITIAL.
          FIND FIRST OCCURRENCE OF 'DATETIME' IN lv_attrname.
          IF sy-subrc = 0.
            <fs_sheet_content>-data_type = 'd'.
          ELSE.
            FIND FIRST OCCURRENCE OF 'TIME' IN lv_attrname.
            IF sy-subrc = 0.
              <fs_sheet_content>-data_type = 't'.
            ELSE.
              FIND FIRST OCCURRENCE OF 'DATE' IN lv_attrname.
              IF sy-subrc = 0.
                <fs_sheet_content>-data_type = 'd'.
              ELSE.
                FIND FIRST OCCURRENCE OF 'CURRENCY' IN lv_attrname.
                IF sy-subrc = 0.
                  <fs_sheet_content>-data_type = 'n'.
                ELSE.
                  FIND FIRST OCCURRENCE OF 'NUMBER' IN lv_attrname.
                  IF sy-subrc = 0.
                    <fs_sheet_content>-data_type = 'n'.
                  ELSE.
                    FIND FIRST OCCURRENCE OF 'PERCENTAGE' IN lv_attrname.
                    IF sy-subrc = 0.
                      <fs_sheet_content>-data_type = 'n'.
                    ENDIF. " Purcentage
                  ENDIF. " Number
                ENDIF. " Currency
              ENDIF. " Date
            ENDIF. " TIME
          ENDIF. " DATETIME
        ENDIF. " lv_attrname IS NOT INITIAL.
      ENDIF. " <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.

* --- Add empty rows
      WHILE lv_row < <fs_sheet_content>-cell_row.
        CONCATENATE lv_string zcl_excel_writer_csv=>eol INTO lv_string.
        lv_row = lv_row + 1.
        lv_col = 1.
      ENDWHILE.

* --- Add empty columns
      WHILE lv_col < <fs_sheet_content>-cell_column.
        CONCATENATE lv_string zcl_excel_writer_csv=>delimiter INTO lv_string.
        lv_col = lv_col + 1.
      ENDWHILE.

* ----- Use format to determine the data type and display format.
      CASE <fs_sheet_content>-data_type.

        WHEN 'd' OR 'D'.
          lc_value = zcl_excel_common=>excel_string_to_date( ip_value = <fs_sheet_content>-cell_value ).
          TRY.
              lv_date = lc_value.
              CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
                EXPORTING
                  date_internal            = lv_date
                IMPORTING
                  date_external            = lv_tmp
                EXCEPTIONS
                  date_internal_is_invalid = 1
                  OTHERS                   = 2.
              IF sy-subrc = 0.
                lc_value = lv_tmp.
              ENDIF.

            CATCH cx_sy_conversion_no_number.

          ENDTRY.

        WHEN 't' OR 'T'.
          lc_value = zcl_excel_common=>excel_string_to_time( ip_value = <fs_sheet_content>-cell_value ).
          WRITE lc_value TO lv_time USING EDIT MASK '__:__:__'.
          lc_value = lv_time.
        WHEN OTHERS.
          lc_value = <fs_sheet_content>-cell_value.

      ENDCASE.

      CONCATENATE zcl_excel_writer_csv=>enclosure zcl_excel_writer_csv=>enclosure INTO lv_tmp.
      CONDENSE lv_tmp.
      REPLACE ALL OCCURRENCES OF zcl_excel_writer_csv=>enclosure IN lc_value WITH lv_tmp.

      FIND FIRST OCCURRENCE OF zcl_excel_writer_csv=>delimiter IN lc_value.
      IF sy-subrc = 0.
        CONCATENATE lv_string zcl_excel_writer_csv=>enclosure lc_value zcl_excel_writer_csv=>enclosure INTO lv_string.
      ELSE.
        CONCATENATE lv_string lc_value INTO lv_string.
      ENDIF.

    ENDLOOP.

    CLEAR ep_content.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text   = lv_string
      IMPORTING
        buffer = ep_content
      EXCEPTIONS
        failed = 1
        OTHERS = 2.

  ENDMETHOD.


  METHOD set_active_sheet_index.
    CLEAR worksheet_name.
    worksheet_index = i_active_worksheet.
  ENDMETHOD.


  METHOD set_active_sheet_index_by_name.
    CLEAR worksheet_index.
    worksheet_name = i_worksheet_name.
  ENDMETHOD.


  METHOD set_delimiter.
    delimiter = ip_value.
  ENDMETHOD.


  METHOD set_enclosure.
    zcl_excel_writer_csv=>enclosure = ip_value.
  ENDMETHOD.


  METHOD set_endofline.
    zcl_excel_writer_csv=>eol = ip_value.
  ENDMETHOD.


  METHOD zif_excel_writer~write_file.
    me->excel = io_excel.
    ep_file = me->create( ).
  ENDMETHOD.
ENDCLASS.
