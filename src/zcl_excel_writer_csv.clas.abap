class ZCL_EXCEL_WRITER_CSV definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
public section.

  interfaces ZIF_EXCEL_WRITER .

  class-methods SET_DELIMITER
    importing
      value(IP_VALUE) type CHAR01 default ';' .
  class-methods SET_ENCLOSURE
    importing
      value(IP_VALUE) type CHAR01 default '"' .
  class-methods SET_ENDOFLINE
    importing
      value(IP_VALUE) type ANY default CL_ABAP_CHAR_UTILITIES=>CR_LF .
  class-methods SET_ACTIVE_SHEET_INDEX
    importing
      !I_ACTIVE_WORKSHEET type ZEXCEL_ACTIVE_WORKSHEET .
  class-methods SET_ACTIVE_SHEET_INDEX_BY_NAME
    importing
      !I_WORKSHEET_NAME type ZEXCEL_WORKSHEETS_NAME .
*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
private section.

  data EXCEL type ref to ZCL_EXCEL .
  class-data DELIMITER type CHAR01 value ';'. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
  class-data ENCLOSURE type CHAR01 value '"'. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
  class-data EOL type CHAR01 value CL_ABAP_CHAR_UTILITIES=>CR_LF. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
  class-data WORKSHEET_NAME type ZEXCEL_WORKSHEETS_NAME .
  class-data WORKSHEET_INDEX type ZEXCEL_ACTIVE_WORKSHEET .

  methods CREATE
    returning
      value(EP_EXCEL) type XSTRING .
  methods CREATE_CSV
    returning
      value(EP_CONTENT) type XSTRING .
ENDCLASS.



CLASS ZCL_EXCEL_WRITER_CSV IMPLEMENTATION.


method CREATE.

* .csv format with ; delimiter

   ep_excel = me->CREATE_CSV( ).

  endmethod.


method CREATE_CSV.

  TYPES: BEGIN OF lty_format,
            cmpname  TYPE SEOCMPNAME,
            attvalue TYPE SEOVALUE,
         END OF lty_format.
  DATA: lt_format TYPE STANDARD TABLE OF lty_format,
        ls_format LIKE LINE OF lt_format,
        lv_date   TYPE DATS,
        lv_tmp    TYPE string,
        lv_time   TYPE CHAR08.

  DATA: lo_iterator       TYPE REF TO cl_object_collection_iterator,
        lo_worksheet      TYPE REF TO zcl_excel_worksheet.

  DATA: lt_cell_data      TYPE zexcel_t_cell_data_unsorted,
        lv_row            TYPE sytabix,
        lv_col            TYPE sytabix,
        lv_string         TYPE string,
        lc_value          TYPE string,
        lv_attrname       TYPE SEOCMPNAME.

  DATA: ls_numfmt         TYPE zexcel_s_style_numfmt,
        lo_style          TYPE REF TO zcl_excel_style.

  FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data.

* --- Retrieve supported cell format
  REFRESH lt_format.
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
  data: current_worksheet_title type ZEXCEL_SHEET_TITLE.

  WHILE lo_iterator->if_object_collection_iterator~has_next( ) EQ abap_true.
    lo_worksheet ?= lo_iterator->if_object_collection_iterator~get_next( ).

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
      WHILE lo_iterator->if_object_collection_iterator~has_next( ) EQ abap_true.
        lo_style ?= lo_iterator->if_object_collection_iterator~get_next( ).
        CHECK lo_style->get_guid( ) = <fs_sheet_content>-cell_style.
        ls_numfmt     = lo_style->number_format->get_structure( ).
        EXIT.
      ENDWHILE.
    ENDIF.
    IF <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.
      " determine data-type
      CLEAR lv_attrname.
      CONCATENATE '''' ls_numfmt-NUMFMT '''' INTO ls_numfmt-NUMFMT.
      TRANSLATE ls_numfmt-numfmt TO UPPER CASE.
      READ TABLE lt_format INTO ls_format WITH KEY attvalue = ls_numfmt-NUMFMT.
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
*      CONCATENATE lv_string cl_abap_char_utilities=>newline INTO lv_string.
*      CONCATENATE lv_string cl_abap_char_utilities=>cr_lf INTO lv_string.
      CONCATENATE lv_string zcl_excel_writer_csv=>eol INTO lv_string.
      lv_row = lv_row + 1.
      lv_col = 1.
    ENDWHILE.

* --- Add empty columns
    WHILE lv_col < <fs_sheet_content>-cell_column.
*      CONCATENATE lv_string ';' INTO lv_string.
      CONCATENATE lv_string zcl_excel_writer_csv=>delimiter INTO lv_string.
      lv_col = lv_col + 1.
    ENDWHILE.

* ----- Use format to determine the data type and display format.
    CASE <fs_sheet_content>-data_type.
*      WHEN 'n' OR 'N'.
*        lc_value = zcl_excel_common=>excel_number_to_string( ip_value = <fs_sheet_content>-cell_value ).

      WHEN 'd' OR 'D'.
        lc_value = zcl_excel_common=>excel_string_to_date( ip_value = <fs_sheet_content>-cell_value ).
        TRY.
          lv_date = lc_value.
          CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
            EXPORTING
              DATE_INTERNAL                  = lv_date
            IMPORTING
              DATE_EXTERNAL                  = lv_tmp
            EXCEPTIONS
              DATE_INTERNAL_IS_INVALID       = 1
              OTHERS                         = 2
                  .
          IF SY-SUBRC = 0.
            lc_value = lv_tmp.
          ENDIF.

        CATCH CX_SY_CONVERSION_NO_NUMBER.

        ENDTRY.

      WHEN 't' OR 'T'.
        lc_value = zcl_excel_common=>excel_string_to_time( ip_value = <fs_sheet_content>-cell_value ).
        write lc_value to lv_time USING EDIT MASK '__:__:__'.
        lc_value = lv_time.
      WHEN OTHERS.
        lc_value = <fs_sheet_content>-cell_value.

    ENDCASE.

*    REPLACE ALL OCCURRENCES OF '"' in lc_value with '""'.
    CONCATENATE zcl_excel_writer_csv=>enclosure zcl_excel_writer_csv=>enclosure INTO lv_tmp.
    CONDENSE lv_tmp.
    REPLACE ALL OCCURRENCES OF zcl_excel_writer_csv=>enclosure in lc_value with lv_tmp.

*    FIND FIRST OCCURRENCE OF ';' IN lc_value.
    FIND FIRST OCCURRENCE OF zcl_excel_writer_csv=>delimiter IN lc_value.
    IF sy-subrc = 0.
      CONCATENATE lv_string zcl_excel_writer_csv=>enclosure lc_value zcl_excel_writer_csv=>enclosure INTO lv_string.
    ELSE.
      CONCATENATE lv_string lc_value INTO lv_string.
    ENDIF.

  ENDLOOP. " AT lt_cell_data

  CLEAR ep_content.

  CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
    EXPORTING
      TEXT           = lv_string
*      MIMETYPE       = ' '
*   ENCODING       =
    IMPORTING
      BUFFER         = ep_content
    EXCEPTIONS
      FAILED         = 1
      OTHERS         = 2
          .

  endmethod.


method SET_ACTIVE_SHEET_INDEX.
  CLEAR WORKSHEET_NAME.
  WORKSHEET_INDEX = i_active_worksheet.
  endmethod.


method SET_ACTIVE_SHEET_INDEX_BY_NAME.
  CLEAR WORKSHEET_INDEX.
  WORKSHEET_NAME = i_worksheet_name.
  endmethod.


method SET_DELIMITER.
  delimiter = ip_value.
  endmethod.


method SET_ENCLOSURE.
  zcl_excel_writer_csv=>enclosure = ip_value.
  endmethod.


method SET_ENDOFLINE.
  zcl_excel_writer_csv=>eol = ip_value.
  endmethod.


method ZIF_EXCEL_WRITER~WRITE_FILE.
  me->excel = io_excel.
  ep_file = me->create( ).
  endmethod.
ENDCLASS.
