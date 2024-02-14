class ZCL_EXCEL_WRITER_CSV definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
public section.

  interfaces ZIF_EXCEL_WRITER .

  class-methods SET_DELIMITER
    importing
      value(IP_VALUE) type C default ';' .
  class-methods SET_ENCLOSURE
    importing
      value(IP_VALUE) type C default '"' .
  class-methods SET_ENDOFLINE
    importing
      value(IP_VALUE) type ANY default CL_ABAP_CHAR_UTILITIES=>CR_LF .
  class-methods SET_ACTIVE_SHEET_INDEX
    importing
      !I_ACTIVE_WORKSHEET type ZEXCEL_ACTIVE_WORKSHEET .
  class-methods SET_ACTIVE_SHEET_INDEX_BY_NAME
    importing
      !I_WORKSHEET_NAME type ZEXCEL_WORKSHEETS_NAME .
  class-methods SET_INITIAL_EXT_DATE
    importing
      !IP_VALUE type CHAR10 default 'DEFAULT' .
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
private section.

  data EXCEL type ref to ZCL_EXCEL .
  class-data DELIMITER type C value ';' ##NO_TEXT.
  class-data ENCLOSURE type C value '"' ##NO_TEXT.
  class-data:
    eol TYPE c LENGTH 2 value CL_ABAP_CHAR_UTILITIES=>CR_LF ##NO_TEXT.
  class-data WORKSHEET_NAME type ZEXCEL_WORKSHEETS_NAME .
  class-data WORKSHEET_INDEX type ZEXCEL_ACTIVE_WORKSHEET .
  class-data INITIAL_EXT_DATE type CHAR10 value 'DEFAULT' ##NO_TEXT.
  constants C_DEFAULT type CHAR10 value 'DEFAULT' ##NO_TEXT.

  methods CREATE
    returning
      value(EP_EXCEL) type XSTRING
    raising
      ZCX_EXCEL .
  methods CREATE_CSV
    returning
      value(EP_CONTENT) type XSTRING
    raising
      ZCX_EXCEL .
ENDCLASS.



CLASS ZCL_EXCEL_WRITER_CSV IMPLEMENTATION.


  METHOD create.

* .csv format with ; delimiter

* Start of insertion # issue 1134 - Dateretention of cellstyles(issue #139)
    me->excel->add_static_styles( ).
* End of insertion # issue 1134 - Dateretention of cellstyles(issue #139)

    ep_excel = me->create_csv( ).

  ENDMETHOD.


  METHOD CREATE_CSV.

    TYPES: BEGIN OF LTY_FORMAT,
             CMPNAME  TYPE SEOCMPNAME,
             ATTVALUE TYPE SEOVALUE,
           END OF LTY_FORMAT.
    DATA: LT_FORMAT TYPE STANDARD TABLE OF LTY_FORMAT,
          LS_FORMAT LIKE LINE OF LT_FORMAT,
          LV_DATE   TYPE D,
          LV_TMP    TYPE STRING,
          LV_TIME   TYPE C LENGTH 8.

    DATA: LO_ITERATOR  TYPE REF TO ZCL_EXCEL_COLLECTION_ITERATOR,
          LO_WORKSHEET TYPE REF TO ZCL_EXCEL_WORKSHEET.

    DATA: LT_CELL_DATA TYPE ZEXCEL_T_CELL_DATA_UNSORTED,
          LV_ROW       TYPE I,
          LV_COL       TYPE I,
          LV_STRING    TYPE STRING,
          LC_VALUE     TYPE STRING,
          LV_ATTRNAME  TYPE SEOCMPNAME.

    DATA: LS_NUMFMT TYPE ZEXCEL_S_STYLE_NUMFMT,
          LO_STYLE  TYPE REF TO ZCL_EXCEL_STYLE.

    FIELD-SYMBOLS: <FS_SHEET_CONTENT> TYPE ZEXCEL_S_CELL_DATA.

* --- Retrieve supported cell format
    SELECT * INTO CORRESPONDING FIELDS OF TABLE LT_FORMAT
      FROM SEOCOMPODF
     WHERE CLSNAME  = 'ZCL_EXCEL_STYLE_NUMBER_FORMAT'
       AND TYPTYPE  = 1
       AND TYPE     = 'ZEXCEL_NUMBER_FORMAT'.

* --- Retrieve SAP date format
    CLEAR LS_FORMAT.
    SELECT DDTEXT INTO LS_FORMAT-ATTVALUE FROM DD07T WHERE DOMNAME    = 'XUDATFM'
                                                       AND DDLANGUAGE = SY-LANGU.
      LS_FORMAT-CMPNAME = 'DATE'.
      CONDENSE LS_FORMAT-ATTVALUE.
      CONCATENATE '''' LS_FORMAT-ATTVALUE '''' INTO LS_FORMAT-ATTVALUE.
      APPEND LS_FORMAT TO LT_FORMAT.
    ENDSELECT.


    LOOP AT LT_FORMAT INTO LS_FORMAT.
      TRANSLATE LS_FORMAT-ATTVALUE TO UPPER CASE.
      MODIFY LT_FORMAT FROM LS_FORMAT.
    ENDLOOP.


* STEP 1: Collect strings from the first worksheet
    LO_ITERATOR = EXCEL->GET_WORKSHEETS_ITERATOR( ).
    DATA: CURRENT_WORKSHEET_TITLE TYPE ZEXCEL_SHEET_TITLE.

    WHILE LO_ITERATOR->HAS_NEXT( ) EQ ABAP_TRUE.
      LO_WORKSHEET ?= LO_ITERATOR->GET_NEXT( ).

      IF WORKSHEET_NAME IS NOT INITIAL.
        CURRENT_WORKSHEET_TITLE = LO_WORKSHEET->GET_TITLE( ).
        CHECK CURRENT_WORKSHEET_TITLE = WORKSHEET_NAME.
      ELSE.
        IF WORKSHEET_INDEX IS INITIAL.
          WORKSHEET_INDEX = 1.
        ENDIF.
        CHECK WORKSHEET_INDEX = SY-INDEX.
      ENDIF.
      APPEND LINES OF LO_WORKSHEET->SHEET_CONTENT TO LT_CELL_DATA.
      EXIT. " Take first worksheet only
    ENDWHILE.

    DELETE LT_CELL_DATA WHERE CELL_FORMULA IS NOT INITIAL. " delete formula content

    SORT LT_CELL_DATA BY CELL_ROW
                         CELL_COLUMN.
    LV_ROW = 1.
    LV_COL = 1.
    CLEAR LV_STRING.
    LOOP AT LT_CELL_DATA ASSIGNING <FS_SHEET_CONTENT>.

*   --- Retrieve Cell Style format and data type
      CLEAR LS_NUMFMT.
      IF <FS_SHEET_CONTENT>-DATA_TYPE IS INITIAL AND <FS_SHEET_CONTENT>-CELL_STYLE IS NOT INITIAL.
        LO_ITERATOR = EXCEL->GET_STYLES_ITERATOR( ).
        WHILE LO_ITERATOR->HAS_NEXT( ) EQ ABAP_TRUE.
          LO_STYLE ?= LO_ITERATOR->GET_NEXT( ).
          CHECK LO_STYLE->GET_GUID( ) = <FS_SHEET_CONTENT>-CELL_STYLE.
          LS_NUMFMT     = LO_STYLE->NUMBER_FORMAT->GET_STRUCTURE( ).
          EXIT.
        ENDWHILE.
      ENDIF.
      IF <FS_SHEET_CONTENT>-DATA_TYPE IS INITIAL AND LS_NUMFMT IS NOT INITIAL.
        " determine data-type
        CLEAR LV_ATTRNAME.
        CONCATENATE '''' LS_NUMFMT-NUMFMT '''' INTO LS_NUMFMT-NUMFMT.
        TRANSLATE LS_NUMFMT-NUMFMT TO UPPER CASE.
        READ TABLE LT_FORMAT INTO LS_FORMAT WITH KEY ATTVALUE = LS_NUMFMT-NUMFMT.
        IF SY-SUBRC = 0.
          LV_ATTRNAME = LS_FORMAT-CMPNAME.
        ENDIF.

        IF LV_ATTRNAME IS NOT INITIAL.
          FIND FIRST OCCURRENCE OF 'DATETIME' IN LV_ATTRNAME.
          IF SY-SUBRC = 0.
            <FS_SHEET_CONTENT>-DATA_TYPE = 'd'.
          ELSE.
            FIND FIRST OCCURRENCE OF 'TIME' IN LV_ATTRNAME.
            IF SY-SUBRC = 0.
              <FS_SHEET_CONTENT>-DATA_TYPE = 't'.
            ELSE.
              FIND FIRST OCCURRENCE OF 'DATE' IN LV_ATTRNAME.
              IF SY-SUBRC = 0.
                <FS_SHEET_CONTENT>-DATA_TYPE = 'd'.
              ELSE.
                FIND FIRST OCCURRENCE OF 'CURRENCY' IN LV_ATTRNAME.
                IF SY-SUBRC = 0.
                  <FS_SHEET_CONTENT>-DATA_TYPE = 'n'.
                ELSE.
                  FIND FIRST OCCURRENCE OF 'NUMBER' IN LV_ATTRNAME.
                  IF SY-SUBRC = 0.
                    <FS_SHEET_CONTENT>-DATA_TYPE = 'n'.
                  ELSE.
                    FIND FIRST OCCURRENCE OF 'PERCENTAGE' IN LV_ATTRNAME.
                    IF SY-SUBRC = 0.
                      <FS_SHEET_CONTENT>-DATA_TYPE = 'n'.
                    ENDIF. " Purcentage
                  ENDIF. " Number
                ENDIF. " Currency
              ENDIF. " Date
            ENDIF. " TIME
          ENDIF. " DATETIME
        ENDIF. " lv_attrname IS NOT INITIAL.
      ENDIF. " <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.

* --- Add empty rows
      WHILE LV_ROW < <FS_SHEET_CONTENT>-CELL_ROW.
        CONCATENATE LV_STRING ZCL_EXCEL_WRITER_CSV=>EOL INTO LV_STRING.
        LV_ROW = LV_ROW + 1.
        LV_COL = 1.
      ENDWHILE.

* --- Add empty columns
      WHILE LV_COL < <FS_SHEET_CONTENT>-CELL_COLUMN.
        CONCATENATE LV_STRING ZCL_EXCEL_WRITER_CSV=>DELIMITER INTO LV_STRING.
        LV_COL = LV_COL + 1.
      ENDWHILE.

* ----- Use format to determine the data type and display format.
      CASE <FS_SHEET_CONTENT>-DATA_TYPE.

        WHEN 'd' OR 'D'.
          IF <FS_SHEET_CONTENT>-CELL_VALUE IS INITIAL AND INITIAL_EXT_DATE <> C_DEFAULT.
            LC_VALUE = INITIAL_EXT_DATE.
          ELSE.
            LC_VALUE = ZCL_EXCEL_COMMON=>EXCEL_STRING_TO_DATE( IP_VALUE = <FS_SHEET_CONTENT>-CELL_VALUE ).
            TRY.
                LV_DATE = LC_VALUE.
                CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
                  EXPORTING
                    DATE_INTERNAL            = LV_DATE
                  IMPORTING
                    DATE_EXTERNAL            = LV_TMP
                  EXCEPTIONS
                    DATE_INTERNAL_IS_INVALID = 1
                    OTHERS                   = 2.
                IF SY-SUBRC = 0.
                  LC_VALUE = LV_TMP.
                ENDIF.

              CATCH CX_SY_CONVERSION_NO_NUMBER.

            ENDTRY.
          ENDIF.

        WHEN 't' OR 'T'.
          LC_VALUE = ZCL_EXCEL_COMMON=>EXCEL_STRING_TO_TIME( IP_VALUE = <FS_SHEET_CONTENT>-CELL_VALUE ).
          WRITE LC_VALUE TO LV_TIME USING EDIT MASK '__:__:__'.
          LC_VALUE = LV_TIME.
        WHEN OTHERS.
          LC_VALUE = <FS_SHEET_CONTENT>-CELL_VALUE.

      ENDCASE.

      CONCATENATE ZCL_EXCEL_WRITER_CSV=>ENCLOSURE ZCL_EXCEL_WRITER_CSV=>ENCLOSURE INTO LV_TMP.
      CONDENSE LV_TMP.
      REPLACE ALL OCCURRENCES OF ZCL_EXCEL_WRITER_CSV=>ENCLOSURE IN LC_VALUE WITH LV_TMP.

      FIND FIRST OCCURRENCE OF ZCL_EXCEL_WRITER_CSV=>DELIMITER IN LC_VALUE.
      IF SY-SUBRC = 0.
        CONCATENATE LV_STRING ZCL_EXCEL_WRITER_CSV=>ENCLOSURE LC_VALUE ZCL_EXCEL_WRITER_CSV=>ENCLOSURE INTO LV_STRING.
      ELSE.
        CONCATENATE LV_STRING LC_VALUE INTO LV_STRING.
      ENDIF.

    ENDLOOP.

    CLEAR EP_CONTENT.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        TEXT   = LV_STRING
      IMPORTING
        BUFFER = EP_CONTENT
      EXCEPTIONS
        FAILED = 1
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


  METHOD SET_INITIAL_EXT_DATE.
    INITIAL_EXT_DATE = IP_VALUE.
  ENDMETHOD.
ENDCLASS.
