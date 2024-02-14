CLASS zcl_excel_writer_csv DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WRITER_2007
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
    CLASS-METHODS set_initial_ext_date
      IMPORTING
        !ip_value TYPE char10 DEFAULT 'DEFAULT' .
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA excel TYPE REF TO zcl_excel .
    CLASS-DATA delimiter TYPE c VALUE ';' ##NO_TEXT.
    CLASS-DATA enclosure TYPE c VALUE '"' ##NO_TEXT.
    CLASS-DATA:
      EOL type C length 2 value CL_ABAP_CHAR_UTILITIES=>CR_LF ##NO_TEXT.
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


  method CREATE.

* .csv format with ; delimiter

* Start of insertion # issue 1134 - Dateretention of cellstyles(issue #139)
    ME->EXCEL->ADD_STATIC_STYLES( ).
* End of insertion # issue 1134 - Dateretention of cellstyles(issue #139)

    EP_EXCEL = ME->CREATE_CSV( ).

  endmethod.


  method CREATE_CSV.

    types: begin of LTY_FORMAT,
             CMPNAME  type SEOCMPNAME,
             ATTVALUE type SEOVALUE,
           end of LTY_FORMAT.
    data: LT_FORMAT type standard table of LTY_FORMAT,
          LS_FORMAT like line of LT_FORMAT,
          LV_DATE   type D,
          LV_TMP    type STRING,
          LV_TIME   type C length 8.

    data: LO_ITERATOR  type ref to ZCL_EXCEL_COLLECTION_ITERATOR,
          LO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET.

    data: LT_CELL_DATA type ZEXCEL_T_CELL_DATA_UNSORTED,
          LV_ROW       type I,
          LV_COL       type I,
          LV_STRING    type STRING,
          LC_VALUE     type STRING,
          LV_ATTRNAME  type SEOCMPNAME.

    data: LS_NUMFMT type ZEXCEL_S_STYLE_NUMFMT,
          LO_STYLE  type ref to ZCL_EXCEL_STYLE.

    field-symbols: <FS_SHEET_CONTENT> type ZEXCEL_S_CELL_DATA.

* --- Retrieve supported cell format
    select * into corresponding fields of table LT_FORMAT
      from SEOCOMPODF
     where CLSNAME  = 'ZCL_EXCEL_STYLE_NUMBER_FORMAT'
       and TYPTYPE  = 1
       and TYPE     = 'ZEXCEL_NUMBER_FORMAT'.

* --- Retrieve SAP date format
    clear LS_FORMAT.
    select DDTEXT into LS_FORMAT-ATTVALUE from DD07T where DOMNAME    = 'XUDATFM'
                                                       and DDLANGUAGE = SY-LANGU.
      LS_FORMAT-CMPNAME = 'DATE'.
      condense LS_FORMAT-ATTVALUE.
      concatenate '''' LS_FORMAT-ATTVALUE '''' into LS_FORMAT-ATTVALUE.
      append LS_FORMAT to LT_FORMAT.
    endselect.


    loop at LT_FORMAT into LS_FORMAT.
      translate LS_FORMAT-ATTVALUE to upper case.
      modify LT_FORMAT from LS_FORMAT.
    endloop.


* STEP 1: Collect strings from the first worksheet
    LO_ITERATOR = EXCEL->GET_WORKSHEETS_ITERATOR( ).
    data: CURRENT_WORKSHEET_TITLE type ZEXCEL_SHEET_TITLE.

    while LO_ITERATOR->HAS_NEXT( ) eq ABAP_TRUE.
      LO_WORKSHEET ?= LO_ITERATOR->GET_NEXT( ).

      if WORKSHEET_NAME is not initial.
        CURRENT_WORKSHEET_TITLE = LO_WORKSHEET->GET_TITLE( ).
        check CURRENT_WORKSHEET_TITLE = WORKSHEET_NAME.
      else.
        if WORKSHEET_INDEX is initial.
          WORKSHEET_INDEX = 1.
        endif.
        check WORKSHEET_INDEX = SY-INDEX.
      endif.
      append lines of LO_WORKSHEET->SHEET_CONTENT to LT_CELL_DATA.
      exit. " Take first worksheet only
    endwhile.

    delete LT_CELL_DATA where CELL_FORMULA is not initial. " delete formula content

    sort LT_CELL_DATA by CELL_ROW
                         CELL_COLUMN.
    LV_ROW = 1.
    LV_COL = 1.
    clear LV_STRING.
    loop at LT_CELL_DATA assigning <FS_SHEET_CONTENT>.

*   --- Retrieve Cell Style format and data type
      clear LS_NUMFMT.
      if <FS_SHEET_CONTENT>-DATA_TYPE is initial and <FS_SHEET_CONTENT>-CELL_STYLE is not initial.
        LO_ITERATOR = EXCEL->GET_STYLES_ITERATOR( ).
        while LO_ITERATOR->HAS_NEXT( ) eq ABAP_TRUE.
          LO_STYLE ?= LO_ITERATOR->GET_NEXT( ).
          check LO_STYLE->GET_GUID( ) = <FS_SHEET_CONTENT>-CELL_STYLE.
          LS_NUMFMT     = LO_STYLE->NUMBER_FORMAT->GET_STRUCTURE( ).
          exit.
        endwhile.
      endif.
      if <FS_SHEET_CONTENT>-DATA_TYPE is initial and LS_NUMFMT is not initial.
        " determine data-type
        clear LV_ATTRNAME.
        concatenate '''' LS_NUMFMT-NUMFMT '''' into LS_NUMFMT-NUMFMT.
        translate LS_NUMFMT-NUMFMT to upper case.
        read table LT_FORMAT into LS_FORMAT with key ATTVALUE = LS_NUMFMT-NUMFMT.
        if SY-SUBRC = 0.
          LV_ATTRNAME = LS_FORMAT-CMPNAME.
        endif.

        if LV_ATTRNAME is not initial.
          find first occurrence of 'DATETIME' in LV_ATTRNAME.
          if SY-SUBRC = 0.
            <FS_SHEET_CONTENT>-DATA_TYPE = 'd'.
          else.
            find first occurrence of 'TIME' in LV_ATTRNAME.
            if SY-SUBRC = 0.
              <FS_SHEET_CONTENT>-DATA_TYPE = 't'.
            else.
              find first occurrence of 'DATE' in LV_ATTRNAME.
              if SY-SUBRC = 0.
                <FS_SHEET_CONTENT>-DATA_TYPE = 'd'.
              else.
                find first occurrence of 'CURRENCY' in LV_ATTRNAME.
                if SY-SUBRC = 0.
                  <FS_SHEET_CONTENT>-DATA_TYPE = 'n'.
                else.
                  find first occurrence of 'NUMBER' in LV_ATTRNAME.
                  if SY-SUBRC = 0.
                    <FS_SHEET_CONTENT>-DATA_TYPE = 'n'.
                  else.
                    find first occurrence of 'PERCENTAGE' in LV_ATTRNAME.
                    if SY-SUBRC = 0.
                      <FS_SHEET_CONTENT>-DATA_TYPE = 'n'.
                    endif. " Purcentage
                  endif. " Number
                endif. " Currency
              endif. " Date
            endif. " TIME
          endif. " DATETIME
        endif. " lv_attrname IS NOT INITIAL.
      endif. " <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.

* --- Add empty rows
      while LV_ROW < <FS_SHEET_CONTENT>-CELL_ROW.
        concatenate LV_STRING ZCL_EXCEL_WRITER_CSV=>EOL into LV_STRING.
        LV_ROW = LV_ROW + 1.
        LV_COL = 1.
      endwhile.

* --- Add empty columns
      while LV_COL < <FS_SHEET_CONTENT>-CELL_COLUMN.
        concatenate LV_STRING ZCL_EXCEL_WRITER_CSV=>DELIMITER into LV_STRING.
        LV_COL = LV_COL + 1.
      endwhile.

* ----- Use format to determine the data type and display format.
      case <FS_SHEET_CONTENT>-DATA_TYPE.

        when 'd' or 'D'.
          if <FS_SHEET_CONTENT>-CELL_VALUE is initial and INITIAL_EXT_DATE <> C_DEFAULT.
            LC_VALUE = INITIAL_EXT_DATE.
          else.
            LC_VALUE = ZCL_EXCEL_COMMON=>EXCEL_STRING_TO_DATE( IP_VALUE = <FS_SHEET_CONTENT>-CELL_VALUE ).
            try.
                LV_DATE = LC_VALUE.
                call function 'CONVERT_DATE_TO_EXTERNAL'
                  exporting
                    DATE_INTERNAL            = LV_DATE
                  importing
                    DATE_EXTERNAL            = LV_TMP
                  exceptions
                    DATE_INTERNAL_IS_INVALID = 1
                    others                   = 2.
                if SY-SUBRC = 0.
                  LC_VALUE = LV_TMP.
                endif.

              catch CX_SY_CONVERSION_NO_NUMBER.

            endtry.
          endif.

        when 't' or 'T'.
          LC_VALUE = ZCL_EXCEL_COMMON=>EXCEL_STRING_TO_TIME( IP_VALUE = <FS_SHEET_CONTENT>-CELL_VALUE ).
          write LC_VALUE to LV_TIME using edit mask '__:__:__'.
          LC_VALUE = LV_TIME.
        when others.
          LC_VALUE = <FS_SHEET_CONTENT>-CELL_VALUE.

      endcase.

      concatenate ZCL_EXCEL_WRITER_CSV=>ENCLOSURE ZCL_EXCEL_WRITER_CSV=>ENCLOSURE into LV_TMP.
      condense LV_TMP.
      replace all occurrences of ZCL_EXCEL_WRITER_CSV=>ENCLOSURE in LC_VALUE with LV_TMP.

      find first occurrence of ZCL_EXCEL_WRITER_CSV=>DELIMITER in LC_VALUE.
      if SY-SUBRC = 0.
        concatenate LV_STRING ZCL_EXCEL_WRITER_CSV=>ENCLOSURE LC_VALUE ZCL_EXCEL_WRITER_CSV=>ENCLOSURE into LV_STRING.
      else.
        concatenate LV_STRING LC_VALUE into LV_STRING.
      endif.

    endloop.

    clear EP_CONTENT.

    call function 'SCMS_STRING_TO_XSTRING'
      exporting
        TEXT   = LV_STRING
      importing
        BUFFER = EP_CONTENT
      exceptions
        FAILED = 1
        others = 2.

  endmethod.


  method SET_ACTIVE_SHEET_INDEX.
    clear WORKSHEET_NAME.
    WORKSHEET_INDEX = I_ACTIVE_WORKSHEET.
  endmethod.


  method SET_ACTIVE_SHEET_INDEX_BY_NAME.
    clear WORKSHEET_INDEX.
    WORKSHEET_NAME = I_WORKSHEET_NAME.
  endmethod.


  method SET_DELIMITER.
    DELIMITER = IP_VALUE.
  endmethod.


  method SET_ENCLOSURE.
    ZCL_EXCEL_WRITER_CSV=>ENCLOSURE = IP_VALUE.
  endmethod.


  method SET_ENDOFLINE.
    ZCL_EXCEL_WRITER_CSV=>EOL = IP_VALUE.
  endmethod.


  method ZIF_EXCEL_WRITER~WRITE_FILE.
    ME->EXCEL = IO_EXCEL.
    EP_FILE = ME->CREATE( ).
  endmethod.


  method SET_INITIAL_EXT_DATE.
    INITIAL_EXT_DATE = IP_VALUE.
  endmethod.
ENDCLASS.
