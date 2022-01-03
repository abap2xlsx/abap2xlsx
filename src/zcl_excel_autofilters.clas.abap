CLASS zcl_excel_autofilters DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_AUTOFILTERS
*"* do not include other source files here!!!

    CONSTANTS c_autofilter TYPE string VALUE '_xlnm._FilterDatabase'. "#EC NOTEXT

    METHODS add
      IMPORTING
        !io_sheet            TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ro_autofilter) TYPE REF TO zcl_excel_autofilter
      RAISING
        zcx_excel .
    METHODS clear .
    METHODS get
      IMPORTING
        !io_worksheet        TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ro_autofilter) TYPE REF TO zcl_excel_autofilter .
    METHODS is_empty
      RETURNING
        VALUE(r_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !io_sheet TYPE any .
    METHODS size
      RETURNING
        VALUE(r_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.

    TYPES:
      BEGIN OF ts_autofilter,
        worksheet  TYPE REF TO zcl_excel_worksheet,
        autofilter TYPE REF TO zcl_excel_autofilter,
      END OF ts_autofilter .
    TYPES:
      tt_autofilters TYPE HASHED TABLE OF ts_autofilter WITH UNIQUE KEY worksheet .

    DATA mt_autofilters TYPE tt_autofilters .
ENDCLASS.



CLASS zcl_excel_autofilters IMPLEMENTATION.


  METHOD add.

    DATA: ls_autofilter LIKE LINE OF me->mt_autofilters.

    FIELD-SYMBOLS: <ls_autofilter> LIKE LINE OF me->mt_autofilters.

    READ TABLE me->mt_autofilters ASSIGNING <ls_autofilter> WITH TABLE KEY worksheet = io_sheet.
    IF sy-subrc = 0.
      RAISE EXCEPTION TYPE zcx_excel. " adding another autofilter to sheet is not allowed
    ENDIF.

    CREATE OBJECT ro_autofilter
      EXPORTING
        io_sheet = io_sheet.

    ls_autofilter-worksheet  = io_sheet.
    ls_autofilter-autofilter = ro_autofilter.
    INSERT ls_autofilter INTO TABLE me->mt_autofilters.


  ENDMETHOD.


  METHOD clear.

    CLEAR me->mt_autofilters.

  ENDMETHOD.


  METHOD get.

    FIELD-SYMBOLS: <ls_autofilter> LIKE LINE OF me->mt_autofilters.

    READ TABLE me->mt_autofilters ASSIGNING <ls_autofilter> WITH TABLE KEY worksheet = io_worksheet.
    IF sy-subrc = 0.
      ro_autofilter = <ls_autofilter>-autofilter.
    ELSE.
      CLEAR ro_autofilter.
    ENDIF.

  ENDMETHOD.


  METHOD is_empty.
    IF me->mt_autofilters IS INITIAL.
      r_empty = abap_true.
    ENDIF.
  ENDMETHOD.


  METHOD remove.

    DATA: lo_worksheet  TYPE REF TO zcl_excel_worksheet.

    DELETE TABLE me->mt_autofilters WITH TABLE KEY worksheet = lo_worksheet.

  ENDMETHOD.


  METHOD size.
    DESCRIBE TABLE me->mt_autofilters LINES r_size.
  ENDMETHOD.
ENDCLASS.
