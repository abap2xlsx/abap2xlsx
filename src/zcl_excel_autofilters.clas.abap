class ZCL_EXCEL_AUTOFILTERS definition
  public
  final
  create public .

public section.
*"* public components of class ZCL_EXCEL_AUTOFILTERS
*"* do not include other source files here!!!
  type-pools ABAP .

  constants C_AUTOFILTER type STRING value '_xlnm._FilterDatabase'. "#EC NOTEXT

  methods ADD
    importing
      !IO_SHEET type ref to ZCL_EXCEL_WORKSHEET
    returning
      value(RO_AUTOFILTER) type ref to ZCL_EXCEL_AUTOFILTER
    raising
      ZCX_EXCEL .
  methods CLEAR .
  methods GET
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET optional
      !I_SHEET_GUID type UUID optional
    returning
      value(RO_AUTOFILTER) type ref to ZCL_EXCEL_AUTOFILTER .
  methods IS_EMPTY
    returning
      value(R_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IO_SHEET type ANY .
  methods SIZE
    returning
      value(R_SIZE) type I .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
protected section.
private section.

  types:
    BEGIN OF ts_autofilter,
      worksheet  TYPE REF TO zcl_excel_worksheet,
      autofilter TYPE REF TO zcl_excel_autofilter,
    END OF ts_autofilter .
  types:
    tt_autofilters TYPE HASHED TABLE OF ts_autofilter WITH UNIQUE KEY worksheet .

  data MT_AUTOFILTERS type TT_AUTOFILTERS .
ENDCLASS.



CLASS ZCL_EXCEL_AUTOFILTERS IMPLEMENTATION.


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
