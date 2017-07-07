class ZCL_EXCEL_ROW definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!
public section.
  type-pools ABAP .

  methods CONSTRUCTOR
    importing
      !IP_INDEX type INT4 default 0 .
  methods GET_COLLAPSED
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET optional
    returning
      value(R_COLLAPSED) type BOOLEAN .
  methods GET_OUTLINE_LEVEL
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET optional
    returning
      value(R_OUTLINE_LEVEL) type INT4 .
  methods GET_ROW_HEIGHT
    returning
      value(R_ROW_HEIGHT) type FLOAT .
  methods GET_ROW_INDEX
    returning
      value(R_ROW_INDEX) type INT4 .
  methods GET_VISIBLE
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET optional
    returning
      value(R_VISIBLE) type BOOLEAN .
  methods GET_XF_INDEX
    returning
      value(R_XF_INDEX) type INT4 .
  methods SET_COLLAPSED
    importing
      !IP_COLLAPSED type BOOLEAN .
  methods SET_OUTLINE_LEVEL
    importing
      !IP_OUTLINE_LEVEL type INT4
    raising
      ZCX_EXCEL .
  methods SET_ROW_HEIGHT
    importing
      !IP_ROW_HEIGHT type SIMPLE
    raising
      ZCX_EXCEL .
  methods SET_ROW_INDEX
    importing
      !IP_INDEX type INT4 .
  methods SET_VISIBLE
    importing
      !IP_VISIBLE type BOOLEAN .
  methods SET_XF_INDEX
    importing
      !IP_XF_INDEX type INT4 .
*"* protected components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!
private section.

  data ROW_INDEX type INT4 .
  data ROW_HEIGHT type FLOAT .
  data VISIBLE type BOOLEAN .
  data OUTLINE_LEVEL type INT4 value 0. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
  data COLLAPSED type BOOLEAN .
  data XF_INDEX type INT4 .
ENDCLASS.



CLASS ZCL_EXCEL_ROW IMPLEMENTATION.


method CONSTRUCTOR.
  " Initialise values
  me->row_index    = ip_index.
  me->row_height   = -1.
  me->visible     = abap_true.
  me->outline_level  = 0.
  me->collapsed   = abap_false.

  " set row dimension as unformatted by default
  me->xf_index = 0.
  endmethod.


METHOD GET_COLLAPSED.

  DATA: lt_row_outlines  TYPE zcl_excel_worksheet=>mty_ts_outlines_row,
        lv_previous_row  TYPE i,
        lv_following_row TYPE i.

  r_collapsed = me->collapsed.

  CHECK r_collapsed = abap_false.  " Maybe new method for outlines is being used
  CHECK io_worksheet IS BOUND.

* If an outline is collapsed ( even inside an outer outline ) the line following the last line
* of the group gets the flag "collapsed"
  IF io_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_off.
    lv_following_row = me->row_index + 1.
    lt_row_outlines = io_worksheet->get_row_outlines( ).
    READ TABLE lt_row_outlines TRANSPORTING NO FIELDS WITH KEY row_from  = lv_following_row " first line of an outline
                                                               collapsed = abap_true.       " that is collapsed
  ELSE.
    lv_previous_row = me->row_index - 1.
    lt_row_outlines = io_worksheet->get_row_outlines( ).
    READ TABLE lt_row_outlines TRANSPORTING NO FIELDS WITH KEY row_to    = lv_previous_row  " last line of an outline
                                                               collapsed = abap_true.       " that is collapsed
  ENDIF.
  CHECK sy-subrc = 0.  " ok - we found it
  r_collapsed = abap_true.


ENDMETHOD.


METHOD GET_OUTLINE_LEVEL.

  DATA: lt_row_outlines TYPE zcl_excel_worksheet=>mty_ts_outlines_row.
  FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF lt_row_outlines.

* if someone has set the outline level explicitly - just use that
  IF me->outline_level IS NOT INITIAL.
    r_outline_level = me->outline_level.
    RETURN.
  ENDIF.
* Maybe we can use the outline information in the worksheet
  CHECK io_worksheet IS BOUND.

  lt_row_outlines = io_worksheet->get_row_outlines( ).
  LOOP AT lt_row_outlines ASSIGNING <ls_row_outline> WHERE row_from <= me->row_index
                                                       AND row_to   >= me->row_index.

    ADD 1 TO r_outline_level.

  ENDLOOP.

ENDMETHOD.


method GET_ROW_HEIGHT.
  r_row_height = me->row_height.
  endmethod.


method GET_ROW_INDEX.
  r_row_index = me->row_index.
  endmethod.


METHOD GET_VISIBLE.

  DATA: lt_row_outlines TYPE zcl_excel_worksheet=>mty_ts_outlines_row.
  FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF lt_row_outlines.

  r_visible = me->visible.
  CHECK r_visible = 'X'.        " Currently visible --> but maybe the new outline methodology will hide it implicitly
  CHECK io_worksheet IS BOUND.  " But we have to see the worksheet to make sure

  lt_row_outlines = io_worksheet->get_row_outlines( ).
  LOOP AT lt_row_outlines ASSIGNING <ls_row_outline> WHERE row_from  <= me->row_index
                                                       AND row_to    >= me->row_index
                                                       AND collapsed =  abap_true.      " row is in a collapsed outline --> not visible
    CLEAR r_visible.
    RETURN. " one hit is enough to ensure invisibility

  ENDLOOP.

ENDMETHOD.


method GET_XF_INDEX.
  r_xf_index = me->xf_index.
  endmethod.


method SET_COLLAPSED.
  me->collapsed = ip_collapsed.
  endmethod.


method SET_OUTLINE_LEVEL.
  IF   ip_outline_level < 0
    OR ip_outline_level > 7.

    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Outline level must range between 0 and 7.'.

  ENDIF.
  me->outline_level = ip_outline_level.
  endmethod.


method SET_ROW_HEIGHT.
  TRY.
      me->row_height = ip_row_height.
    CATCH cx_sy_conversion_no_number.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Unable to interpret ip_row_height as number'.
  ENDTRY.
  endmethod.


method SET_ROW_INDEX.
  me->row_index = ip_index.
  endmethod.


method SET_VISIBLE.
  me->visible = ip_visible.
  endmethod.


method SET_XF_INDEX.
  me->XF_INDEX = ip_XF_INDEX.
  endmethod.
ENDCLASS.
