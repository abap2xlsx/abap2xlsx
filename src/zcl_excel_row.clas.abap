CLASS zcl_excel_row DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS constructor
      IMPORTING
        !ip_index TYPE int4 DEFAULT 0 .
    METHODS get_collapsed
      IMPORTING
        !io_worksheet      TYPE REF TO zcl_excel_worksheet OPTIONAL
      RETURNING
        VALUE(r_collapsed) TYPE abap_bool .
    METHODS get_outline_level
      IMPORTING
        !io_worksheet          TYPE REF TO zcl_excel_worksheet OPTIONAL
      RETURNING
        VALUE(r_outline_level) TYPE int4 .
    METHODS get_row_height
      RETURNING
        VALUE(r_row_height) TYPE f .
    METHODS get_custom_height
      RETURNING
        VALUE(r_custom_height) TYPE abap_bool .
    METHODS get_row_index
      RETURNING
        VALUE(r_row_index) TYPE int4 .
    METHODS get_visible
      IMPORTING
        !io_worksheet    TYPE REF TO zcl_excel_worksheet OPTIONAL
      RETURNING
        VALUE(r_visible) TYPE abap_bool .
    METHODS get_xf_index
      RETURNING
        VALUE(r_xf_index) TYPE int4 .
    METHODS set_collapsed
      IMPORTING
        !ip_collapsed TYPE abap_bool .
    METHODS set_outline_level
      IMPORTING
        !ip_outline_level TYPE int4
      RAISING
        zcx_excel .
    METHODS set_row_height
      IMPORTING
        !ip_row_height    TYPE simple
        !ip_custom_height TYPE abap_bool DEFAULT abap_true
      RAISING
        zcx_excel .
    METHODS set_row_index
      IMPORTING
        !ip_index TYPE int4 .
    METHODS set_visible
      IMPORTING
        !ip_visible TYPE abap_bool .
    METHODS set_xf_index
      IMPORTING
        !ip_xf_index TYPE int4 .
*"* protected components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA row_index TYPE int4 .
    DATA row_height TYPE f .
    DATA visible TYPE abap_bool .
    DATA outline_level TYPE int4 VALUE 0. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
    DATA collapsed TYPE abap_bool .
    DATA xf_index TYPE int4 .
    DATA custom_height TYPE abap_bool .
ENDCLASS.



CLASS zcl_excel_row IMPLEMENTATION.


  METHOD constructor.
    " Initialise values
    me->row_index    = ip_index.
    me->row_height   = -1.
    me->visible     = abap_true.
    me->outline_level  = 0.
    me->collapsed   = abap_false.

    " set row dimension as unformatted by default
    me->xf_index = 0.
    me->custom_height = abap_false.
  ENDMETHOD.


  METHOD get_collapsed.

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


  METHOD get_outline_level.

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


  METHOD get_row_height.
    r_row_height = me->row_height.
  ENDMETHOD.


  METHOD get_custom_height.
    r_custom_height = me->custom_height.
  ENDMETHOD.


  METHOD get_row_index.
    r_row_index = me->row_index.
  ENDMETHOD.


  METHOD get_visible.

    DATA: lt_row_outlines TYPE zcl_excel_worksheet=>mty_ts_outlines_row.
    FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF lt_row_outlines.

    r_visible = me->visible.
    CHECK r_visible = abap_true.  " Currently visible --> but maybe the new outline methodology will hide it implicitly
    CHECK io_worksheet IS BOUND.  " But we have to see the worksheet to make sure

    lt_row_outlines = io_worksheet->get_row_outlines( ).
    LOOP AT lt_row_outlines ASSIGNING <ls_row_outline> WHERE row_from  <= me->row_index
                                                         AND row_to    >= me->row_index
                                                         AND collapsed =  abap_true.      " row is in a collapsed outline --> not visible
      CLEAR r_visible.
      RETURN. " one hit is enough to ensure invisibility

    ENDLOOP.

  ENDMETHOD.


  METHOD get_xf_index.
    r_xf_index = me->xf_index.
  ENDMETHOD.


  METHOD set_collapsed.
    me->collapsed = ip_collapsed.
  ENDMETHOD.


  METHOD set_outline_level.
    IF   ip_outline_level < 0
      OR ip_outline_level > 7.

      zcx_excel=>raise_text( 'Outline level must range between 0 and 7.' ).

    ENDIF.
    me->outline_level = ip_outline_level.
  ENDMETHOD.


  METHOD set_row_height.
    DATA: height TYPE f.
    TRY.
        height = ip_row_height.
        IF height <= 0.
          zcx_excel=>raise_text( 'Please supply a positive number as row-height' ).
        ENDIF.
        me->row_height = ip_row_height.
      CATCH cx_sy_conversion_no_number.
        zcx_excel=>raise_text( 'Unable to interpret ip_row_height as number' ).
    ENDTRY.
    me->custom_height = ip_custom_height.
  ENDMETHOD.


  METHOD set_row_index.
    me->row_index = ip_index.
  ENDMETHOD.


  METHOD set_visible.
    me->visible = ip_visible.
  ENDMETHOD.


  METHOD set_xf_index.
    me->xf_index = ip_xf_index.
  ENDMETHOD.
ENDCLASS.
