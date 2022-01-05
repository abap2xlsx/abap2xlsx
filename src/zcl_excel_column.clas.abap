CLASS zcl_excel_column DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS constructor
      IMPORTING
        !ip_index     TYPE zexcel_cell_column_alpha
        !ip_worksheet TYPE REF TO zcl_excel_worksheet
        !ip_excel     TYPE REF TO zcl_excel
      RAISING
        zcx_excel .
    METHODS get_auto_size
      RETURNING
        VALUE(r_auto_size) TYPE abap_bool .
    METHODS get_collapsed
      RETURNING
        VALUE(r_collapsed) TYPE abap_bool .
    METHODS get_column_index
      RETURNING
        VALUE(r_column_index) TYPE int4 .
    METHODS get_outline_level
      RETURNING
        VALUE(r_outline_level) TYPE int4 .
    METHODS get_visible
      RETURNING
        VALUE(r_visible) TYPE abap_bool .
    METHODS get_width
      RETURNING
        VALUE(r_width) TYPE f .
    METHODS get_xf_index
      RETURNING
        VALUE(r_xf_index) TYPE int4 .
    METHODS set_auto_size
      IMPORTING
        !ip_auto_size    TYPE abap_bool
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_collapsed
      IMPORTING
        !ip_collapsed    TYPE abap_bool
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_column_index
      IMPORTING
        !ip_index        TYPE zexcel_cell_column_alpha
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS set_outline_level
      IMPORTING
        !ip_outline_level TYPE int4 .
    METHODS set_visible
      IMPORTING
        !ip_visible      TYPE abap_bool
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_width
      IMPORTING
        !ip_width        TYPE simple
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS set_xf_index
      IMPORTING
        !ip_xf_index     TYPE int4
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_column_style_by_guid
      IMPORTING
        !ip_style_guid TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    METHODS get_column_style_guid
      RETURNING
        VALUE(ep_style_guid) TYPE zexcel_cell_style
      RAISING
        zcx_excel .
*"* protected components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA column_index TYPE int4 .
    DATA width TYPE f .
    DATA auto_size TYPE abap_bool .
    DATA visible TYPE abap_bool .
    DATA outline_level TYPE int4 .
    DATA collapsed TYPE abap_bool .
    DATA xf_index TYPE int4 .
    DATA style_guid TYPE zexcel_cell_style .
    DATA excel TYPE REF TO zcl_excel .
    DATA worksheet TYPE REF TO zcl_excel_worksheet .
ENDCLASS.



CLASS zcl_excel_column IMPLEMENTATION.


  METHOD constructor.
    me->column_index = zcl_excel_common=>convert_column2int( ip_index ).
    me->width         = -1.
    me->auto_size     = abap_false.
    me->visible       = abap_true.
    me->outline_level = 0.
    me->collapsed     = abap_false.
    me->excel         = ip_excel.        "ins issue #157 - Allow Style for columns
    me->worksheet     = ip_worksheet.    "ins issue #157 - Allow Style for columns

    " set default index to cellXf
    me->xf_index = 0.

  ENDMETHOD.


  METHOD get_auto_size.
    r_auto_size = me->auto_size.
  ENDMETHOD.


  METHOD get_collapsed.
    r_collapsed = me->collapsed.
  ENDMETHOD.


  METHOD get_column_index.
    r_column_index = me->column_index.
  ENDMETHOD.


  METHOD get_column_style_guid.
    IF me->style_guid IS NOT INITIAL.
      ep_style_guid = me->style_guid.
    ELSE.
      ep_style_guid = me->worksheet->zif_excel_sheet_properties~get_style( ).
    ENDIF.
  ENDMETHOD.


  METHOD get_outline_level.
    r_outline_level = me->outline_level.
  ENDMETHOD.


  METHOD get_visible.
    r_visible = me->visible.
  ENDMETHOD.


  METHOD get_width.
    r_width = me->width.
  ENDMETHOD.


  METHOD get_xf_index.
    r_xf_index = me->xf_index.
  ENDMETHOD.


  METHOD set_auto_size.
    me->auto_size = ip_auto_size.
    io_column = me.
  ENDMETHOD.


  METHOD set_collapsed.
    me->collapsed = ip_collapsed.
    io_column = me.
  ENDMETHOD.


  METHOD set_column_index.
    me->column_index = zcl_excel_common=>convert_column2int( ip_index ).
    io_column = me.
  ENDMETHOD.


  METHOD set_column_style_by_guid.
    DATA: stylemapping TYPE zexcel_s_stylemapping.

    IF me->excel IS NOT BOUND.
      zcx_excel=>raise_text( 'Internal error - reference to ZCL_EXCEL not bound' ).
    ENDIF.
    TRY.
        stylemapping = me->excel->get_style_to_guid( ip_style_guid ).
        me->style_guid = stylemapping-guid.

      CATCH zcx_excel .
        RETURN.  " leave as is in case of error
    ENDTRY.

  ENDMETHOD.


  METHOD set_outline_level.
    me->outline_level = ip_outline_level.
  ENDMETHOD.


  METHOD set_visible.
    me->visible = ip_visible.
    io_column = me.
  ENDMETHOD.


  METHOD set_width.
    TRY.
        me->width = ip_width.
        io_column = me.
      CATCH cx_sy_conversion_no_number.
        zcx_excel=>raise_text( 'Unable to interpret width as number' ).
    ENDTRY.
  ENDMETHOD.


  METHOD set_xf_index.
    me->xf_index = ip_xf_index.
    io_column = me.
  ENDMETHOD.
ENDCLASS.
