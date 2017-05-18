class ZCL_EXCEL_COLUMN definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!
public section.
  type-pools ABAP .

  methods CONSTRUCTOR
    importing
      !IP_INDEX type ZEXCEL_CELL_COLUMN_ALPHA
      !IP_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
      !IP_EXCEL type ref to ZCL_EXCEL .
  methods GET_AUTO_SIZE
    returning
      value(R_AUTO_SIZE) type ABAP_BOOL .
  methods GET_COLLAPSED
    returning
      value(R_COLLAPSED) type ABAP_BOOL .
  methods GET_COLUMN_INDEX
    returning
      value(R_COLUMN_INDEX) type INT4 .
  methods GET_OUTLINE_LEVEL
    returning
      value(R_OUTLINE_LEVEL) type INT4 .
  methods GET_VISIBLE
    returning
      value(R_VISIBLE) type ABAP_BOOL .
  methods GET_WIDTH
    returning
      value(R_WIDTH) type FLOAT .
  methods GET_XF_INDEX
    returning
      value(R_XF_INDEX) type INT4 .
  methods SET_AUTO_SIZE
    importing
      !IP_AUTO_SIZE type ABAP_BOOL
    returning
      value(IO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods SET_COLLAPSED
    importing
      !IP_COLLAPSED type ABAP_BOOL
    returning
      value(IO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods SET_COLUMN_INDEX
    importing
      !IP_INDEX type ZEXCEL_CELL_COLUMN_ALPHA
    returning
      value(IO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods SET_OUTLINE_LEVEL
    importing
      !IP_OUTLINE_LEVEL type INT4 .
  methods SET_VISIBLE
    importing
      !IP_VISIBLE type ABAP_BOOL
    returning
      value(IO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods SET_WIDTH
    importing
      !IP_WIDTH type SIMPLE
    returning
      value(IO_COLUMN) type ref to ZCL_EXCEL_COLUMN
    raising
      ZCX_EXCEL .
  methods SET_XF_INDEX
    importing
      !IP_XF_INDEX type INT4
    returning
      value(IO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods SET_COLUMN_STYLE_BY_GUID
    importing
      !IP_STYLE_GUID type ZEXCEL_CELL_STYLE
    raising
      ZCX_EXCEL .
  methods GET_COLUMN_STYLE_GUID
    returning
      value(EP_STYLE_GUID) type ZEXCEL_CELL_STYLE
    raising
      ZCX_EXCEL .
*"* protected components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!
private section.

  data COLUMN_INDEX type INT4 .
  data WIDTH type FLOAT .
  data AUTO_SIZE type ABAP_BOOL .
  data VISIBLE type ABAP_BOOL .
  data OUTLINE_LEVEL type INT4 .
  data COLLAPSED type ABAP_BOOL .
  data XF_INDEX type INT4 .
  data STYLE_GUID type ZEXCEL_CELL_STYLE .
  data EXCEL type ref to ZCL_EXCEL .
  data WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
ENDCLASS.



CLASS ZCL_EXCEL_COLUMN IMPLEMENTATION.


method CONSTRUCTOR.
  me->column_index = zcl_excel_common=>convert_column2int( ip_index ).
  me->width         = -1.
  me->auto_size     = abap_false.
  me->visible       = abap_true.
  me->outline_level	= 0.
  me->collapsed     = abap_false.
  me->excel         = ip_excel.        "ins issue #157 - Allow Style for columns
  me->worksheet     = ip_worksheet.    "ins issue #157 - Allow Style for columns

  " set default index to cellXf
  me->xf_index = 0.

  endmethod.


method GET_AUTO_SIZE.
  r_auto_size = me->auto_size.
  endmethod.


method GET_COLLAPSED.
  r_Collapsed = me->Collapsed.
  endmethod.


method GET_COLUMN_INDEX.
  r_column_index = me->column_index.
  endmethod.


method GET_COLUMN_STYLE_GUID.
  IF me->style_guid IS NOT INITIAL.
    ep_style_guid = me->style_guid.
  ELSE.
    ep_style_guid = me->worksheet->zif_excel_sheet_properties~get_style( ).
  ENDIF.
  endmethod.


method GET_OUTLINE_LEVEL.
  r_outline_level = me->outline_level.
  endmethod.


method GET_VISIBLE.
  r_Visible = me->Visible.
  endmethod.


method GET_WIDTH.
  r_WIDTH = me->WIDTH.
  endmethod.


method GET_XF_INDEX.
  r_xf_index = me->xf_index.
  endmethod.


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


method SET_COLUMN_STYLE_BY_GUID.
  DATA: stylemapping TYPE zexcel_s_stylemapping.

  IF me->excel IS NOT BOUND.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Internal error - reference to ZCL_EXCEL not bound'.
  ENDIF.
  TRY.
      stylemapping = me->excel->get_style_to_guid( ip_style_guid ).
      me->style_guid = stylemapping-guid.

    CATCH zcx_excel .
      EXIT.  " leave as is in case of error
  ENDTRY.

  endmethod.


method SET_OUTLINE_LEVEL.
  me->outline_level = ip_outline_level.
  endmethod.


METHOD set_visible.
  me->visible = ip_visible.
  io_column = me.
ENDMETHOD.


METHOD set_width.
  TRY.
      me->width = ip_width.
      io_column = me.
    CATCH cx_sy_conversion_no_number.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Unable to interpret width as number'.
  ENDTRY.
ENDMETHOD.


METHOD set_xf_index.
  me->xf_index = ip_xf_index.
  io_column = me.
ENDMETHOD.
ENDCLASS.
