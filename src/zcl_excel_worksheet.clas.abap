class ZCL_EXCEL_WORKSHEET definition
  public
  create public .

public section.
*"* public components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
  type-pools ABAP .
  type-pools SLIS .
  type-pools SOI .

  interfaces ZIF_EXCEL_SHEET_PRINTSETTINGS .
  interfaces ZIF_EXCEL_SHEET_PROPERTIES .
  interfaces ZIF_EXCEL_SHEET_PROTECTION .
  interfaces ZIF_EXCEL_SHEET_VBA_PROJECT .

  types:
    begin of  MTY_S_OUTLINE_ROW,
          row_from  type i,
          row_to    type i,
          collapsed type abap_bool,
        end of mty_s_outline_row .
  types:
    MTY_TS_OUTLINES_ROW type sorted table of MTY_S_OUTLINE_ROW with unique key row_from row_to .

  constants C_BREAK_COLUMN type ZEXCEL_BREAK value 2. "#EC NOTEXT
  constants C_BREAK_NONE type ZEXCEL_BREAK value 0. "#EC NOTEXT
  constants C_BREAK_ROW type ZEXCEL_BREAK value 1. "#EC NOTEXT
  data EXCEL type ref to ZCL_EXCEL read-only .
  data PRINT_GRIDLINES type ZEXCEL_PRINT_GRIDLINES read-only value ABAP_FALSE. "#EC NOTEXT
  data SHEET_CONTENT type ZEXCEL_T_CELL_DATA .
  data SHEET_SETUP type ref to ZCL_EXCEL_SHEET_SETUP .
  data SHOW_GRIDLINES type ZEXCEL_SHOW_GRIDLINES read-only value ABAP_TRUE. "#EC NOTEXT
  data SHOW_ROWCOLHEADERS type ZEXCEL_SHOW_GRIDLINES read-only value ABAP_TRUE. "#EC NOTEXT
  data STYLES type ZEXCEL_T_SHEET_STYLE .
  data TABCOLOR type ZEXCEL_S_TABCOLOR read-only .

  methods ADD_DRAWING
    importing
      !IP_DRAWING type ref to ZCL_EXCEL_DRAWING .
  methods ADD_NEW_COLUMN
    importing
      !IP_COLUMN type SIMPLE
    returning
      value(EO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods ADD_NEW_STYLE_COND
    returning
      value(EO_STYLE_COND) type ref to ZCL_EXCEL_STYLE_COND .
  methods ADD_NEW_DATA_VALIDATION
    returning
      value(EO_DATA_VALIDATION) type ref to ZCL_EXCEL_DATA_VALIDATION .
  methods ADD_NEW_RANGE
    returning
      value(EO_RANGE) type ref to ZCL_EXCEL_RANGE .
  methods ADD_NEW_ROW
    importing
      !IP_ROW type SIMPLE
    returning
      value(EO_ROW) type ref to ZCL_EXCEL_ROW .
  methods BIND_ALV
    importing
      !IO_ALV type ref to OBJECT
      !IT_TABLE type STANDARD TABLE
      !I_TOP type I default 1
      !I_LEFT type I default 1
      !TABLE_STYLE type ZEXCEL_TABLE_STYLE optional
      !I_TABLE type ABAP_BOOL default ABAP_TRUE
    raising
      ZCX_EXCEL .
  methods BIND_ALV_OLE2
    importing
      !I_DOCUMENT_URL type CHAR255 default SPACE
      !I_XLS type C default SPACE
      !I_SAVE_PATH type STRING
      !IO_ALV type ref to CL_GUI_ALV_GRID
      !IT_LISTHEADER type SLIS_T_LISTHEADER optional
      !I_TOP type I default 1
      !I_LEFT type I default 1
      !I_COLUMNS_HEADER type C default 'X'
      !I_COLUMNS_AUTOFIT type C default 'X'
      !I_FORMAT_COL_HEADER type SOI_FORMAT_ITEM optional
      !I_FORMAT_SUBTOTAL type SOI_FORMAT_ITEM optional
      !I_FORMAT_TOTAL type SOI_FORMAT_ITEM optional
    exceptions
      MISS_GUIDE
      EX_TRANSFER_KKBLO_ERROR
      FATAL_ERROR
      INV_DATA_RANGE
      DIM_MISMATCH_VKEY
      DIM_MISMATCH_SEMA
      ERROR_IN_SEMA .
  methods BIND_TABLE
    importing
      !IP_TABLE type STANDARD TABLE
      !IT_FIELD_CATALOG type ZEXCEL_T_FIELDCATALOG optional
      !IS_TABLE_SETTINGS type ZEXCEL_S_TABLE_SETTINGS optional
      value(IV_DEFAULT_DESCR) type C optional
    exporting
      !ES_TABLE_SETTINGS type ZEXCEL_S_TABLE_SETTINGS
    raising
      ZCX_EXCEL .
  methods CALCULATE_COLUMN_WIDTHS
    raising
      ZCX_EXCEL .
  methods CHANGE_CELL_STYLE
    importing
      !IP_CELL type ZEXCEL_CELL_REFERENCE optional
      !IP_COLUMN type SIMPLE optional
      !IP_ROW type ZEXCEL_CELL_ROW optional
      !IP_COLUMN_TO type SIMPLE optional
      !IP_ROW_TO type ZEXCEL_CELL_ROW optional
      !IP_COMPLETE type ZEXCEL_S_CSTYLE_COMPLETE optional
      !IP_XCOMPLETE type ZEXCEL_S_CSTYLEX_COMPLETE optional
      !IP_FONT type ZEXCEL_S_CSTYLE_FONT optional
      !IP_XFONT type ZEXCEL_S_CSTYLEX_FONT optional
      !IP_FILL type ZEXCEL_S_CSTYLE_FILL optional
      !IP_XFILL type ZEXCEL_S_CSTYLEX_FILL optional
      !IP_BORDERS type ZEXCEL_S_CSTYLE_BORDERS optional
      !IP_XBORDERS type ZEXCEL_S_CSTYLEX_BORDERS optional
      !IP_ALIGNMENT type ZEXCEL_S_CSTYLE_ALIGNMENT optional
      !IP_XALIGNMENT type ZEXCEL_S_CSTYLEX_ALIGNMENT optional
      !IP_NUMBER_FORMAT_FORMAT_CODE type ZEXCEL_NUMBER_FORMAT optional
      !IP_PROTECTION type ZEXCEL_S_CSTYLE_PROTECTION optional
      !IP_XPROTECTION type ZEXCEL_S_CSTYLEX_PROTECTION optional
      !IP_FONT_BOLD type FLAG optional
      !IP_FONT_COLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_FONT_COLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_FONT_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_FONT_COLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_FONT_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_FONT_FAMILY type ZEXCEL_STYLE_FONT_FAMILY optional
      !IP_FONT_ITALIC type FLAG optional
      !IP_FONT_NAME type ZEXCEL_STYLE_FONT_NAME optional
      !IP_FONT_SCHEME type ZEXCEL_STYLE_FONT_SCHEME optional
      !IP_FONT_SIZE type ZEXCEL_STYLE_FONT_SIZE optional
      !IP_FONT_STRIKETHROUGH type FLAG optional
      !IP_FONT_UNDERLINE type FLAG optional
      !IP_FONT_UNDERLINE_MODE type ZEXCEL_STYLE_FONT_UNDERLINE optional
      !IP_FILL_FILLTYPE type ZEXCEL_FILL_TYPE optional
      !IP_FILL_ROTATION type ZEXCEL_ROTATION optional
      !IP_FILL_FGCOLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_FILL_FGCOLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_FILL_FGCOLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_FILL_FGCOLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_FILL_FGCOLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_FILL_BGCOLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_FILL_BGCOLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_FILL_BGCOLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_FILL_BGCOLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_FILL_BGCOLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_BORDERS_ALLBORDERS type ZEXCEL_S_CSTYLE_BORDER optional
      !IP_FILL_GRADTYPE_TYPE type ZEXCEL_S_GRADIENT_TYPE-TYPE optional
      !IP_FILL_GRADTYPE_DEGREE type ZEXCEL_S_GRADIENT_TYPE-DEGREE optional
      !IP_XBORDERS_ALLBORDERS type ZEXCEL_S_CSTYLEX_BORDER optional
      !IP_BORDERS_DIAGONAL type ZEXCEL_S_CSTYLE_BORDER optional
      !IP_FILL_GRADTYPE_BOTTOM type ZEXCEL_S_GRADIENT_TYPE-BOTTOM optional
      !IP_FILL_GRADTYPE_TOP type ZEXCEL_S_GRADIENT_TYPE-TOP optional
      !IP_XBORDERS_DIAGONAL type ZEXCEL_S_CSTYLEX_BORDER optional
      !IP_BORDERS_DIAGONAL_MODE type ZEXCEL_DIAGONAL optional
      !IP_FILL_GRADTYPE_RIGHT type ZEXCEL_S_GRADIENT_TYPE-RIGHT optional
      !IP_BORDERS_DOWN type ZEXCEL_S_CSTYLE_BORDER optional
      !IP_FILL_GRADTYPE_LEFT type ZEXCEL_S_GRADIENT_TYPE-LEFT optional
      !IP_FILL_GRADTYPE_POSITION1 type ZEXCEL_S_GRADIENT_TYPE-POSITION1 optional
      !IP_XBORDERS_DOWN type ZEXCEL_S_CSTYLEX_BORDER optional
      !IP_BORDERS_LEFT type ZEXCEL_S_CSTYLE_BORDER optional
      !IP_FILL_GRADTYPE_POSITION2 type ZEXCEL_S_GRADIENT_TYPE-POSITION2 optional
      !IP_FILL_GRADTYPE_POSITION3 type ZEXCEL_S_GRADIENT_TYPE-POSITION3 optional
      !IP_XBORDERS_LEFT type ZEXCEL_S_CSTYLEX_BORDER optional
      !IP_BORDERS_RIGHT type ZEXCEL_S_CSTYLE_BORDER optional
      !IP_XBORDERS_RIGHT type ZEXCEL_S_CSTYLEX_BORDER optional
      !IP_BORDERS_TOP type ZEXCEL_S_CSTYLE_BORDER optional
      !IP_XBORDERS_TOP type ZEXCEL_S_CSTYLEX_BORDER optional
      !IP_ALIGNMENT_HORIZONTAL type ZEXCEL_ALIGNMENT optional
      !IP_ALIGNMENT_VERTICAL type ZEXCEL_ALIGNMENT optional
      !IP_ALIGNMENT_TEXTROTATION type ZEXCEL_TEXT_ROTATION optional
      !IP_ALIGNMENT_WRAPTEXT type FLAG optional
      !IP_ALIGNMENT_SHRINKTOFIT type FLAG optional
      !IP_ALIGNMENT_INDENT type ZEXCEL_INDENT optional
      !IP_PROTECTION_HIDDEN type ZEXCEL_CELL_PROTECTION optional
      !IP_PROTECTION_LOCKED type ZEXCEL_CELL_PROTECTION optional
      !IP_BORDERS_ALLBORDERS_STYLE type ZEXCEL_BORDER optional
      !IP_BORDERS_ALLBORDERS_COLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_BORDERS_ALLBO_COLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_BORDERS_ALLBO_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_BORDERS_ALLBO_COLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_BORDERS_ALLBO_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_BORDERS_DIAGONAL_STYLE type ZEXCEL_BORDER optional
      !IP_BORDERS_DIAGONAL_COLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_BORDERS_DIAGONAL_COLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_BORDERS_DIAGONAL_COLOR_INDE type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_BORDERS_DIAGONAL_COLOR_THEM type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_BORDERS_DIAGONAL_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_BORDERS_DOWN_STYLE type ZEXCEL_BORDER optional
      !IP_BORDERS_DOWN_COLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_BORDERS_DOWN_COLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_BORDERS_DOWN_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_BORDERS_DOWN_COLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_BORDERS_DOWN_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_BORDERS_LEFT_STYLE type ZEXCEL_BORDER optional
      !IP_BORDERS_LEFT_COLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_BORDERS_LEFT_COLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_BORDERS_LEFT_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_BORDERS_LEFT_COLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_BORDERS_LEFT_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_BORDERS_RIGHT_STYLE type ZEXCEL_BORDER optional
      !IP_BORDERS_RIGHT_COLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_BORDERS_RIGHT_COLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_BORDERS_RIGHT_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_BORDERS_RIGHT_COLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_BORDERS_RIGHT_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
      !IP_BORDERS_TOP_STYLE type ZEXCEL_BORDER optional
      !IP_BORDERS_TOP_COLOR type ZEXCEL_S_STYLE_COLOR optional
      !IP_BORDERS_TOP_COLOR_RGB type ZEXCEL_STYLE_COLOR_ARGB optional
      !IP_BORDERS_TOP_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
      !IP_BORDERS_TOP_COLOR_THEME type ZEXCEL_STYLE_COLOR_THEME optional
      !IP_BORDERS_TOP_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
    returning
      value(EP_GUID) type ZEXCEL_CELL_STYLE
    raising
      ZCX_EXCEL .
  methods CONSTRUCTOR
    importing
      !IP_EXCEL type ref to ZCL_EXCEL
      !IP_TITLE type ZEXCEL_SHEET_TITLE optional
    raising
      ZCX_EXCEL .
  methods DELETE_MERGE
    importing
      !IP_CELL_COLUMN type SIMPLE optional
      !IP_CELL_ROW type ZEXCEL_CELL_ROW optional .
  methods DELETE_ROW_OUTLINE
    importing
      !IV_ROW_FROM type I
      !IV_ROW_TO type I
    raising
      ZCX_EXCEL .
  methods FREEZE_PANES
    importing
      !IP_NUM_COLUMNS type I optional
      !IP_NUM_ROWS type I optional
    raising
      ZCX_EXCEL .
  methods GET_ACTIVE_CELL
    returning
      value(EP_ACTIVE_CELL) type STRING
    raising
      ZCX_EXCEL .
  methods GET_CELL
    importing
      !IP_CELL type CSEQUENCE optional
      !IP_COLUMN type SIMPLE optional
      !IP_ROW type ZEXCEL_CELL_ROW optional
    exporting
      !EP_VALUE type ZEXCEL_CELL_VALUE
      !EP_RC type SYSUBRC
      !EP_STYLE type ref to ZCL_EXCEL_STYLE
      !EP_GUID type ZEXCEL_CELL_STYLE
      !EP_FORMULA type ZEXCEL_CELL_FORMULA
      !ET_RTF type ZEXCEL_T_RTF
    raising
      ZCX_EXCEL .
  methods GET_COLUMN
    importing
      !IP_COLUMN type SIMPLE
    returning
      value(EO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods GET_COLUMNS
    returning
      value(EO_COLUMNS) type ref to ZCL_EXCEL_COLUMNS .
  methods GET_COLUMNS_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_STYLE_COND_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_DATA_VALIDATIONS_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_DATA_VALIDATIONS_SIZE
    returning
      value(EP_SIZE) type I .
  methods GET_DEFAULT_COLUMN
    returning
      value(EO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
  methods GET_DEFAULT_EXCEL_DATE_FORMAT
    returning
      value(EP_DEFAULT_EXCEL_DATE_FORMAT) type ZEXCEL_NUMBER_FORMAT .
  methods GET_DEFAULT_EXCEL_TIME_FORMAT
    returning
      value(EP_DEFAULT_EXCEL_TIME_FORMAT) type ZEXCEL_NUMBER_FORMAT .
  methods GET_DEFAULT_ROW
    returning
      value(EO_ROW) type ref to ZCL_EXCEL_ROW .
  methods GET_DIMENSION_RANGE
    returning
      value(EP_DIMENSION_RANGE) type STRING
    raising
      ZCX_EXCEL .
  methods GET_DRAWINGS
    importing
      !IP_TYPE type ZEXCEL_DRAWING_TYPE optional
    returning
      value(R_DRAWINGS) type ref to ZCL_EXCEL_DRAWINGS .
  methods GET_DRAWINGS_ITERATOR
    importing
      !IP_TYPE type ZEXCEL_DRAWING_TYPE
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_FREEZE_CELL
    exporting
      !EP_ROW type ZEXCEL_CELL_ROW
      !EP_COLUMN type ZEXCEL_CELL_COLUMN .
  methods GET_GUID
    returning
      value(EP_GUID) type UUID .
  methods GET_HIGHEST_COLUMN
    returning
      value(R_HIGHEST_COLUMN) type ZEXCEL_CELL_COLUMN
    raising
      ZCX_EXCEL .
  methods GET_HIGHEST_ROW
    returning
      value(R_HIGHEST_ROW) type INT4
    raising
      ZCX_EXCEL .
  methods GET_HYPERLINKS_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_HYPERLINKS_SIZE
    returning
      value(EP_SIZE) type I .
  methods GET_MERGE
    returning
      value(MERGE_RANGE) type STRING_TABLE
    raising
      ZCX_EXCEL .
  methods GET_PAGEBREAKS
    returning
      value(RO_PAGEBREAKS) type ref to ZCL_EXCEL_WORKSHEET_PAGEBREAKS
    raising
      ZCX_EXCEL .
  methods GET_RANGES_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_ROW
    importing
      !IP_ROW type INT4
    returning
      value(EO_ROW) type ref to ZCL_EXCEL_ROW .
  methods GET_ROWS
    returning
      value(EO_ROWS) type ref to ZCL_EXCEL_ROWS .
  methods GET_ROWS_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_ROW_OUTLINES
    returning
      value(RT_ROW_OUTLINES) type MTY_TS_OUTLINES_ROW .
  methods GET_STYLE_COND
    importing
      !IP_GUID type ZEXCEL_CELL_STYLE
    returning
      value(EO_STYLE_COND) type ref to ZCL_EXCEL_STYLE_COND .
  methods GET_TABCOLOR
    returning
      value(EV_TABCOLOR) type ZEXCEL_S_TABCOLOR .
  methods GET_TABLES_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_TABLES_SIZE
    returning
      value(EP_SIZE) type I .
  methods GET_TITLE
    importing
      !IP_ESCAPED type FLAG default ''
    returning
      value(EP_TITLE) type ZEXCEL_SHEET_TITLE .
  methods IS_CELL_MERGED
    importing
      !IP_COLUMN type SIMPLE
      !IP_ROW type ZEXCEL_CELL_ROW
    returning
      value(RP_IS_MERGED) type ABAP_BOOL
    raising
      ZCX_EXCEL .
  methods SET_CELL
    importing
      !IP_CELL type CSEQUENCE optional
      !IP_COLUMN type SIMPLE optional
      !IP_ROW type ZEXCEL_CELL_ROW optional
      !IP_COLUMN_TO type SIMPLE optional
      !IP_ROW_TO type ZEXCEL_CELL_ROW optional
      !IP_VALUE type SIMPLE optional
      !IP_FORMULA type ZEXCEL_CELL_FORMULA optional
      !IP_STYLE type ANY optional
      !IP_HYPERLINK type ref to ZCL_EXCEL_HYPERLINK optional
      !IP_DATA_TYPE type ZEXCEL_CELL_DATA_TYPE optional
      !IP_ABAP_TYPE type ABAP_TYPEKIND optional
      !IT_RTF type ZEXCEL_T_RTF optional
    raising
      ZCX_EXCEL .
  methods SET_CELL_FORMULA
    importing
      !IP_CELL type CSEQUENCE optional
      !IP_COLUMN type SIMPLE optional
      !IP_ROW type ZEXCEL_CELL_ROW optional
      !IP_FORMULA type ZEXCEL_CELL_FORMULA
    raising
      ZCX_EXCEL .
  methods SET_CELL_STYLE
    importing
      !IP_CELL type CSEQUENCE optional
      !IP_COLUMN type SIMPLE optional
      !IP_ROW type ZEXCEL_CELL_ROW optional
      !IP_STYLE type ANY
    raising
      ZCX_EXCEL .
  methods SET_COLUMN_WIDTH
    importing
      !IP_COLUMN type SIMPLE
      !IP_WIDTH_FIX type SIMPLE default 0
      !IP_WIDTH_AUTOSIZE type FLAG default 'X'
    raising
      ZCX_EXCEL .
  methods SET_DEFAULT_EXCEL_DATE_FORMAT
    importing
      !IP_DEFAULT_EXCEL_DATE_FORMAT type ZEXCEL_NUMBER_FORMAT
    raising
      ZCX_EXCEL .
  methods SET_MERGE
    importing
      !IP_CELL type CSEQUENCE optional
      !IP_COLUMN_START type SIMPLE default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MIN_COL
      !IP_COLUMN_END type SIMPLE default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MAX_COL
      !IP_ROW type ZEXCEL_CELL_ROW default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MIN_ROW
      !IP_ROW_TO type ZEXCEL_CELL_ROW default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MAX_ROW
    raising
      ZCX_EXCEL .
  methods SET_PRINT_GRIDLINES
    importing
      !I_PRINT_GRIDLINES type ZEXCEL_PRINT_GRIDLINES .
  methods SET_ROW_HEIGHT
    importing
      !IP_ROW type SIMPLE
      !IP_HEIGHT_FIX type SIMPLE
    raising
      ZCX_EXCEL .
  methods SET_ROW_OUTLINE
    importing
      !IV_ROW_FROM type I
      !IV_ROW_TO type I
      !IV_COLLAPSED type ABAP_BOOL
    raising
      ZCX_EXCEL .
  methods SET_SHOW_GRIDLINES
    importing
      !I_SHOW_GRIDLINES type ZEXCEL_SHOW_GRIDLINES .
  methods SET_SHOW_ROWCOLHEADERS
    importing
      !I_SHOW_ROWCOLHEADERS type ZEXCEL_SHOW_ROWCOLHEADER .
  methods SET_TABCOLOR
    importing
      !IV_TABCOLOR type ZEXCEL_S_TABCOLOR .
  methods SET_TABLE
    importing
      !IP_TABLE type STANDARD TABLE
      !IP_HDR_STYLE type ZEXCEL_CELL_STYLE optional
      !IP_BODY_STYLE type ZEXCEL_CELL_STYLE optional
      !IP_TABLE_TITLE type STRING
      !IP_TOP_LEFT_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA default 'B'
      !IP_TOP_LEFT_ROW type ZEXCEL_CELL_ROW default 3
      !IP_TRANSPOSE type XFELD optional
      !IP_NO_HEADER type XFELD optional
    raising
      ZCX_EXCEL .
  methods SET_TITLE
    importing
      !IP_TITLE type ZEXCEL_SHEET_TITLE
    raising
      ZCX_EXCEL .
  methods GET_TABLE
    importing
      !IV_SKIPPED_ROWS type INT4 default 0
      !IV_SKIPPED_COLS type INT4 default 0
    exporting
      !ET_TABLE type STANDARD TABLE
    raising
      ZCX_EXCEL .
protected section.
private section.

  types:
    BEGIN OF mty_s_font_metric,
      char       TYPE c LENGTH 1,
      char_width TYPE tdcwidths,
    END OF mty_s_font_metric .
  types:
    mty_th_font_metrics
           TYPE HASHED TABLE OF mty_s_font_metric
           WITH UNIQUE KEY char .
  types:
    BEGIN OF mty_s_font_cache,
      font_name       TYPE zexcel_style_font_name,
      font_height     TYPE tdfontsize,
      flag_bold       TYPE abap_bool,
      flag_italic     TYPE abap_bool,
      th_font_metrics TYPE mty_th_font_metrics,
    END OF mty_s_font_cache .
  types:
    mty_th_font_cache
           TYPE HASHED TABLE OF mty_s_font_cache
           WITH UNIQUE KEY font_name font_height flag_bold flag_italic .
  types:
*  types:
*    mty_ts_row_dimension TYPE SORTED TABLE OF zexcel_s_worksheet_rowdimensio WITH UNIQUE KEY row .
    BEGIN OF mty_merge,
      row_from TYPE i,
      row_to   TYPE i,
      col_from TYPE i,
      col_to   TYPE i,
    END OF mty_merge .
  types:
    mty_ts_merge TYPE SORTED TABLE OF mty_merge WITH UNIQUE KEY table_line .

*"* private components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
  data ACTIVE_CELL type ZEXCEL_S_CELL_DATA .
  data CHARTS type ref to ZCL_EXCEL_DRAWINGS .
  data COLUMNS type ref to ZCL_EXCEL_COLUMNS .
  data ROW_DEFAULT type ref to ZCL_EXCEL_ROW .
  data COLUMN_DEFAULT type ref to ZCL_EXCEL_COLUMN .
  data STYLES_COND type ref to ZCL_EXCEL_STYLES_COND .
  data DATA_VALIDATIONS type ref to ZCL_EXCEL_DATA_VALIDATIONS .
  data DEFAULT_EXCEL_DATE_FORMAT type ZEXCEL_NUMBER_FORMAT .
  data DEFAULT_EXCEL_TIME_FORMAT type ZEXCEL_NUMBER_FORMAT .
  data DRAWINGS type ref to ZCL_EXCEL_DRAWINGS .
  data FREEZE_PANE_CELL_COLUMN type ZEXCEL_CELL_COLUMN .
  data FREEZE_PANE_CELL_ROW type ZEXCEL_CELL_ROW .
  data GUID type UUID .
  data HYPERLINKS type ref to CL_OBJECT_COLLECTION .
  data LOWER_CELL type ZEXCEL_S_CELL_DATA .
  data MO_PAGEBREAKS type ref to ZCL_EXCEL_WORKSHEET_PAGEBREAKS .
  class-data MTH_FONT_CACHE type MTY_TH_FONT_CACHE .
  data MT_MERGED_CELLS type MTY_TS_MERGE .
  data MT_ROW_OUTLINES type MTY_TS_OUTLINES_ROW .
  data PRINT_TITLE_COL_FROM type ZEXCEL_CELL_COLUMN_ALPHA .
  data PRINT_TITLE_COL_TO type ZEXCEL_CELL_COLUMN_ALPHA .
  data PRINT_TITLE_ROW_FROM type ZEXCEL_CELL_ROW .
  data PRINT_TITLE_ROW_TO type ZEXCEL_CELL_ROW .
  data RANGES type ref to ZCL_EXCEL_RANGES .
  data ROWS type ref to ZCL_EXCEL_ROWS .
  data TABLES type ref to CL_OBJECT_COLLECTION .
  data TITLE type ZEXCEL_SHEET_TITLE value 'Worksheet'. "#EC NOTEXT
  data UPPER_CELL type ZEXCEL_S_CELL_DATA .

  methods CALCULATE_CELL_WIDTH
    importing
      !IP_COLUMN type SIMPLE
      !IP_ROW type ZEXCEL_CELL_ROW
    returning
      value(EP_WIDTH) type FLOAT
    raising
      ZCX_EXCEL .
  methods GENERATE_TITLE
    returning
      value(EP_TITLE) type ZEXCEL_SHEET_TITLE .
  methods GET_VALUE_TYPE
    importing
      !IP_VALUE type SIMPLE
    exporting
      !EP_VALUE type SIMPLE
      !EP_VALUE_TYPE type ABAP_TYPEKIND .
  methods PRINT_TITLE_SET_RANGE .
  methods UPDATE_DIMENSION_RANGE
    raising
      ZCX_EXCEL .
  methods DETERMINE_RANGE
    importing
      !IP_CELL type CSEQUENCE optional
      !IP_COLUMN type SIMPLE optional
      !IP_ROW type ZEXCEL_CELL_ROW optional
      !IP_COLUMN_TO type SIMPLE optional
      !IP_ROW_TO type ZEXCEL_CELL_ROW optional
    exporting
      !EV_COLUMN_I_START type ZEXCEL_CELL_COLUMN
      !EV_COLUMN_A_START type ZEXCEL_CELL_COLUMN_ALPHA
      !EV_ROW_START type ZEXCEL_CELL_ROW
      !EV_COLUMN_I_END type ZEXCEL_CELL_COLUMN
      !EV_COLUMN_A_END type ZEXCEL_CELL_COLUMN_ALPHA
      !EV_ROW_END type ZEXCEL_CELL_ROW .
  methods CHECK_RTF
    importing
      !IP_VALUE type SIMPLE
      !IT_RTF type ZEXCEL_T_RTF .
ENDCLASS.



CLASS ZCL_EXCEL_WORKSHEET IMPLEMENTATION.


method ADD_DRAWING.
  CASE ip_drawing->get_type( ).
    WHEN zcl_excel_drawing=>type_image.
      drawings->include( ip_drawing ).
    WHEN zcl_excel_drawing=>type_chart.
      charts->include( ip_drawing ).
  ENDCASE.
  endmethod.


METHOD add_new_column.
  DATA: lv_column_alpha TYPE zexcel_cell_column_alpha.

  lv_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column ).

  CREATE OBJECT eo_column
    EXPORTING
      ip_index     = lv_column_alpha
      ip_excel     = me->excel
      ip_worksheet = me.
  columns->add( eo_column ).
ENDMETHOD.


method ADD_NEW_DATA_VALIDATION.

  CREATE OBJECT eo_data_validation.
  data_validations->add( eo_data_validation ).
  endmethod.


METHOD add_new_range.
* Create default blank range
  CREATE OBJECT eo_range.
  ranges->add( eo_range ).
ENDMETHOD.


METHOD add_new_row.
  CREATE OBJECT eo_row
    EXPORTING
      ip_index     = ip_row.
  rows->add( eo_row ).
ENDMETHOD.


METHOD add_new_style_cond.
  CREATE OBJECT eo_style_cond.
  styles_cond->add( eo_style_cond ).
ENDMETHOD.


method BIND_ALV.
  data: lo_converter type ref to zcl_excel_converter.

  create object lo_converter.

  try.
      lo_converter->convert(
        exporting
          io_alv         = io_alv
          it_table       = it_table
          i_row_int      = i_top
          i_column_int   = i_left
          i_table        = i_table
          i_style_table  = table_style
          io_worksheet   = me
        changing
          co_excel       = excel ).
    catch zcx_excel .
  endtry.

  endmethod.


method BIND_ALV_OLE2.
*--------------------------------------------------------------------*
* Method description:
*   Method use to export a CL_GUI_ALV_GRID object to xlsx/xls file
*   with list header and  characteristics of ALV field catalog such as:
*     + Total, group's subtotal
*     + Quantity fields, amount fields (dependent fields)
*     + No_out, no_zero, ...
* Technique use in method:
*   SAP Desktop Office Integration (DOI)
*--------------------------------------------------------------------*

* Data for session 0: DOI constructor
* ------------------------------------------

  data: lo_control  type ref to I_OI_CONTAINER_CONTROL.
  data: lo_proxy    type ref to I_OI_DOCUMENT_PROXY.
  data: lo_spreadsheet type ref to I_OI_SPREADSHEET.
  data: lo_error    type ref to I_OI_ERROR.
  data: lc_retcode  type SOI_RET_STRING.
  data: li_has      type i. "Proxy has spreadsheet interface?
  data: l_is_closed type i.

* Data for session 1: Get LVC data from ALV object
* ------------------------------------------

  data: l_has_activex,
        l_doctype_excel_sheet(11) type c.

* LVC
  data: lt_fieldcat_lvc       type LVC_T_FCAT.
  data: wa_fieldcat_lvc       type lvc_s_fcat.
  data: lt_sort_lvc           type LVC_T_SORT.
  data: lt_filter_idx_lvc     type LVC_T_FIDX.
  data: lt_GROUPLEVELS_LVC    type LVC_T_GRPL.

* KKBLO
  DATA: LT_FIELDCAT_KKBLO     Type  KKBLO_T_FIELDCAT.
  DATA: LT_SORT_KKBLO         Type  KKBLO_T_SORTINFO.
  DATA: LT_GROUPLEVELS_KKBLO  Type  KKBLO_T_GROUPLEVELS.
  DATA: LT_FILTER_IDX_KKBLO   Type  KKBLO_T_SFINFO.
  data: wa_listheader         like line of it_listheader.

* Subtotal
  data: lt_collect00          type ref to data.
  data: lt_collect01          type ref to data.
  data: lt_collect02          type ref to data.
  data: lt_collect03          type ref to data.
  data: lt_collect04          type ref to data.
  data: lt_collect05          type ref to data.
  data: lt_collect06          type ref to data.
  data: lt_collect07          type ref to data.
  data: lt_collect08          type ref to data.
  data: lt_collect09          type ref to data.

* data table name
  data: l_tabname             type  kkblo_tabname.

* local object
  data: lo_grid               type ref to lcl_gui_alv_grid.

* data table get from ALV
  data: lt_alv                  type ref to data.

* total / subtotal data
  field-symbols: <f_collect00>  type standard table.
  field-symbols: <f_collect01>  type standard table.
  field-symbols: <f_collect02>  type standard table.
  field-symbols: <f_collect03>  type standard table.
  field-symbols: <f_collect04>  type standard table.
  field-symbols: <f_collect05>  type standard table.
  field-symbols: <f_collect06>  type standard table.
  field-symbols: <f_collect07>  type standard table.
  field-symbols: <f_collect08>  type standard table.
  field-symbols: <f_collect09>  type standard table.

* table before append subtotal lines
  field-symbols: <f_alv_tab>    type standard table.

* data for session 2: sort, filter and calculate total/subtotal
* ------------------------------------------

* table to save index of subotal / total line in excel tanle
* this ideal to control index of subtotal / total line later
* for ex, when get subtotal / total line to format
  types: begin of st_subtot_indexs,
          index type i,
        end of st_subtot_indexs.
  data: lt_subtot_indexs type table of st_subtot_indexs.
  data: wa_subtot_indexs like line of lt_subtot_indexs.

* data table after append subtotal
  data: lt_excel                type ref to data.

  data: l_tabix                 type i.
  data: l_save_index            type i.

* dyn subtotal table name
  data: l_collect               type string.

* subtotal range, to format subtotal (and total)
  data: subranges               type soi_range_list.
  data: subrangeitem            type soi_range_item.
  data: l_sub_index             type i.


* table after append subtotal lines
  field-symbols: <f_excel_tab>  type standard table.
  field-symbols: <f_excel_line> type any.

* dyn subtotal tables
  field-symbols: <f_collect_tab>      type standard table.
  field-symbols: <f_collect_line>     type any.

  field-symbols: <f_filter_idx_line>  like line of LT_FILTER_IDX_KKBLO.
  field-symbols: <f_fieldcat_line>    like line of LT_FIELDCAT_KKBLO.
  field-symbols: <f_grouplevels_line> like line of LT_GROUPLEVELS_KKBLO.
  field-symbols: <f_line>             type any.

* Data for session 3: map data to semantic table
* ------------------------------------------

  types: begin of st_column_index,
          fieldname type kkblo_fieldname,
          tabname   type kkblo_tabname,
          col like sy-index,
         end of st_column_index.

* columns index
  data: lt_column_index   type table of st_column_index.
  data: wa_column_index   like line of lt_column_index.

* table of dependent field ( currency and quantity unit field)
  data: lt_fieldcat_depf  type kkblo_t_fieldcat.
  data: wa_fieldcat_depf  type kkblo_fieldcat.

* XXL interface:
* -XXL: contain exporting columns characteristic
  data: lt_sema type table of gxxlt_s initial size 0.
  data: wa_sema like line of lt_sema.

* -XXL interface: header
  data: lt_hkey type table of gxxlt_h initial size 0.
  data: wa_hkey like line of lt_hkey.

* -XXL interface: header keys
  data: lt_vkey type table of gxxlt_v initial size 0.
  data: wa_vkey like line of lt_vkey.

* Number of H Keys: number of key columns
  data: l_n_hrz_keys      type  i.
* Number of data columns in the list object: non-key columns no
  data: l_n_att_cols      type  i.
* Number of V Keys: number of header row
  data: l_n_vrt_keys      type  i.

* curency to format amount
  data: lt_tcurx          type table of tcurx.
  data: wa_tcurx          like line of lt_tcurx.
  data: l_def             type flag. " currency / quantity flag
  data: wa_t006           type t006. " decimal place of unit

  data: l_num             type i. " table columns number
  data: l_typ             type c. " table type
  data: wa                type ref to data.
  data: l_int             type i.
  data: l_counter         type i.

  field-symbols: <f_excel_column>     type any.
  field-symbols: <f_fcat_column>      type any.

* Data for session 4: write to excel
* ------------------------------------------

  data: sema_type         type  c.

  data l_error           type ref to c_oi_proxy_error.
  data count              type i.
  data datac              type i.
  data datareal           type i. " exporting column number
  data vkeycount          type i.
  data all type i.
  data mit type i         value 1.  " index of recent row?
  data li_col_pos type i  value 1.  " column position
  data li_col_num type i.           " table columns number
  field-symbols: <line>   type any.
  field-symbols: <item>   type any.

  data td                 type sydes_desc.

  data: typ.
  data: ranges             type soi_range_list.
  data: rangeitem          type soi_range_item.
  data: contents           type soi_generic_table.
  data: contentsitem       type soi_generic_item.
  data: semaitem           type gxxlt_s.
  data: hkeyitem           type gxxlt_h.
  data: vkeyitem           type gxxlt_v.
  data: li_commentary_rows type i.  "row number of title lines + 1
  data: lo_error_w         type ref to  i_oi_error.
  data: l_retcode          type soi_ret_string.
  data: no_flush           type c value 'X'.
  data: li_head_top        type i. "header rows position

* Data for session 5: Save and clode document
* ------------------------------------------

  data: li_document_size   type i.
  data: ls_path            type RLGRAP-FILENAME.

* MACRO: Close_document
*-------------------------------------------

  DEFINE close_document.
    clear: l_is_closed.
    IF lo_proxy is not initial.

* check proxy detroyed adi

      call method lo_proxy->is_destroyed
        IMPORTING
          ret_value = l_is_closed.

* if dun detroyed yet: close -> release proxy

      IF l_is_closed is initial.
        call method lo_proxy->close_document
*        EXPORTING
*          do_save = do_save
          IMPORTING
            error       = lo_error
            retcode     = lc_retcode.
      ENDIF.

      call method lo_proxy->release_document
        IMPORTING
          error   = lo_error
          retcode = lC_retcode.

    else.
      lc_retcode = c_oi_errors=>ret_document_not_open.
    ENDIF.

* Detroy control container

    IF lo_control is not initial.
      CALL METHOD lo_control->destroy_control.
    ENDIF.

    clear:
      lo_spreadsheet,
      lo_proxy,
      lo_control.

* free local

    clear: l_is_closed.

  END-OF-DEFINITION.

* Macro to catch DOI error
*-------------------------------------------

  DEFINE error_doi.
    if lc_retcode ne c_oi_errors=>ret_ok.
      close_document.
      call method lo_error->raise_message
        EXPORTING
          type = 'E'.
      clear: lo_error.
    endif.
  END-OF-DEFINITION.

*--------------------------------------------------------------------*
* SESSION 0: DOI CONSTRUCTOR
*--------------------------------------------------------------------*

* check active windown

  call function 'GUI_HAS_ACTIVEX'
    IMPORTING
      return = l_has_activex.

  if l_has_activex is initial.
    raise MISS_GUIDE.
  endif.

*   Get Container Object of Screen

  call method c_oi_container_control_creator=>get_container_control
    IMPORTING
      control = lo_control
      retcode = lC_retcode.

  error_doi.

* Initialize Container control

  CALL METHOD lo_control->init_control
    EXPORTING
      parent                   = CL_GUI_CONTAINER=>DEFAULT_SCREEN
      r3_application_name      = ''
      inplace_enabled          = 'X'
      no_flush                 = 'X'
      register_on_close_event  = 'X'
      register_on_custom_event = 'X'
    IMPORTING
      error                    = lO_ERROR
      retcode                  = lc_retcode.

  error_doi.

* Get Proxy Document:
* check exist of document proxy, if exist -> close first

  if not lo_proxy is initial.
    close_document.
  endif.

  IF i_xls is not initial.
* xls format, doctype = soi_doctype_excel97_sheet
    l_doctype_excel_sheet = 'Excel.Sheet.8'.
  else.
* xlsx format, doctype = soi_doctype_excel_sheet
    l_doctype_excel_sheet = 'Excel.Sheet'.
  ENDIF.

  CALL METHOD lo_control->get_document_proxy
    EXPORTING
      document_type      = l_doctype_excel_sheet
      register_container = 'X'
    IMPORTING
      document_proxy     = lo_proxy
      error              = lO_ERROR
      retcode            = lc_retcode.

  error_doi.

  IF I_DOCUMENT_URL is initial.

* create new excel document

    call method lo_proxy->create_document
      EXPORTING
        create_view_data = 'X'
        open_inplace     = 'X'
        no_flush         = 'X'
      IMPORTING
        ERROR            = lO_ERROR
        retcode          = lc_retcode.

    error_doi.

  else.

* Read excel template for i_DOCUMENT_URL
* this excel template can be store in local or server

    CALL METHOD lo_proxy->open_document
      EXPORTING
        document_url = i_document_url
        open_inplace = 'X'
        no_flush     = 'X'
      IMPORTING
        error        = lo_error
        retcode      = lc_retcode.

    error_doi.

  endif.

* Check Spreadsheet Interface of Document Proxy

  CALL METHOD lo_proxy->has_spreadsheet_interface
    IMPORTING
      is_available = li_has
      error        = lO_ERROR
      retcode      = lc_retcode.

  error_doi.

* create Spreadsheet object

  CHECK li_has IS NOT INITIAL.

  CALL METHOD lo_proxy->get_spreadsheet_interface
    IMPORTING
      sheet_interface = lo_spreadsheet
      error           = lO_ERROR
      retcode         = lc_retcode.

  error_doi.

*--------------------------------------------------------------------*
* SESSION 1: GET LVC DATA FROM ALV OBJECT
*--------------------------------------------------------------------*

* data table

  create object lo_grid
    EXPORTING
      i_parent = CL_GUI_CONTAINER=>SCREEN0.

  call method lo_grid->get_alv_attributes
    EXPORTING
      io_grid  = io_alv
    IMPORTING
      Et_table = lt_alv.

  assign lt_alv->* to <f_alv_tab>.

* fieldcat

  CALL METHOD iO_alv->GET_FRONTEND_FIELDCATALOG
    IMPORTING
      ET_FIELDCATALOG = lt_fieldcat_LVC.

* table name

  loop at lt_fieldcat_LVC into wa_fieldcat_lvc
  where not tabname is initial.
    l_tabname = wa_fieldcat_lvc-tabname.
    exit.
  endloop.

  if sy-subrc ne 0.
    l_tabname = '1'.
  endif.
  clear: wa_fieldcat_lvc.

* sort table

  CALL METHOD IO_ALV->GET_SORT_CRITERIA
    IMPORTING
      ET_SORT = lt_sort_lvc.


* filter index

  CALL METHOD IO_ALV->GET_FILTERED_ENTRIES
    IMPORTING
      ET_FILTERED_ENTRIES = lt_filter_idx_lvc.

* group level + subtotal

  CALL METHOD IO_ALV->GET_SUBTOTALS
    IMPORTING
      EP_COLLECT00   = lt_collect00
      EP_COLLECT01   = lt_collect01
      EP_COLLECT02   = lt_collect02
      EP_COLLECT03   = lt_collect03
      EP_COLLECT04   = lt_collect04
      EP_COLLECT05   = lt_collect05
      EP_COLLECT06   = lt_collect06
      EP_COLLECT07   = lt_collect07
      EP_COLLECT08   = lt_collect08
      EP_COLLECT09   = lt_collect09
      ET_GROUPLEVELS = lt_GROUPLEVELS_LVC.

  assign lt_collect00->* to <f_collect00>.
  assign lt_collect01->* to <f_collect01>.
  assign lt_collect02->* to <f_collect02>.
  assign lt_collect03->* to <f_collect03>.
  assign lt_collect04->* to <f_collect04>.
  assign lt_collect05->* to <f_collect05>.
  assign lt_collect06->* to <f_collect06>.
  assign lt_collect07->* to <f_collect07>.
  assign lt_collect08->* to <f_collect08>.
  assign lt_collect09->* to <f_collect09>.

* transfer to KKBLO struct

  CALL FUNCTION 'LVC_TRANSFER_TO_KKBLO'
    EXPORTING
      IT_FIELDCAT_LVC           = lt_fieldcat_lvc
      IT_SORT_LVC               = lt_sort_lvc
      IT_FILTER_INDEX_LVC       = lt_filter_idx_lvc
      IT_GROUPLEVELS_LVC        = lt_grouplevels_lvc
    IMPORTING
      ET_FIELDCAT_KKBLO         = lt_fieldcat_kkblo
      ET_SORT_KKBLO             = lt_sort_kkblo
      ET_FILTERED_ENTRIES_KKBLO = lt_filter_idx_kkblo
      ET_GROUPLEVELS_KKBLO      = lt_grouplevels_kkblo
    TABLES
      IT_DATA                   = <f_alv_tab>
    EXCEPTIONS
      IT_DATA_MISSING           = 1
      IT_FIELDCAT_LVC_MISSING   = 2
      OTHERS                    = 3.
  IF SY-SUBRC <> 0.
    raise ex_transfer_KKBLO_ERROR.
  ENDIF.

  clear:
    wa_fieldcat_lvc,
    lt_fieldcat_lvc,
    lt_sort_lvc,
    lt_filter_idx_lvc,
    lt_GROUPLEVELS_LVC.

  clear:
    lo_grid.


*--------------------------------------------------------------------*
* SESSION 2: SORT, FILTER AND CALCULATE TOTAL / SUBTOTAL
*--------------------------------------------------------------------*

* append subtotal & total line

  create data lt_excel like <f_ALV_TAB>.
  assign lt_excel->* to <f_excel_tab>.

  loop at <f_alv_tab> assigning <f_line>.
    l_save_index = sy-tabix.

* filter base on filter index table

    read table LT_FILTER_IDX_KKBLO assigning <f_filter_idx_line>
    with key index = l_save_index
    binary search.
    if sy-subrc ne 0.
      append <f_line> to <f_excel_tab>.
    endif.

* append subtotal lines

    read table LT_GROUPLEVELS_KKBLO assigning <f_grouplevels_line>
    with key index_to = l_save_index
    binary search.
    if sy-subrc = 0.
      l_tabix = sy-tabix.
      do.
        if <f_grouplevels_line>-subtot eq 'X' and
           <f_grouplevels_line>-hide_level is initial and
           <f_grouplevels_line>-cindex_from ne 0.

* dynamic append subtotal line to excel table base on grouplevel table
* ex <f_GROUPLEVELS_line>-level = 1
* then <f_collect_tab> = '<F_COLLECT01>'

          l_collect = <f_grouplevels_line>-level.
          condense l_collect.
          concatenate '<F_COLLECT0'
                      l_collect '>'
*                      '->*'
                      into l_collect.

          assign (l_collect) to <f_collect_tab>.

* incase there're more than 1 total line of group, at the same level
* for example: subtotal of multi currency

          LOOP AT <f_collect_tab> assigning <f_collect_line>.
            IF  sy-tabix between <f_grouplevels_line>-cindex_from
                            and  <f_grouplevels_line>-cindex_to.


              append <f_collect_line> to <f_excel_tab>.

* save subtotal lines index

              wa_subtot_indexs-index = sy-tabix.
              append wa_subtot_indexs to lt_subtot_indexs.

* append sub total ranges table for format later

              add 1 to l_sub_index.
              subrangeitem-name     =  l_sub_index.
              condense subrangeitem-name.
              concatenate 'SUBTOT'
                          subrangeitem-name
                          into subrangeitem-name.

              subrangeitem-rows     = wa_subtot_indexs-index.
              subrangeitem-columns  = 1.            " start col
              append subrangeitem to subranges.
              clear: subrangeitem.

            ENDIF.
          ENDLOOP.
          unassign: <f_collect_tab>.
          unassign: <f_collect_line>.
          clear: l_collect.
        endif.

* check next subtotal level of group

        unassign: <f_grouplevels_line>.
        add 1 to l_tabix.

        read table LT_GROUPLEVELS_KKBLO assigning <f_grouplevels_line>
        index l_tabix.
        if sy-subrc ne 0
        or <f_grouplevels_line>-index_to ne l_save_index.
          exit.
        endif.

        unassign:
          <f_collect_tab>,
          <f_collect_line>.

      enddo.
    endif.

    clear:
      l_tabix,
      l_save_index.

    unassign:
      <f_filter_idx_line>,
      <f_grouplevels_line>.

  endloop.

* free local data

  unassign:
    <f_line>,
    <f_collect_tab>,
    <f_collect_line>,
    <f_fieldcat_line>.

* append grand total line

  IF <f_collect00> is assigned.
    assign <f_collect00> to <f_collect_tab>.
    if <f_collect_tab> is not initial.
      LOOP AT <f_collect_tab> assigning <f_collect_line>.

        append <f_collect_line> to <f_excel_tab>.

* save total line index

        wa_subtot_indexs-index = sy-tabix.
        append wa_subtot_indexs to lt_subtot_indexs.

* append grand total range (to format)

        add 1 to l_sub_index.
        subrangeitem-name     =  l_sub_index.
        condense subrangeitem-name.
        concatenate 'TOTAL'
                    subrangeitem-name
                    into subrangeitem-name.

        subrangeitem-rows     = wa_subtot_indexs-index.
        subrangeitem-columns  = 1.            " start col
        append subrangeitem to subranges.
      ENDLOOP.
    endif.
  ENDIF.

  clear:
    subrangeitem,
    LT_SORT_KKBLO,
    <f_collect00>,
    <f_collect01>,
    <f_collect02>,
    <f_collect03>,
    <f_collect04>,
    <f_collect05>,
    <f_collect06>,
    <f_collect07>,
    <f_collect08>,
    <f_collect09>.

  unassign:
    <f_collect00>,
    <f_collect01>,
    <f_collect02>,
    <f_collect03>,
    <f_collect04>,
    <f_collect05>,
    <f_collect06>,
    <f_collect07>,
    <f_collect08>,
    <f_collect09>,
    <f_collect_tab>,
    <f_collect_line>.

*--------------------------------------------------------------------*
* SESSION 3: MAP DATA TO SEMANTIC TABLE
*--------------------------------------------------------------------*

* get dependent field field: currency and quantity

  create data wa like line of <f_excel_tab>.
  assign wa->* to <f_excel_line>.

  describe field <f_excel_line> type l_typ components l_num.

  do l_num times.
    l_save_index = sy-index.
    assign component l_save_index of structure <f_excel_line>
    to <f_excel_column>.
    if sy-subrc ne 0.
      message e059(0k) with 'FATAL ERROR' raising fatal_error.
    endif.

    loop at LT_FIELDCAT_KKBLO assigning <f_fieldcat_line>
    where tabname = l_tabname.
      assign component <f_fieldcat_line>-fieldname
      of structure <f_excel_line> to <f_fcat_column>.

      describe distance between <f_excel_column> and <f_fcat_column>
      into l_int in byte mode.

* append column index
* this columns index is of table, not fieldcat

      if l_int = 0.
        wa_column_index-fieldname = <f_fieldcat_line>-fieldname.
        wa_column_index-tabname   = <f_fieldcat_line>-tabname.
        wa_column_index-col       = l_save_index.
        append wa_column_index to lt_column_index.
      endif.

* append dependent fields (currency and quantity unit)

      if <f_fieldcat_line>-cfieldname is not initial.
        clear wa_fieldcat_depf.
        wa_fieldcat_depf-fieldname = <f_fieldcat_line>-cfieldname.
        wa_fieldcat_depf-tabname   = <f_fieldcat_line>-ctabname.
        collect wa_fieldcat_depf into lt_fieldcat_depf.
      endif.

      if <f_fieldcat_line>-qfieldname is not initial.
        clear wa_fieldcat_depf.
        wa_fieldcat_depf-fieldname = <f_fieldcat_line>-qfieldname.
        wa_fieldcat_depf-tabname   = <f_fieldcat_line>-qtabname.
        collect wa_fieldcat_depf into lt_fieldcat_depf.
      endif.

* rewrite field data type

      if <f_fieldcat_line>-inttype = 'X'
      and <f_fieldcat_line>-datatype(3) = 'INT'.
        <f_fieldcat_line>-inttype = 'I'.
      endif.

    endloop.

    clear: l_save_index.
    unassign: <f_fieldcat_line>.

  enddo.

* build semantic tables

  l_n_hrz_keys = 1.

*   Get keyfigures

  loop at LT_FIELDCAT_KKBLO assigning <f_fieldcat_line>
  where tabname = l_tabname
  and tech ne 'X'
  and no_out ne 'X'.

    clear wa_sema.
    clear wa_hkey.

*   Units belong to keyfigures -> display as str

    read table lt_fieldcat_depf into wa_fieldcat_depf with key
    fieldname = <f_fieldcat_line>-fieldname
    tabname   = <f_fieldcat_line>-tabname.

    if sy-subrc = 0.
      wa_sema-col_typ = 'STR'.
      wa_sema-col_ops = 'DFT'.

*   Keyfigures

    else.
      case <f_fieldcat_line>-datatype.
        when 'QUAN'.
          wa_sema-col_typ = 'N03'.

          if <f_fieldcat_line>-no_sum ne 'X'.
            wa_sema-col_ops = 'ADD'.
          else.
            wa_sema-col_ops = 'NOP'. " no dependent field
          endif.

        when 'DATS'.
          wa_sema-col_typ = 'DAT'.
          wa_sema-col_ops = 'NOP'.

        when 'CHAR' OR 'UNIT' OR 'CUKY'. " Added fieldformats UNIT and CUKY - dd. 26-10-2012 Wouter Heuvelmans
          wa_sema-col_typ = 'STR'.
          wa_sema-col_ops = 'DFT'.   " dependent field

*   incase numeric, ex '00120' -> display as '12'

        when 'NUMC'.
          wa_sema-col_typ = 'STR'.
          wa_sema-col_ops = 'DFT'.

        when others.
          wa_sema-col_typ = 'NUM'.

          if <f_fieldcat_line>-no_sum ne 'X'.
            wa_sema-col_ops = 'ADD'.
          else.
            wa_sema-col_ops = 'NOP'.
          endif.
      endcase.
    endif.

    l_counter = l_counter + 1.
    l_n_att_cols = l_n_att_cols + 1.

    wa_sema-col_no = l_counter.

    read table lt_column_index into wa_column_index with key
    fieldname = <f_fieldcat_line>-fieldname
    tabname   = <f_fieldcat_line>-tabname.

    if sy-subrc = 0.
      wa_sema-col_src = wa_column_index-col.
    else.
      raise fatal_error.
    endif.

* columns index of ref currency field in table

    if not <f_fieldcat_line>-cfieldname is initial.
      read table lt_column_index into wa_column_index with key
      fieldname = <f_fieldcat_line>-cfieldname
      tabname   = <f_fieldcat_line>-ctabname.

      if sy-subrc = 0.
        wa_sema-col_cur = wa_column_index-col.
      endif.

* quantities fields
* treat as currency when display on excel

    elseif not <f_fieldcat_line>-qfieldname is initial.
      read table lt_column_index into wa_column_index with key
      fieldname = <f_fieldcat_line>-qfieldname
      tabname   = <f_fieldcat_line>-qtabname.
      if sy-subrc = 0.
        wa_sema-col_cur = wa_column_index-col.
      endif.

    endif.

*   Treat of fixed currency in the fieldcatalog for column

    data: l_num_help(2) type n.

    if not <f_fieldcat_line>-currency is initial.

      select * from tcurx into table lt_tcurx.
      sort lt_tcurx.
      read table lt_tcurx into wa_tcurx
                 with key currkey = <f_fieldcat_line>-currency.
      if sy-subrc = 0.
        l_num_help = wa_tcurx-currdec.
        concatenate 'N' l_num_help into wa_sema-col_typ.
        wa_sema-col_cur = sy-tabix * ( -1 ).
      endif.

    endif.

    wa_hkey-col_no    = l_n_att_cols.
    wa_hkey-row_no    = l_n_hrz_keys.
    wa_hkey-col_name  = <f_fieldcat_line>-reptext.
    append wa_hkey to lt_hkey.
    append wa_sema to lt_sema.

  endloop.

* free local data

  clear:
    lt_column_index,
    wa_column_index,
    lt_fieldcat_depf,
    wa_fieldcat_depf,
    lt_tcurx,
    wa_tcurx,
    l_num,
    l_typ,
    wa,
    l_int,
    l_counter.

  unassign:
    <f_fieldcat_line>,
    <f_excel_line>,
    <f_excel_column>,
    <f_fcat_column>.

*--------------------------------------------------------------------*
* SESSION 4: WRITE TO EXCEL
*--------------------------------------------------------------------*

  clear: wa_tcurx.
  refresh: lt_tcurx.

*   if spreadsheet dun have proxy yet

  if li_has is initial.
    l_retcode = c_oi_errors=>ret_interface_not_supported.
    call method c_oi_errors=>create_error_for_retcode
      EXPORTING
        retcode  = l_retcode
        no_flush = no_flush
      IMPORTING
        error    = lo_error_w.
    exit.
  endif.

  create object l_error
    EXPORTING
      object_name = 'OLE_DOCUMENT_PROXY'
      method_name = 'get_ranges_names'.

  call method c_oi_errors=>add_error
    EXPORTING
      error = l_error.


  describe table lt_sema lines datareal.
  describe table <f_excel_tab> lines datac.
  describe table lt_vkey lines vkeycount.

  if datac = 0.
    raise inv_data_range.
  endif.


  if vkeycount ne l_n_vrt_keys.
    raise dim_mismatch_vkey.
  endif.

  all = l_n_vrt_keys + l_n_att_cols.

  if datareal ne all.
    raise dim_mismatch_sema.
  endif.

  data: decimal type c.

* get decimal separator format ('.', ',', ...) in Office config

  call method lo_proxy->get_application_property
    EXPORTING
      property_name    = 'INTERNATIONAL'
      subproperty_name = 'DECIMAL_SEPARATOR'
    CHANGING
      retvalue         = decimal.

  data: wa_usr type usr01.
  select * from usr01 into wa_usr where bname = sy-uname.
  endselect.

  data: comma_elim(4) type c.
  field-symbols <g> type any.
  data search_item(4) value '   #'.

  concatenate ',' decimal '.' decimal into comma_elim.

  data help type i. " table (with subtotal) line number

  help = datac.

  data: rowmax type i value 1.    " header row number
  data: columnmax type i value 0. " header columns number

  loop at lt_hkey into hkeyitem.
    if hkeyitem-col_no > columnmax.
      columnmax = hkeyitem-col_no.
    endif.

    if hkeyitem-row_no > rowmax.
      rowmax = hkeyitem-row_no.
    endif.
  endloop.

  data: hkeycolumns type i. " header columns no

  hkeycolumns = columnmax.

  if hkeycolumns <   l_n_att_cols.
    hkeycolumns = l_n_att_cols.
  endif.

  columnmax = 0.

  loop at lt_vkey into vkeyitem.
    if vkeyitem-col_no > columnmax.
      columnmax = vkeyitem-col_no.
    endif.
  endloop.

  data overflow type i value 1.
  data testname(10) type c.
  data temp2 type i.                " 1st item row position in excel
  data realmit type i value 1.
  data realoverflow type i value 1. " row index in content

  call method lo_spreadsheet->screen_update
    EXPORTING
      updating = ''.

  call method lo_spreadsheet->load_lib.

  data: str(40) type c. " range names of columns range (w/o col header)
  data: rows type i.    " row postion of 1st item line in ecxel

* calculate row position of data table

  describe table iT_LISTHEADER lines li_commentary_rows.

* if grid had title, add 1 empy line between title and table

  if li_commentary_rows ne 0.
    add 1 to li_commentary_rows.
  endif.

* add top position of block data

  li_commentary_rows = li_commentary_rows + i_top - 1.

* write header (commentary rows)

  data: li_commentary_row_index type i value 1.
  data: li_content_index type i value 1.
  data: ls_index(10) type c.
  data  ls_commentary_range(40) type c value 'TITLE'.
  data: li_font_bold    type i.
  data: li_font_italic  type i.
  data: li_font_size    type i.

  loop at iT_LISTHEADER into wa_listheader.
    li_commentary_row_index = i_top + li_content_index - 1.
    ls_index = li_content_index.
    condense ls_index.
    concatenate ls_commentary_range(5) ls_index
                into ls_commentary_range.
    condense ls_commentary_range.

* insert title range

    call method lo_spreadsheet->insert_range_dim
      EXPORTING
        name     = ls_commentary_range
        top      = li_commentary_row_index
        left     = i_left
        rows     = 1
        columns  = 1
        no_flush = no_flush.

* format range

    case wa_listheader-typ.
      when 'H'. "title
        li_font_size    = 16.
        li_font_bold    = 1.
        li_font_italic  = -1.
      when 'S'. "subtile
        li_font_size = -1.
        li_font_bold    = 1.
        li_font_italic  = -1.
      when others. "'A' comment
        li_font_size = -1.
        li_font_bold    = -1.
        li_font_italic  = 1.
    endcase.

    call method lo_spreadsheet->set_font
      EXPORTING
        rangename = ls_commentary_range
        family    = ''
        size      = li_font_size
        bold      = li_font_bold
        italic    = li_font_italic
        align     = 0
        no_flush  = no_flush.

* title: range content

    rangeitem-name = ls_commentary_range.
    rangeitem-columns = 1.
    rangeitem-rows = 1.
    append rangeitem to ranges.

    contentsitem-row    = li_content_index.
    contentsitem-column = 1.
    concatenate wa_listheader-key
                wa_listheader-info
                into contentsitem-value
                separated by space.
    condense contentsitem-value.
    append contentsitem to contents.

    add 1 to li_content_index.

    clear:
      rangeitem,
      contentsitem,
      ls_index.

  endloop.

* set range data title

  call method lo_spreadsheet->set_ranges_data
    EXPORTING
      ranges   = ranges
      contents = contents
      no_flush = no_flush.

  refresh:
     ranges,
     contents.

  rows = rowmax + li_commentary_rows + 1.

  all = wa_usr-datfm.
  all = all + 3.

  loop at lt_sema into semaitem.
    if semaitem-col_typ = 'DAT' or semaitem-col_typ = 'MON' or
       semaitem-col_typ = 'N00' or semaitem-col_typ = 'N01' or
       semaitem-col_typ = 'N01' or semaitem-col_typ = 'N02' or
       semaitem-col_typ = 'N03' or semaitem-col_typ = 'PCT' or
       semaitem-col_typ = 'STR' or semaitem-col_typ = 'NUM'.
      clear str.
      str = semaitem-col_no.
      condense str.
      concatenate 'DATA' str into str.
      mit = semaitem-col_no.
      li_col_pos = semaitem-col_no + i_left - 1.

* range from data1 to data(n), for each columns of table

      call method lo_spreadsheet->insert_range_dim
        EXPORTING
          name     = str
          top      = rows
          left     = li_col_pos
          rows     = help
          columns  = 1
          no_flush = no_flush.

      data dec type i value -1.
      data typeinfo type sydes_typeinfo.
      loop at <f_excel_tab> assigning <line>.
        assign component semaitem-col_no of structure <line> to <item>.
        describe field <item> into td.
        read table td-types index 1 into typeinfo.
        if typeinfo-type = 'P'.
          dec = typeinfo-decimals.
        elseif typeinfo-type = 'I'.
          dec = 0.
        endif.

        describe field <line> type typ components count.
        mit = 1.
        do count times.
          if mit = semaitem-col_src.
            assign component sy-index of structure <line> to <item>.
            describe field <item> into td.
            read table td-types index 1 into typeinfo.
            if typeinfo-type = 'P'.
              dec = typeinfo-decimals.
            endif.
            exit.
          endif.
          mit = mit + 1.
        enddo.
        exit.
      endloop.

* format for each columns of table (w/o columns headers)

      if semaitem-col_typ = 'DAT'.
        if semaitem-col_no > vkeycount.
          call method lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = all
              no_flush  = no_flush.
        else.
          call method lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 0
              no_flush  = no_flush.
        endif.
      elseif semaitem-col_typ = 'STR'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 0
            no_flush  = no_flush.
      elseif semaitem-col_typ = 'MON'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 10
            no_flush  = no_flush.
      elseif semaitem-col_typ = 'N00'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 1
            decimals  = 0
            no_flush  = no_flush.
      elseif semaitem-col_typ = 'N01'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 1
            decimals  = 1
            no_flush  = no_flush.
      elseif semaitem-col_typ = 'N02'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 1
            decimals  = 2
            no_flush  = no_flush.
      elseif semaitem-col_typ = 'N03'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 1
            decimals  = 3
            no_flush  = no_flush.
      elseif semaitem-col_typ = 'N04'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 1
            decimals  = 4
            no_flush  = no_flush.
      elseif semaitem-col_typ = 'NUM'.
        if dec eq -1.
          call method lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 1
              decimals  = 2
              no_flush  = no_flush.
        else.
          call method lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 1
              decimals  = dec
              no_flush  = no_flush.
        endif.
      elseif semaitem-col_typ = 'PCT'.
        call method lo_spreadsheet->set_format
          EXPORTING
            rangename = str
            currency  = ''
            typ       = 3
            decimals  = 0
            no_flush  = no_flush.
      endif.

    endif.
  endloop.

* get item contents for set_range_data method
* get currency cell also

  mit = 1.

  data: currcells type soi_cell_table.
  data: curritem  type soi_cell_item.

  curritem-rows = 1.
  curritem-columns = 1.
  curritem-front = -1.
  curritem-back = -1.
  curritem-font = ''.
  curritem-size = -1.
  curritem-bold = -1.
  curritem-italic = -1.
  curritem-align = -1.
  curritem-frametyp = -1.
  curritem-framecolor = -1.
  curritem-currency = ''.
  curritem-number = 1.
  curritem-input = -1.

  data: const type i.

*   Change for Correction request
*    Initial 10000 lines are missing in Excel Export
*    if there are only 2 columns in exported List object.

  if datareal gt 2.
    const = 20000 / datareal.
  else.
    const = 20000 / ( datareal + 2 ).
  endif.

  data: lines type i.
  data: innerlines type i.
  data: counter type i.
  data: curritem2 like curritem.
  data: curritem3 like curritem.
  data: length type i.
  data: found.

* append content table (for method set_range_content)

  loop at <f_excel_tab> assigning <line>.

* save line index to compare with lt_subtot_indexs,
* to discover line is a subtotal / totale line or not
* ex use to set 'dun display zero in subtotal / total line'

    l_save_index = sy-tabix.

    do datareal times.
      read table lt_sema into semaitem with key col_no = sy-index.
      if semaitem-col_src ne 0.
        assign component semaitem-col_src
               of structure <line> to <item>.
      else.
        assign component sy-index
               of structure <line> to <item>.
      endif.

      contentsitem-row = realoverflow.

      if sy-subrc = 0.
        move semaitem-col_ops to search_item(3).
        search 'ADD#CNT#MIN#MAX#AVG#NOP#DFT#'
                          for search_item.
        if sy-subrc ne 0.
          raise error_in_sema.
        endif.
        move semaitem-col_typ to search_item(3).
        search 'NUM#N00#N01#N02#N03#N04#PCT#DAT#MON#STR#'
                          for search_item.
        if sy-subrc ne 0.
          raise error_in_sema.
        endif.
        contentsitem-column = sy-index.
        if semaitem-col_typ eq 'DAT' or semaitem-col_typ eq 'MON'.
          if semaitem-col_no > vkeycount.

            " Hinweis 512418
            " EXCEL bezieht Datumsangaben
            " auf den 31.12.1899, behandelt
            " aber 1900 als ein Schaltjahr
            " d.h. ab 1.3.1900 korrekt
            " 1.3.1900 als Zahl = 61

            data: genesis type d value '18991230'.
            data: number_of_days type p.
* change for date in char format & sema_type = X
            data: temp_date type d.

            if not <item> is initial and not <item> co ' ' and not
            <item> co '0'.
* change for date in char format & sema_type = X starts
              if sema_type = 'X'.
                describe field <item> type typ.
                if typ = 'C'.
                  temp_date = <item>.
                  number_of_days = temp_date - genesis.
                else.
                  number_of_days = <item> - genesis.
                endif.
              else.
                number_of_days = <item> - genesis.
              endif.
* change for date in char format & sema_type = X ends
              if number_of_days < 61.
                number_of_days = number_of_days - 1.
              endif.

              set country 'DE'.
              write number_of_days to contentsitem-value
              no-grouping
                                        left-justified.
              set country space.
              translate contentsitem-value using comma_elim.
            else.
              clear contentsitem-value.
            endif.
          else.
            move <item> to contentsitem-value.
          endif.
        elseif semaitem-col_typ eq 'NUM' or
               semaitem-col_typ eq 'N00' or
               semaitem-col_typ eq 'N01' or
               semaitem-col_typ eq 'N02' or
               semaitem-col_typ eq 'N03' or
               semaitem-col_typ eq 'N04' or
               semaitem-col_typ eq 'PCT'.
          set country 'DE'.
          describe field <item> type typ.

          if semaitem-col_cur is initial.
            if typ ne 'F'.
              write <item> to contentsitem-value no-grouping
                                                 no-sign decimals 14.
            else.
              write <item> to contentsitem-value no-grouping
                                                 no-sign.
            endif.
          else.
* Treat of fixed curreny for column >>Y9CK007319
            if semaitem-col_cur < 0.
              semaitem-col_cur = semaitem-col_cur * ( -1 ).
              select * from tcurx into table lt_tcurx.
              sort lt_tcurx.
              read table lt_tcurx into
                                  wa_tcurx index semaitem-col_cur.
              if sy-subrc = 0.
                if typ ne 'F'.
                  write <item> to contentsitem-value no-grouping
                   currency wa_tcurx-currkey no-sign decimals 14.
                else.
                  write <item> to contentsitem-value no-grouping
                   currency wa_tcurx-currkey no-sign.
                endif.
              endif.
            else.
              assign component semaitem-col_cur
                   of structure <line> to <g>.
* mit = index of recent row
              curritem-top  = rowmax + mit + li_commentary_rows.

              li_col_pos =  sy-index + i_left - 1.
              curritem-left = li_col_pos.

* if filed is quantity field (qfieldname ne space)
* or amount field (cfieldname ne space), then format decimal place
* corresponding with config

              clear: l_def.
              read table LT_FIELDCAT_KKBLO assigning <f_fieldcat_line>
              with key  tabname = l_tabname
                        tech    = space
                        no_out  = space
                        col_pos = semaitem-col_no.
              IF sy-subrc = 0.
                IF <f_fieldcat_line>-cfieldname is not initial.
                  l_def = 'C'.
                else."if <f_fieldcat_line>-qfieldname is not initial.
                  l_def = 'Q'.
                ENDIF.
              ENDIF.

* if field is amount field
* exporting of amount field base on currency decimal table: TCURX
              IF l_def = 'C'. "field is amount field
                select single * from tcurx into wa_tcurx
                  where currkey = <g>.
* if amount ref to un-know currency -> default decimal  = 2
                if sy-subrc eq 0.
                  curritem-decimals = wa_tcurx-currdec.
                else.
                  curritem-decimals = 2.
                endif.

                append curritem to currcells.
                if typ ne 'F'.
                  write <item> to contentsitem-value
                                      currency <g>
                     no-sign no-grouping.
                else.
                  write <item> to contentsitem-value
                     decimals 14      currency <g>
                     no-sign no-grouping.
                endif.

* if field is quantity field
* exporting of quantity field base on quantity decimal table: T006

              else."if l_def = 'Q'. " field is quantity field
                clear: wa_t006.
                select single * from t006 into wa_t006
                  where MSEHI = <g>.
* if quantity ref to un-know unit-> default decimal  = 2
                if sy-subrc eq 0.
                  curritem-decimals = wa_t006-decan.
                else.
                  curritem-decimals = 2.
                endif.
                append curritem to currcells.

                write <item> to contentsitem-value
                                    unit <g>
                   no-sign no-grouping.
                condense contentsitem-value.

              ENDIF.

            endif.                                          "Y9CK007319
          endif.
          condense contentsitem-value.

* add function fieldcat-no zero display

          loop at LT_FIELDCAT_KKBLO assigning <f_fieldcat_line>
          where tabname = l_tabname
          and   tech ne 'X'
          and   no_out ne 'X'.
            if <f_fieldcat_line>-col_pos = semaitem-col_no.
              if <f_fieldcat_line>-no_zero = 'X'.
                if <item> = '0'.
                  clear: contentsitem-value.
                endif.

* dun display zero in total/subtotal line too

              else.
                clear: wa_subtot_indexs.
                read table lt_subtot_indexs into wa_subtot_indexs
                with key index = l_save_index.
                IF sy-subrc = 0 AND <item> = '0'.
                  clear: contentsitem-value.
                ENDIF.
              endif.
            endif.
          endloop.
          unassign: <f_fieldcat_line>.

          if <item> lt 0.
            search contentsitem-value for 'E'.
            if sy-fdpos eq 0.

* use prefix notation for signed numbers

              translate contentsitem-value using '- '.
              condense contentsitem-value no-gaps.
              concatenate '-' contentsitem-value
                         into contentsitem-value.
            else.
              concatenate '-' contentsitem-value
                         into contentsitem-value.
            endif.
          endif.
          set country space.
* Hier wird nur die korrekte Kommaseparatierung gemacht, wenn die
* Zeichen einer
* Zahl enthalten sind. Das ist für Timestamps, die auch ":" enthalten.
* Für die
* darf keine Kommaseparierung stattfinden.
* Changing for correction request - Y6BK041073
          if contentsitem-value co '0123456789.,-+E '.
            translate contentsitem-value using comma_elim.
          endif.
        else.
          clear contentsitem-value.

* if type is not numeric -> dun display with zero

          write <item> to contentsitem-value no-zero.

          shift contentsitem-value left deleting leading space.

        endif.
        append contentsitem to contents.
      endif.
    enddo.

    realmit = realmit + 1.
    realoverflow = realoverflow + 1.

    mit = mit + 1.
*   overflow = current row index in content table
    overflow = overflow + 1.
  endloop.

  unassign: <f_fieldcat_line>.

* set item range for set_range_data method

  testname = mit / const.
  condense testname.

  concatenate 'TEST' testname into testname.

  realoverflow = realoverflow - 1.
  realmit = realmit - 1.
  help = realoverflow.

  rangeitem-name = testname.
  rangeitem-columns = datareal.
  rangeitem-rows = help.
  append rangeitem to ranges.

* insert item range dim

  temp2 = rowmax + 1 + li_commentary_rows + realmit - realoverflow.

* items data

  call method lo_spreadsheet->insert_range_dim
    EXPORTING
      name     = testname
      top      = temp2
      left     = i_left
      rows     = help
      columns  = datareal
      no_flush = no_flush.

* get columns header contents for set_range_data method
* export columns header only if no columns header option = space

  data: rowcount type i.
  data: columncount type i.

  if i_columns_header = 'X'.

* append columns header to contents: hkey

    rowcount = 1.
    do rowmax times.
      columncount = 1.
      do hkeycolumns times.
        loop at lt_hkey into hkeyitem where col_no = columncount
                                         and row_no   = rowcount.
        endloop.
        if sy-subrc = 0.
          str = hkeyitem-col_name.
          contentsitem-value = hkeyitem-col_name.
        else.
          contentsitem-value = str.
        endif.
        contentsitem-column = columncount.
        contentsitem-row = rowcount.
        append contentsitem to contents.
        columncount = columncount + 1.
      enddo.
      rowcount = rowcount + 1.
    enddo.

* incase columns header in multiline

    data: rowmaxtemp type i.
    if rowmax > 1.
      rowmaxtemp = rowmax - 1.
      rowcount = 1.
      do rowmaxtemp times.
        columncount = 1.
        do columnmax times.
          contentsitem-column = columncount.
          contentsitem-row    = rowcount.
          contentsitem-value  = ''.
          append contentsitem to contents.
          columncount = columncount + 1.
        enddo.
        rowcount = rowcount + 1.
      enddo.
    endif.

* append columns header to contents: vkey

    columncount = 1.
    do columnmax times.
      loop at lt_vkey into vkeyitem where col_no = columncount.
      endloop.
      contentsitem-value = vkeyitem-col_name.
      contentsitem-row = rowmax.
      contentsitem-column = columncount.
      append contentsitem to contents.
      columncount = columncount + 1.
    enddo.
*--------------------------------------------------------------------*
* set header range for method set_range_data
* insert header keys range dim

    li_head_top = li_commentary_rows + 1.
    li_col_pos = i_left.

* insert range headers

    if hkeycolumns ne 0.
      rangeitem-name = 'TESTHKEY'.
      rangeitem-rows = rowmax.
      rangeitem-columns = hkeycolumns.
      append rangeitem to ranges.
      clear: rangeitem.

      call method lo_spreadsheet->insert_range_dim
        EXPORTING
          name     = 'TESTHKEY'
          top      = li_head_top
          left     = li_col_pos
          rows     = rowmax
          columns  = hkeycolumns
          no_flush = no_flush.
    endif.
  endif.

* format for columns header + total + subtotal
* ------------------------------------------

  help = rowmax + realmit. " table + header lines

  data: lt_format     type soi_format_table.
  data: wa_format     like line of lt_format.
  data: wa_format_temp like line of lt_format.

  field-symbols: <f_source> type any.
  field-symbols: <f_des>    type any.

* columns header format

  wa_format-front       = -1.
  wa_format-back        = 15. "grey
  wa_format-font        = space.
  wa_format-size        = -1.
  wa_format-bold        = 1.
  wa_format-align       = 0.
  wa_format-frametyp    = -1.
  wa_format-framecolor  = -1.

* get column header format from input record
* -> map input format

  if i_columns_header = 'X'.
    wa_format-name        = 'TESTHKEY'.
    if i_format_col_header is not initial.
      describe field i_format_col_header type l_typ components
      li_col_num.
      do li_col_num times.
        if sy-index ne 1. " dun map range name
          assign component sy-index of structure i_format_col_header
          to <f_source>.
          if <f_source> is not initial.
            assign component sy-index of structure wa_format to <f_des>.
            <f_des> = <f_source>.
            unassign: <f_des>.
          endif.
          unassign: <f_source>.
        endif.
      enddo.

      clear: li_col_num.
    endif.

    append wa_format to lt_format.
  endif.

* Zusammenfassen der Spalten mit gleicher Nachkommastellenzahl
* collect vertical cells (col)  with the same number of decimal places
* to increase perfomance in currency cell format

  describe table currcells lines lines.
  lines = lines - 1.
  do lines times.
    describe table currcells lines innerlines.
    innerlines = innerlines - 1.
    sort currcells by left top.
    clear found.
    do innerlines times.
      read table currcells index sy-index into curritem.
      counter = sy-index + 1.
      read table currcells index counter into curritem2.
      if curritem-left eq curritem2-left.
        length = curritem-top + curritem-rows.
        if length eq curritem2-top and curritem-decimals eq curritem2-decimals.
          move curritem to curritem3.
          curritem3-rows = curritem3-rows + curritem2-rows.
          curritem-left = -1.
          modify currcells index sy-index from curritem.
          curritem2-left = -1.
          modify currcells index counter from curritem2.
          append curritem3 to currcells.
          found = 'X'.
        endif.
      endif.
    enddo.
    if found is initial.
      exit.
    endif.
    delete currcells where left = -1.
  enddo.

* Zusammenfassen der Zeilen mit gleicher Nachkommastellenzahl
* collect horizontal cells (row) with the same number of decimal places
* to increase perfomance in currency cell format

  describe table currcells lines lines.
  lines = lines - 1.
  do lines times.
    describe table currcells lines innerlines.
    innerlines = innerlines - 1.
    sort currcells by top left.
    clear found.
    do innerlines times.
      read table currcells index sy-index into curritem.
      counter = sy-index + 1.
      read table currcells index counter into curritem2.
      if curritem-top eq curritem2-top and curritem-rows eq
      curritem2-rows.
        length = curritem-left + curritem-columns.
        if length eq curritem2-left and curritem-decimals eq curritem2-decimals.
          move curritem to curritem3.
          curritem3-columns = curritem3-columns + curritem2-columns.
          curritem-left = -1.
          modify currcells index sy-index from curritem.
          curritem2-left = -1.
          modify currcells index counter from curritem2.
          append curritem3 to currcells.
          found = 'X'.
        endif.
      endif.
    enddo.
    if found is initial.
      exit.
    endif.
    delete currcells where left = -1.
  enddo.
* Ende der Zusammenfassung


* item data: format for currency cell, corresponding with currency

  call method lo_spreadsheet->cell_format
    EXPORTING
      cells    = currcells
      no_flush = no_flush.

* item data: write item table content

  call method lo_spreadsheet->set_ranges_data
    EXPORTING
      ranges   = ranges
      contents = contents
      no_flush = no_flush.

* whole table range to format all table

  if i_columns_header = 'X'.
    li_head_top = li_commentary_rows + 1.
  else.
    li_head_top = li_commentary_rows + 2.
    help = help - 1.
  endif.

  call method lo_spreadsheet->insert_range_dim
    EXPORTING
      name     = 'WHOLE_TABLE'
      top      = li_head_top
      left     = i_left
      rows     = help
      columns  = datareal
      no_flush = no_flush.

* columns width auto fix
* this parameter = space in case use with exist template

  IF i_columns_autofit = 'X'.
    call method lo_spreadsheet->fit_widest
      EXPORTING
        name     = 'WHOLE_TABLE'
        no_flush = no_flush.
  ENDIF.

* frame
* The parameter has 8 bits
*0 Left margin
*1 Top marginT
*2 Bottom margin
*3 Right margin
*4 Horizontal line
*5 Vertical line
*6 Thinness
*7 Thickness
* here 127 = 1111111 6-5-4-3-2-1 mean Thin-ver-hor-right-bot-top-left

* ( final DOI method call, set no_flush = space
* equal to call method CL_GUI_CFW=>FLUSH )

  call method lo_spreadsheet->set_frame
    EXPORTING
      rangename = 'WHOLE_TABLE'
      typ       = 127
      color     = 1
      no_flush  = space
    IMPORTING
      error     = lo_error
      retcode   = lc_retcode.

  error_doi.

* reformat subtotal / total line after format wholw table

  loop at subranges into subrangeitem.
    l_sub_index = subrangeitem-rows + li_commentary_rows + rowmax.

    call method lo_spreadsheet->insert_range_dim
      EXPORTING
        name     = subrangeitem-name
        left     = i_left
        top      = l_sub_index
        rows     = 1
        columns  = datareal
        no_flush = no_flush.

    wa_format-name    = subrangeitem-name.

*   default format:
*     - clolor: subtotal = light yellow, subtotal = yellow
*     - frame: box

    IF  subrangeitem-name(3) = 'SUB'.
      wa_format-back = 36. "subtotal line
      wa_format_temp = i_format_subtotal.
    else.
      wa_format-back = 27. "total line
      wa_format_temp = i_format_total.
    endif.
    wa_format-FRAMETYP = 79.
    wa_format-FRAMEcolor = 1.
    wa_format-number  = -1.
    wa_format-align   = -1.

*   get subtoal + total format from intput parameter
*   overwrite default format

    if wa_format_temp is not initial.
      describe field wa_format_temp type l_typ components li_col_num.
      do li_col_num times.
        if sy-index ne 1. " dun map range name
          assign component sy-index of structure wa_format_temp
          to <f_source>.
          if <f_source> is not initial.
            assign component sy-index of structure wa_format to <f_des>.
            <f_des> = <f_source>.
            unassign: <f_des>.
          endif.
          unassign: <f_source>.
        endif.
      enddo.

      clear: li_col_num.
    endif.

    append wa_format to lt_format.
    clear: wa_format-name.
    clear: l_sub_index.
    clear: wa_format_temp.

  endloop.

  if lt_format[] is not initial.
    call method lo_spreadsheet->set_ranges_format
      EXPORTING
        formattable = lt_format
        no_flush    = no_flush.
    refresh: lt_format.
  endif.
*--------------------------------------------------------------------*
  call method lo_spreadsheet->screen_update
    EXPORTING
      updating = 'X'.

  call method c_oi_errors=>flush_errors.

  lo_error_w = l_error.
  lc_retcode = lo_error_w->error_code.

** catch no_flush -> led to dump ( optional )
*    go_error = l_error.
*    gc_retcode = go_error->error_code.
*    error_doi.

  clear:
    lt_sema,
    wa_sema,
    lt_hkey,
    wa_hkey,
    lt_vkey,
    wa_vkey,
    l_n_hrz_keys,
    l_n_att_cols,
    l_n_vrt_keys,
    count,
    datac,
    datareal,
    vkeycount,
    all,
    mit,
    li_col_pos,
    li_col_num,
    ranges,
    rangeitem,
    contents,
    contentsitem,
    semaitem,
    hkeyitem,
    vkeyitem,
    li_commentary_rows,
    l_retcode,
    li_head_top,
    <f_excel_tab>.

  clear:
     lo_error_w.

  unassign:
  <line>,
  <item>,
  <f_excel_tab>.

*--------------------------------------------------------------------*
* SESSION 5: SAVE AND CLOSE FILE
*--------------------------------------------------------------------*

* ex of save path: 'FILE://C:\temp\test.xlsx'
  concatenate 'FILE://' I_save_path
              into ls_path.

  call method lo_proxy->save_document_to_url
    EXPORTING
      no_flush      = 'X'
      url           = ls_path
    IMPORTING
      error         = lo_error
      retcode       = lc_retcode
    CHANGING
      document_size = li_document_size.

  error_doi.

* if save successfully -> raise successful message
*  message i499(sy) with 'Document is Exported to ' p_path.
  message i499(sy) with 'Data has been exported successfully'.

  clear:
    ls_path,
    li_document_size.

  close_document.
  endmethod.


METHOD bind_table.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmöcker,      (wi p)              2012-12-01
*              - ...
*          aligning code
*          message made to support multilinguality
*--------------------------------------------------------------------*
* issue #237   - Check if overlapping areas exist
*              - Alessandro Iannacci                        2012-12-01
* changes:     - Added raise if overlaps are detected
*--------------------------------------------------------------------*

  CONSTANTS:
              lc_top_left_column              TYPE zexcel_cell_column_alpha VALUE 'A',
              lc_top_left_row                 TYPE zexcel_cell_row VALUE 1.

  DATA:
              lv_row_int                      TYPE zexcel_cell_row,
              lv_first_row                    TYPE zexcel_cell_row,
              lv_last_row                     TYPE zexcel_cell_row,
              lv_column_int                   TYPE zexcel_cell_column,
              lv_column_alpha                 TYPE zexcel_cell_column_alpha,
              lt_field_catalog                TYPE zexcel_t_fieldcatalog,
              lv_id                           TYPE i,
              lv_rows                         TYPE i,
              lv_formula                      TYPE string,
              ls_settings                     TYPE zexcel_s_table_settings,
              lo_table                        TYPE REF TO zcl_excel_table,
              lt_column_name_buffer           TYPE SORTED TABLE OF string WITH UNIQUE KEY table_line,
              lv_value                        TYPE string,
              lv_value_lowercase              TYPE string,
              lv_syindex                      TYPE char3,
              lv_errormessage                 TYPE string,                                            "ins issue #237

              lv_columns                      TYPE i,
              lt_columns                      TYPE zexcel_t_fieldcatalog,
              lv_maxcol                       TYPE i,
              lv_maxrow                       TYPE i,
              lo_iterator                     TYPE REF TO cl_object_collection_iterator,
              lo_style_cond                   TYPE REF TO zcl_excel_style_cond,
              lo_curtable                     TYPE REF TO zcl_excel_table.

  FIELD-SYMBOLS:
              <ls_field_catalog>              TYPE zexcel_s_fieldcatalog,
              <ls_field_catalog_custom>       TYPE zexcel_s_fieldcatalog,
              <fs_table_line>                 TYPE any,
              <fs_fldval>                     TYPE any.

  ls_settings = is_table_settings.

  IF ls_settings-top_left_column IS INITIAL.
    ls_settings-top_left_column = lc_top_left_column.
  ENDIF.

  IF ls_settings-table_style IS INITIAL.
    ls_settings-table_style = zcl_excel_table=>builtinstyle_medium2.
  ENDIF.

  IF ls_settings-top_left_row IS INITIAL.
    ls_settings-top_left_row = lc_top_left_row.
  ENDIF.

  IF it_field_catalog IS NOT SUPPLIED.
    lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = ip_table ).
  ELSE.
    lt_field_catalog = it_field_catalog.
  ENDIF.

  SORT lt_field_catalog BY position.

*--------------------------------------------------------------------*
*  issue #237   Check if overlapping areas exist  Start
*--------------------------------------------------------------------*
  "Get the number of columns for the current table
  lt_columns = lt_field_catalog.
  DELETE lt_columns WHERE dynpfld NE abap_true.
  DESCRIBE TABLE lt_columns LINES lv_columns.

  "Calculate the top left row of the current table
  lv_column_int = zcl_excel_common=>convert_column2int( ls_settings-top_left_column ).
  lv_row_int    = ls_settings-top_left_row.

  "Get number of row for the current table
  DESCRIBE TABLE ip_table LINES lv_rows.

  "Calculate the bottom right row for the current table
  lv_maxcol                       = lv_column_int + lv_columns - 1.
  lv_maxrow                       = lv_row_int    + lv_rows - 1.
  ls_settings-bottom_right_column = zcl_excel_common=>convert_column2alpha( lv_maxcol ).
  ls_settings-bottom_right_row    = lv_maxrow.

  lv_column_int                   = zcl_excel_common=>convert_column2int( ls_settings-top_left_column ).

  lo_iterator = me->tables->if_object_collection~get_iterator( ).
  WHILE lo_iterator->if_object_collection_iterator~has_next( ) EQ abap_true.

    lo_curtable ?= lo_iterator->if_object_collection_iterator~get_next( ).
    IF  (    (  ls_settings-top_left_row     GE lo_curtable->settings-top_left_row                             AND ls_settings-top_left_row     LE lo_curtable->settings-bottom_right_row )
          OR
             (  ls_settings-bottom_right_row GE lo_curtable->settings-top_left_row                             AND ls_settings-bottom_right_row LE lo_curtable->settings-bottom_right_row )
        )
      AND
        (    (  lv_column_int GE zcl_excel_common=>convert_column2int( lo_curtable->settings-top_left_column ) AND lv_column_int LE zcl_excel_common=>convert_column2int( lo_curtable->settings-bottom_right_column ) )
          OR
             (  lv_maxcol     GE zcl_excel_common=>convert_column2int( lo_curtable->settings-top_left_column ) AND lv_maxcol     LE zcl_excel_common=>convert_column2int( lo_curtable->settings-bottom_right_column ) )
        ).
      lv_errormessage = 'Table overlaps with previously bound table and will not be added to worksheet.'(400).
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = lv_errormessage.
    ENDIF.

  ENDWHILE.
*--------------------------------------------------------------------*
*  issue #237   Check if overlapping areas exist  End
*--------------------------------------------------------------------*

  CREATE OBJECT lo_table.
  lo_table->settings = ls_settings.
  lo_table->set_data( ir_data = ip_table ).
  lv_id = me->excel->get_next_table_id( ).
  lo_table->set_id( iv_id = lv_id ).
*  lo_table->fieldcat = lt_field_catalog[].

  me->tables->add( lo_table ).

* It is better to loop column by column (only visible column)
  LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog> WHERE dynpfld EQ abap_true.

    lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

    " Due restrinction of new table object we cannot have two column with the same name
    " Check if a column with the same name exists, if exists add a counter
    " If no medium description is provided we try to use small or long
*    lv_value = <ls_field_catalog>-scrtext_m.
    FIELD-SYMBOLS: <scrtxt1> TYPE any,
                   <scrtxt2> TYPE any,
                   <scrtxt3> TYPE any.

    CASE iv_default_descr.
      WHEN 'M'.
        ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt1>.
        ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt2>.
        ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
      WHEN 'S'.
        ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt1>.
        ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt2>.
        ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
      WHEN 'L'.
        ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt1>.
        ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt2>.
        ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt3>.
      WHEN OTHERS.
        ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt1>.
        ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt2>.
        ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
    ENDCASE.


    IF <scrtxt1> IS NOT INITIAL.
      lv_value = <scrtxt1>.
      <ls_field_catalog>-scrtext_l = lv_value.
    ELSEIF <scrtxt2> IS NOT INITIAL.
      lv_value = <scrtxt2>.
      <ls_field_catalog>-scrtext_l = lv_value.
    ELSEIF <scrtxt3> IS NOT INITIAL.
      lv_value = <scrtxt3>.
      <ls_field_catalog>-scrtext_l = lv_value.
    ELSE.
      lv_value = 'Column'.  " default value as Excel does
      <ls_field_catalog>-scrtext_l = lv_value.
    ENDIF.
    WHILE 1 = 1.
      lv_value_lowercase = lv_value.
      TRANSLATE lv_value_lowercase TO LOWER CASE.
      READ TABLE lt_column_name_buffer TRANSPORTING NO FIELDS WITH KEY table_line = lv_value_lowercase BINARY SEARCH.
      IF sy-subrc <> 0.
        <ls_field_catalog>-scrtext_l = lv_value.
        INSERT lv_value_lowercase INTO TABLE lt_column_name_buffer.
        EXIT.
      ELSE.
        lv_syindex = sy-index.
        CONCATENATE <ls_field_catalog>-scrtext_l lv_syindex INTO lv_value.
      ENDIF.

    ENDWHILE.
    " First of all write column header
    IF <ls_field_catalog>-style_header IS NOT INITIAL.
      me->set_cell( ip_column = lv_column_alpha
                    ip_row    = lv_row_int
                    ip_value  = lv_value
                    ip_style  = <ls_field_catalog>-style_header ).
    ELSE.
      me->set_cell( ip_column = lv_column_alpha
                    ip_row    = lv_row_int
                    ip_value  = lv_value ).
    ENDIF.

    ADD 1 TO lv_row_int.
    LOOP AT ip_table ASSIGNING <fs_table_line>.

      ASSIGN COMPONENT <ls_field_catalog>-fieldname OF STRUCTURE <fs_table_line> TO <fs_fldval>.
      " issue #290 Add formula support in table
      IF <ls_field_catalog>-formula EQ abap_true.
        IF <ls_field_catalog>-style IS NOT INITIAL.
          IF <ls_field_catalog>-abap_type IS NOT INITIAL.
            me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = <fs_fldval>
                        ip_abap_type = <ls_field_catalog>-abap_type
                        ip_style    = <ls_field_catalog>-style ).
          ELSE.
            me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval>
                          ip_style    = <ls_field_catalog>-style ).
          ENDIF.
        ELSEIF <ls_field_catalog>-abap_type IS NOT INITIAL.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = <fs_fldval>
                        ip_abap_type = <ls_field_catalog>-abap_type ).
        ELSE.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = <fs_fldval> ).
        ENDIF.
      ELSE.
        IF <ls_field_catalog>-style IS NOT INITIAL.
          IF <ls_field_catalog>-abap_type IS NOT INITIAL.
            me->set_cell( ip_column = lv_column_alpha
                        ip_row    = lv_row_int
                        ip_value  = <fs_fldval>
                        ip_abap_type = <ls_field_catalog>-abap_type
                        ip_style  = <ls_field_catalog>-style ).
          ELSE.
            me->set_cell( ip_column = lv_column_alpha
                          ip_row    = lv_row_int
                          ip_value  = <fs_fldval>
                          ip_style  = <ls_field_catalog>-style ).
          ENDIF.
        ELSE.
          IF <ls_field_catalog>-abap_type IS NOT INITIAL.
            me->set_cell( ip_column = lv_column_alpha
                        ip_row    = lv_row_int
                        ip_abap_type = <ls_field_catalog>-abap_type
                        ip_value  = <fs_fldval> ).
          ELSE.
            me->set_cell( ip_column = lv_column_alpha
                          ip_row    = lv_row_int
                          ip_value  = <fs_fldval> ).
          ENDIF.
        ENDIF.
      ENDIF.
      ADD 1 TO lv_row_int.

    ENDLOOP.
    IF sy-subrc <> 0. "create empty row if table has no data
      me->set_cell( ip_column = lv_column_alpha
                    ip_row    = lv_row_int
                    ip_value  = space ).
      ADD 1 TO lv_row_int.
    ENDIF.

*--------------------------------------------------------------------*
    " totals
*--------------------------------------------------------------------*
    IF <ls_field_catalog>-totals_function IS NOT INITIAL.
      lv_formula = lo_table->get_totals_formula( ip_column = <ls_field_catalog>-scrtext_l ip_function = <ls_field_catalog>-totals_function ).
      IF <ls_field_catalog>-style_total IS NOT INITIAL.
        me->set_cell( ip_column   = lv_column_alpha
                      ip_row      = lv_row_int
                      ip_formula  = lv_formula
                      ip_style    = <ls_field_catalog>-style_total ).
      ELSE.
        me->set_cell( ip_column   = lv_column_alpha
                      ip_row      = lv_row_int
                      ip_formula  = lv_formula ).
      ENDIF.
    ENDIF.

    lv_row_int = ls_settings-top_left_row.
    ADD 1 TO lv_column_int.

*--------------------------------------------------------------------*
    " conditional formatting
*--------------------------------------------------------------------*
    IF <ls_field_catalog>-style_cond IS NOT INITIAL.
      lv_first_row    = ls_settings-top_left_row + 1. " +1 to exclude header
      lv_last_row     = ls_settings-top_left_row + lv_rows.
      lo_style_cond = me->get_style_cond( <ls_field_catalog>-style_cond ).
      lo_style_cond->set_range( ip_start_column  = lv_column_alpha
                                ip_start_row     = lv_first_row
                                ip_stop_column   = lv_column_alpha
                                ip_stop_row      = lv_last_row ).
    ENDIF.

  ENDLOOP.

*--------------------------------------------------------------------*
  " Set field catalog
*--------------------------------------------------------------------*
  lo_table->fieldcat = lt_field_catalog[].

  es_table_settings = ls_settings.
  es_table_settings-bottom_right_column = lv_column_alpha.
  " >> Issue #291
  IF ip_table IS INITIAL.
    es_table_settings-bottom_right_row    = ls_settings-top_left_row + 2.           "Last rows
  ELSE.
    es_table_settings-bottom_right_row    = ls_settings-top_left_row + lv_rows + 1. "Last rows
  ENDIF.
  " << Issue #291

ENDMETHOD.


METHOD calculate_cell_width.
*--------------------------------------------------------------------*
* issue #293   - Roberto Bianco
*              - Christian Assig                            2014-03-14
*
* changes: - Calculate widths using SAPscript font metrics
*            (transaction SE73)
*          - Calculate the width of dates
*          - Add additional width for auto filter buttons
*          - Add cell padding to simulate Excel behavior
*--------------------------------------------------------------------*

  CONSTANTS:
    lc_default_font_name   TYPE zexcel_style_font_name VALUE 'Calibri', "#EC NOTEXT
    lc_default_font_height TYPE tdfontsize VALUE '110',
    lc_excel_cell_padding  TYPE float VALUE '0.75'.

  DATA: ld_cell_value                 TYPE zexcel_cell_value,
        ld_current_character          TYPE c LENGTH 1,
        ld_style_guid                 TYPE zexcel_cell_style,
        ls_stylemapping               TYPE zexcel_s_stylemapping,
        lo_table_object               TYPE REF TO object,
        lo_table                      TYPE REF TO zcl_excel_table,
        ld_table_top_left_column      TYPE zexcel_cell_column,
        ld_table_bottom_right_column  TYPE zexcel_cell_column,
        ld_flag_contains_auto_filter  TYPE abap_bool VALUE abap_false,
        ld_flag_bold                  TYPE abap_bool VALUE abap_false,
        ld_flag_italic                TYPE abap_bool VALUE abap_false,
        ld_date                       TYPE d,
        ld_date_char                  TYPE c LENGTH 50,
        ld_font_height                TYPE tdfontsize VALUE lc_default_font_height,
        lt_itcfc                      TYPE STANDARD TABLE OF itcfc,
        ld_offset                     TYPE i,
        ld_length                     TYPE i,
        ld_uccp                       TYPE i,
        ls_font_metric                TYPE mty_s_font_metric,
        ld_width_from_font_metrics    TYPE i,
        ld_font_family                TYPE itcfh-tdfamily,
        ld_font_name                  TYPE zexcel_style_font_name VALUE lc_default_font_name,
        lt_font_families              LIKE STANDARD TABLE OF ld_font_family,
        ls_font_cache                 TYPE mty_s_font_cache.

  FIELD-SYMBOLS: <ls_font_cache>  TYPE mty_s_font_cache,
                 <ls_font_metric> TYPE mty_s_font_metric,
                 <ls_itcfc>       TYPE itcfc.

  " Determine cell content and cell style
  me->get_cell( EXPORTING ip_column = ip_column
                          ip_row    = ip_row
                IMPORTING ep_value  = ld_cell_value
                          ep_guid   = ld_style_guid ).

  " ABAP2XLSX uses tables to define areas containing headers and
  " auto-filters. Find out if the current cell is in the header
  " of one of these tables.
  LOOP AT me->tables->collection INTO lo_table_object.
    " Downcast: OBJECT -> ZCL_EXCEL_TABLE
    lo_table ?= lo_table_object.

    " Convert column letters to corresponding integer values
    ld_table_top_left_column =
      zcl_excel_common=>convert_column2int(
        lo_table->settings-top_left_column ).

    ld_table_bottom_right_column =
      zcl_excel_common=>convert_column2int(
        lo_table->settings-bottom_right_column ).

    " Is the current cell part of the table header?
    IF ip_column BETWEEN ld_table_top_left_column AND
                         ld_table_bottom_right_column AND
       ip_row    EQ lo_table->settings-top_left_row.
      " Current cell is part of the table header
      " -> Assume that an auto filter is present and that the font is
      "    bold
      ld_flag_contains_auto_filter = abap_true.
      ld_flag_bold = abap_true.
    ENDIF.
  ENDLOOP.

  " If a style GUID is present, read style attributes
  IF ld_style_guid IS NOT INITIAL.
    TRY.
        " Read style attributes
        ls_stylemapping = me->excel->get_style_to_guid( ld_style_guid ).

        " If the current cell contains the default date format,
        " convert the cell value to a date and calculate its length
        IF ls_stylemapping-complete_style-number_format-format_code =
           zcl_excel_style_number_format=>c_format_date_std.

          " Convert excel date to ABAP date
          ld_date =
            zcl_excel_common=>excel_string_to_date( ld_cell_value ).

          " Format ABAP date using user's formatting settings
          WRITE ld_date TO ld_date_char.

          " Remember the formatted date to calculate the cell size
          ld_cell_value = ld_date_char.

        ENDIF.

        " Read the font size and convert it to the font height
        " used by SAPscript (multiplication by 10)
        IF ls_stylemapping-complete_stylex-font-size = abap_true.
          ld_font_height = ls_stylemapping-complete_style-font-size * 10.
        ENDIF.

        " If set, remember the font name
        IF ls_stylemapping-complete_stylex-font-name = abap_true.
          ld_font_name = ls_stylemapping-complete_style-font-name.
        ENDIF.

        " If set, remember whether font is bold and italic.
        IF ls_stylemapping-complete_stylex-font-bold = abap_true.
          ld_flag_bold = ls_stylemapping-complete_style-font-bold.
        ENDIF.

        IF ls_stylemapping-complete_stylex-font-italic = abap_true.
          ld_flag_italic = ls_stylemapping-complete_style-font-italic.
        ENDIF.

      CATCH zcx_excel.                                  "#EC NO_HANDLER
        " Style GUID is present, but style was not found
        " Continue with default values

    ENDTRY.
  ENDIF.

  " Check if the same font (font name and font attributes) was already
  " used before
  READ TABLE mth_font_cache
    WITH TABLE KEY
      font_name   = ld_font_name
      font_height = ld_font_height
      flag_bold   = ld_flag_bold
      flag_italic = ld_flag_italic
    ASSIGNING <ls_font_cache>.

  IF sy-subrc <> 0.
    " Font is used for the first time
    " Add the font to our local font cache
    ls_font_cache-font_name   = ld_font_name.
    ls_font_cache-font_height = ld_font_height.
    ls_font_cache-flag_bold   = ld_flag_bold.
    ls_font_cache-flag_italic = ld_flag_italic.
    INSERT ls_font_cache INTO TABLE mth_font_cache
      ASSIGNING <ls_font_cache>.

    " Determine the SAPscript font family name from the Excel
    " font name
    SELECT tdfamily
      FROM tfo01
      INTO TABLE lt_font_families
      UP TO 1 ROWS
      WHERE tdtext = ld_font_name
      ORDER BY PRIMARY KEY.

    " Check if a matching font family was found
    " Fonts can be uploaded from TTF files using transaction SE73
    IF lines( lt_font_families ) > 0.
      READ TABLE lt_font_families INDEX 1 INTO ld_font_family.

      " Load font metrics (returns a table with the size of each letter
      " in the font)
      CALL FUNCTION 'LOAD_FONT'
        EXPORTING
          family      = ld_font_family
          height      = ld_font_height
          printer     = 'SWIN'
          bold        = ld_flag_bold
          italic      = ld_flag_italic
        TABLES
          metric      = lt_itcfc
        EXCEPTIONS
          font_family = 1
          codepage    = 2
          device_type = 3
          OTHERS      = 4.
      IF sy-subrc <> 0.
        CLEAR lt_itcfc.
      ENDIF.

      " For faster access, convert each character number to the actual
      " character, and store the characters and their sizes in a hash
      " table
      LOOP AT lt_itcfc ASSIGNING <ls_itcfc>.
        ld_uccp = <ls_itcfc>-cpcharno.
        ls_font_metric-char =
          cl_abap_conv_in_ce=>uccpi( ld_uccp ).
        ls_font_metric-char_width = <ls_itcfc>-tdcwidths.
        INSERT ls_font_metric
          INTO TABLE <ls_font_cache>-th_font_metrics.
      ENDLOOP.

    ENDIF.
  ENDIF.

  " Calculate the cell width
  " If available, use font metrics
  IF lines( <ls_font_cache>-th_font_metrics ) = 0.
    " Font metrics are not available
    " -> Calculate the cell width using only the font size
    ld_length = strlen( ld_cell_value ).
    ep_width = ld_length * ld_font_height / lc_default_font_height + lc_excel_cell_padding.

  ELSE.
    " Font metrics are available

    " Calculate the size of the text by adding the sizes of each
    " letter
    ld_length = strlen( ld_cell_value ).
    DO ld_length TIMES.
      " Subtract 1, because the first character is at offset 0
      ld_offset = sy-index - 1.

      " Read the current character from the cell value
      ld_current_character = ld_cell_value+ld_offset(1).

      " Look up the size of the current letter
      READ TABLE <ls_font_cache>-th_font_metrics
        WITH TABLE KEY char = ld_current_character
        ASSIGNING <ls_font_metric>.
      IF sy-subrc = 0.
        " The size of the letter is known
        " -> Add the actual size of the letter
        ADD <ls_font_metric>-char_width TO ld_width_from_font_metrics.
      ELSE.
        " The size of the letter is unknown
        " -> Add the font height as the default letter size
        ADD ld_font_height TO ld_width_from_font_metrics.
      ENDIF.
    ENDDO.

    " Add cell padding (Excel makes columns a bit wider than the space
    " that is needed for the text itself) and convert unit
    " (division by 100)
    ep_width = ld_width_from_font_metrics / 100 + lc_excel_cell_padding.
  ENDIF.

  " If the current cell contains an auto filter, make it a bit wider.
  " The size used by the auto filter button does not depend on the font
  " size.
  IF ld_flag_contains_auto_filter = abap_true.
    ADD 2 TO ep_width.
  ENDIF.

ENDMETHOD.


METHOD calculate_column_widths.
  TYPES:
  BEGIN OF t_auto_size,
    col_index TYPE int4,
    width     TYPE float,
  END   OF t_auto_size.
  TYPES: tt_auto_size TYPE TABLE OF t_auto_size.

  DATA: lo_column_iterator  TYPE REF TO cl_object_collection_iterator,
        lo_column           TYPE REF TO zcl_excel_column.

  DATA: auto_size   TYPE flag.
  DATA: auto_sizes  TYPE tt_auto_size.
  DATA: count       TYPE int4.
  DATA: highest_row TYPE int4.
  DATA: width       TYPE float.

  FIELD-SYMBOLS: <auto_size>        LIKE LINE OF auto_sizes.

  lo_column_iterator = me->get_columns_iterator( ).
  WHILE lo_column_iterator->has_next( ) = abap_true.
    lo_column ?= lo_column_iterator->get_next( ).
    auto_size = lo_column->get_auto_size( ).
    IF auto_size = abap_true.
      APPEND INITIAL LINE TO auto_sizes ASSIGNING <auto_size>.
      <auto_size>-col_index = lo_column->get_column_index( ).
      <auto_size>-width     = -1.
    ENDIF.
  ENDWHILE.

  " There is only something to do if there are some auto-size columns
  IF NOT auto_sizes IS INITIAL.
    highest_row = me->get_highest_row( ).
    LOOP AT auto_sizes ASSIGNING <auto_size>.
      count = 1.
      WHILE count <= highest_row.
* Do not check merged cells
        IF is_cell_merged(
            ip_column    = <auto_size>-col_index
            ip_row       = count ) = abap_false.
          width = calculate_cell_width( ip_column = <auto_size>-col_index     " issue #155 - less restrictive typing for ip_column
                                        ip_row    = count ).
          IF width > <auto_size>-width.
            <auto_size>-width = width.
          ENDIF.
        ENDIF.
        count = count + 1.
      ENDWHILE.
      lo_column = me->get_column( <auto_size>-col_index ). " issue #155 - less restrictive typing for ip_column
      lo_column->set_width( <auto_size>-width ).
    ENDLOOP.
  ENDIF.

ENDMETHOD.


METHOD change_cell_style.
  " issue # 139
  DATA: stylemapping    TYPE zexcel_s_stylemapping,

        complete_style  TYPE zexcel_s_cstyle_complete,
        complete_stylex TYPE zexcel_s_cstylex_complete,

        borderx         TYPE zexcel_s_cstylex_border,
        l_guid          TYPE zexcel_cell_style.   "issue # 177

  DATA:
    lv_column                 TYPE zexcel_cell_column_alpha,
    lv_column_int             TYPE zexcel_cell_column,
    lv_row                    TYPE zexcel_cell_row,
    lv_column_int_start       TYPE zexcel_cell_column,
    lv_column_int_end         TYPE zexcel_cell_column,
    lv_row_start              TYPE zexcel_cell_row,
    lv_row_end                TYPE zexcel_cell_row.

* We have a lot of parameters.  Use some macros to make the coding more structured

  DEFINE clear_initial_colorxfields.
    if &1-rgb is initial.
      clear &2-rgb.
    endif.
    if &1-indexed is initial.
      clear &2-indexed.
    endif.
    if &1-theme is initial.
      clear &2-theme.
    endif.
    if &1-tint is initial.
      clear &2-tint.
    endif.
  END-OF-DEFINITION.

  DEFINE move_supplied_borders.
    if ip_&1 is supplied.  " only act if parameter was supplied
      if ip_x&1 is supplied.  "
        borderx = ip_x&1.          " use supplied x-parameter
      else.
        clear borderx with 'X'.
* clear in a way that would be expected to work easily
        if ip_&1-border_style is  initial.
          clear borderx-border_style.
        endif.
        clear_initial_colorxfields ip_&1-border_color borderx-border_color.
      endif.
      move-corresponding ip_&1   to complete_style-&2.
      move-corresponding borderx to complete_stylex-&2.
    endif.
  END-OF-DEFINITION.

  DEFINE move_supplied_singlestyles.
    if ip_&1 is supplied.
      complete_style-&2 = ip_&1.
      complete_stylex-&2 = 'X'.
    endif.
  END-OF-DEFINITION.

  " #524 accept Excel-style cell reference parameter IP_CELL
  determine_range(
    EXPORTING ip_cell           = ip_cell
              ip_column         = ip_column
              ip_row            = ip_row
              ip_column_to      = ip_column_to
              ip_row_to         = ip_row_to
    IMPORTING ev_column_i_start = lv_column_int_start
              ev_row_start      = lv_row_start
              ev_column_i_end   = lv_column_int_end
              ev_row_end        = lv_row_end ).

  lv_column_int = lv_column_int_start.
  lv_row        = lv_row_start.

  DO. " loop over each cell in range
    " lv_column, lv_column_int, lv_row contain cell address
    lv_column = zcl_excel_common=>convert_column2alpha( lv_column_int ).

    CLEAR: stylemapping, complete_style, complete_stylex, borderx, l_guid.

*   First get current stylsettings
    TRY.
        me->get_cell( EXPORTING ip_column = lv_column  " Cell Column
                                ip_row    = lv_row      " Cell Row
                      IMPORTING ep_guid   = l_guid )." Cell Value ).  "issue # 177


        stylemapping = me->excel->get_style_to_guid( l_guid ).        "issue # 177
        complete_style  = stylemapping-complete_style.
        complete_stylex = stylemapping-complete_stylex.
      CATCH zcx_excel.
*   Error --> use submitted style
    ENDTRY.

*    move_supplied_multistyles: complete.
    IF ip_complete IS SUPPLIED.
      IF ip_xcomplete IS NOT SUPPLIED.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = 'Complete styleinfo has to be supplied with corresponding X-field'.
      ENDIF.
      MOVE-CORRESPONDING ip_complete  TO complete_style.
      MOVE-CORRESPONDING ip_xcomplete TO complete_stylex.
    ENDIF.



    IF ip_font IS SUPPLIED.
      DATA: fontx LIKE ip_xfont.
      IF ip_xfont IS SUPPLIED.
        fontx = ip_xfont.
      ELSE.
*   Only supplied values should be used - exception: Flags bold and italic strikethrough underline
        MOVE 'X' TO: fontx-bold,
                     fontx-italic,
                     fontx-strikethrough,
                     fontx-underline_mode.
        CLEAR fontx-color WITH 'X'.
        clear_initial_colorxfields ip_font-color fontx-color.
        IF ip_font-family IS NOT INITIAL.
          fontx-family = 'X'.
        ENDIF.
        IF ip_font-name IS NOT INITIAL.
          fontx-name = 'X'.
        ENDIF.
        IF ip_font-scheme IS NOT INITIAL.
          fontx-scheme = 'X'.
        ENDIF.
        IF ip_font-size IS NOT INITIAL.
          fontx-size = 'X'.
        ENDIF.
        IF ip_font-underline_mode IS NOT INITIAL.
          fontx-underline_mode = 'X'.
        ENDIF.
      ENDIF.
      MOVE-CORRESPONDING ip_font  TO complete_style-font.
      MOVE-CORRESPONDING fontx    TO complete_stylex-font.
*   Correction for undeline mode
    ENDIF.

    IF ip_fill IS SUPPLIED.
      DATA: fillx LIKE ip_xfill.
      IF ip_xfill IS SUPPLIED.
        fillx = ip_xfill.
      ELSE.
        CLEAR fillx WITH 'X'.
        IF ip_fill-filltype IS INITIAL.
          CLEAR fillx-filltype.
        ENDIF.
        clear_initial_colorxfields ip_fill-fgcolor fillx-fgcolor.
        clear_initial_colorxfields ip_fill-bgcolor fillx-bgcolor.

      ENDIF.
      MOVE-CORRESPONDING ip_fill  TO complete_style-fill.
      MOVE-CORRESPONDING fillx    TO complete_stylex-fill.
    ENDIF.


    IF ip_borders IS SUPPLIED.
      DATA: bordersx LIKE ip_xborders.
      IF ip_xborders IS SUPPLIED.
        bordersx = ip_xborders.
      ELSE.
        CLEAR bordersx WITH 'X'.
        IF ip_borders-allborders-border_style IS INITIAL.
          CLEAR bordersx-allborders-border_style.
        ENDIF.
        IF ip_borders-diagonal-border_style IS INITIAL.
          CLEAR bordersx-diagonal-border_style.
        ENDIF.
        IF ip_borders-down-border_style IS INITIAL.
          CLEAR bordersx-down-border_style.
        ENDIF.
        IF ip_borders-left-border_style IS INITIAL.
          CLEAR bordersx-left-border_style.
        ENDIF.
        IF ip_borders-right-border_style IS INITIAL.
          CLEAR bordersx-right-border_style.
        ENDIF.
        IF ip_borders-top-border_style IS INITIAL.
          CLEAR bordersx-top-border_style.
        ENDIF.
        clear_initial_colorxfields ip_borders-allborders-border_color bordersx-allborders-border_color.
        clear_initial_colorxfields ip_borders-diagonal-border_color   bordersx-diagonal-border_color.
        clear_initial_colorxfields ip_borders-down-border_color       bordersx-down-border_color.
        clear_initial_colorxfields ip_borders-left-border_color       bordersx-left-border_color.
        clear_initial_colorxfields ip_borders-right-border_color      bordersx-right-border_color.
        clear_initial_colorxfields ip_borders-top-border_color        bordersx-top-border_color.

      ENDIF.
      MOVE-CORRESPONDING ip_borders  TO complete_style-borders.
      MOVE-CORRESPONDING bordersx    TO complete_stylex-borders.
    ENDIF.

    IF ip_alignment IS SUPPLIED.
      DATA: alignmentx LIKE ip_xalignment.
      IF ip_xalignment IS SUPPLIED.
        alignmentx = ip_xalignment.
      ELSE.
        CLEAR alignmentx WITH 'X'.
        IF ip_alignment-horizontal IS INITIAL.
          CLEAR alignmentx-horizontal.
        ENDIF.
        IF ip_alignment-vertical IS INITIAL.
          CLEAR alignmentx-vertical.
        ENDIF.
      ENDIF.
      MOVE-CORRESPONDING ip_alignment  TO complete_style-alignment.
      MOVE-CORRESPONDING alignmentx    TO complete_stylex-alignment.
    ENDIF.

    IF ip_protection IS SUPPLIED.
      MOVE-CORRESPONDING ip_protection  TO complete_style-protection.
      IF ip_xprotection IS SUPPLIED.
        MOVE-CORRESPONDING ip_xprotection TO complete_stylex-protection.
      ELSE.
        IF ip_protection-hidden IS NOT INITIAL.
          complete_stylex-protection-hidden = 'X'.
        ENDIF.
        IF ip_protection-locked IS NOT INITIAL.
          complete_stylex-protection-locked = 'X'.
        ENDIF.
      ENDIF.
    ENDIF.


    move_supplied_borders    : borders_allborders borders-allborders,
                               borders_diagonal   borders-diagonal  ,
                               borders_down       borders-down      ,
                               borders_left       borders-left      ,
                               borders_right      borders-right     ,
                               borders_top        borders-top       .

    move_supplied_singlestyles: number_format_format_code     number_format-format_code,
                                font_bold                     font-bold,
                                font_color                    font-color,
                                font_color_rgb                font-color-rgb,
                                font_color_indexed            font-color-indexed,
                                font_color_theme              font-color-theme,
                                font_color_tint               font-color-tint,

                                font_family                   font-family,
                                font_italic                   font-italic,
                                font_name                     font-name,
                                font_scheme                   font-scheme,
                                font_size                     font-size,
                                font_strikethrough            font-strikethrough,
                                font_underline                font-underline,
                                font_underline_mode           font-underline_mode,
                                fill_filltype                 fill-filltype,
                                fill_rotation                 fill-rotation,
                                fill_fgcolor                  fill-fgcolor,
                                fill_fgcolor_rgb              fill-fgcolor-rgb,
                                fill_fgcolor_indexed          fill-fgcolor-indexed,
                                fill_fgcolor_theme            fill-fgcolor-theme,
                                fill_fgcolor_tint             fill-fgcolor-tint,

                                fill_bgcolor                  fill-bgcolor,
                                fill_bgcolor_rgb              fill-bgcolor-rgb,
                                fill_bgcolor_indexed          fill-bgcolor-indexed,
                                fill_bgcolor_theme            fill-bgcolor-theme,
                                fill_bgcolor_tint             fill-bgcolor-tint,

                                fill_gradtype_type            fill-gradtype-TYPE,
                                fill_gradtype_degree          fill-gradtype-DEGREE,
                                fill_gradtype_bottom          fill-gradtype-BOTTOM,
                                fill_gradtype_left            fill-gradtype-LEFT,
                                fill_gradtype_top             fill-gradtype-TOP,
                                fill_gradtype_right           fill-gradtype-RIGHT,
                                fill_gradtype_position1       fill-gradtype-POSITION1,
                                fill_gradtype_position2       fill-gradtype-POSITION2,
                                fill_gradtype_position3       fill-gradtype-POSITION3,



                                borders_diagonal_mode         borders-diagonal_mode,
                                alignment_horizontal          alignment-horizontal,
                                alignment_vertical            alignment-vertical,
                                alignment_textrotation        alignment-textrotation,
                                alignment_wraptext            alignment-wraptext,
                                alignment_shrinktofit         alignment-shrinktofit,
                                alignment_indent              alignment-indent,
                                protection_hidden             protection-hidden,
                                protection_locked             protection-locked,

                                borders_allborders_style      borders-allborders-border_style,
                                borders_allborders_color      borders-allborders-border_color,
                                borders_allbo_color_rgb       borders-allborders-border_color-rgb,
                                borders_allbo_color_indexed   borders-allborders-border_color-indexed,
                                borders_allbo_color_theme     borders-allborders-border_color-theme,
                                borders_allbo_color_tint      borders-allborders-border_color-tint,

                                borders_diagonal_style        borders-diagonal-border_style,
                                borders_diagonal_color        borders-diagonal-border_color,
                                borders_diagonal_color_rgb    borders-diagonal-border_color-rgb,
                                borders_diagonal_color_inde   borders-diagonal-border_color-indexed,
                                borders_diagonal_color_them   borders-diagonal-border_color-theme,
                                borders_diagonal_color_tint   borders-diagonal-border_color-tint,

                                borders_down_style            borders-down-border_style,
                                borders_down_color            borders-down-border_color,
                                borders_down_color_rgb        borders-down-border_color-rgb,
                                borders_down_color_indexed    borders-down-border_color-indexed,
                                borders_down_color_theme      borders-down-border_color-theme,
                                borders_down_color_tint       borders-down-border_color-tint,

                                borders_left_style            borders-left-border_style,
                                borders_left_color            borders-left-border_color,
                                borders_left_color_rgb        borders-left-border_color-rgb,
                                borders_left_color_indexed    borders-left-border_color-indexed,
                                borders_left_color_theme      borders-left-border_color-theme,
                                borders_left_color_tint       borders-left-border_color-tint,

                                borders_right_style           borders-right-border_style,
                                borders_right_color           borders-right-border_color,
                                borders_right_color_rgb       borders-right-border_color-rgb,
                                borders_right_color_indexed   borders-right-border_color-indexed,
                                borders_right_color_theme     borders-right-border_color-theme,
                                borders_right_color_tint      borders-right-border_color-tint,

                                borders_top_style             borders-top-border_style,
                                borders_top_color             borders-top-border_color,
                                borders_top_color_rgb         borders-top-border_color-rgb,
                                borders_top_color_indexed     borders-top-border_color-indexed,
                                borders_top_color_theme       borders-top-border_color-theme,
                                borders_top_color_tint        borders-top-border_color-tint.


*   Now we have a completly filled styles.
*   This can be used to get the guid
*   Return guid if requested.  Might be used if copy&paste of styles is requested
    ep_guid = me->excel->get_static_cellstyle_guid( ip_cstyle_complete  = complete_style
                                                    ip_cstylex_complete = complete_stylex  ).
    me->set_cell_style( ip_column = lv_column
                        ip_row    = lv_row
                        ip_style  = ep_guid ).

    " move to next column
    lv_column_int = lv_column_int + 1.
    " move to new row if column is out of range
    IF lv_column_int > lv_column_int_end.
      lv_column_int = lv_column_int_start.
      lv_row = lv_row + 1.
    ENDIF.
    " exit if row is out of range
    IF lv_row > lv_row_end.
      EXIT. " DO
    ENDIF.
  ENDDO.

ENDMETHOD.


  method CHECK_RTF.

    DATA: lv_rtf_length TYPE I.
    DATA: lv_value      TYPE STRING.
    DATA: lv_val_length TYPE I.
    FIELD-SYMBOLS: <rtf> LIKE LINE OF IT_RTF.

    CHECK it_rtf[] IS NOT INITIAL.

    CLEAR: lv_rtf_length.
    LOOP AT IT_RTF ASSIGNING <rtf>.
      IF lv_rtf_length <> <rtf>-offset.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = 'Gaps or overlaps in RTF data offset/length specs'.
      ENDIF.
      lv_rtf_length = <rtf>-offset + <rtf>-length.
    ENDLOOP.

    lv_value = ip_value.
    lv_val_length = strlen( lv_value ).
    IF lv_val_length <> lv_rtf_length.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'RTF specs length is not equal to value length'.
    ENDIF.

  endmethod.


METHOD constructor.
  DATA: lv_title TYPE zexcel_sheet_title.

  me->excel = ip_excel.

*  CALL FUNCTION 'GUID_CREATE'                                    " del issue #379 - function is outdated in newer releases
*    IMPORTING
*      ev_guid_16 = me->guid.
  me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).        " ins issue #379 - replacement for outdated function call

  IF ip_title IS NOT INITIAL.
    lv_title = ip_title.
  ELSE.
*    lv_title = me->guid.             " del issue #154 - Names of worksheets
    lv_title = me->generate_title( ). " ins issue #154 - Names of worksheets
  ENDIF.

  me->set_title( ip_title = lv_title ).

  CREATE OBJECT sheet_setup.
  CREATE OBJECT styles_cond.
  CREATE OBJECT data_validations.
  CREATE OBJECT tables.
  CREATE OBJECT columns.
  CREATE OBJECT rows.
  CREATE OBJECT ranges. " issue #163
  CREATE OBJECT mo_pagebreaks.
  CREATE OBJECT drawings
    EXPORTING
      ip_type = zcl_excel_drawing=>type_image.
  CREATE OBJECT charts
    EXPORTING
      ip_type = zcl_excel_drawing=>type_chart.
  me->zif_excel_sheet_protection~initialize( ).
  me->zif_excel_sheet_properties~initialize( ).
  CREATE OBJECT hyperlinks.

* initialize active cell coordinates
  active_cell-cell_row = 1.
  active_cell-cell_column = 1.

* inizialize dimension range
  lower_cell-cell_row     = 1.
  lower_cell-cell_column  = 1.
  upper_cell-cell_row     = 1.
  upper_cell-cell_column  = 1.

ENDMETHOD.


METHOD delete_merge.

  DATA: lv_column TYPE i.
*--------------------------------------------------------------------*
* If cell information is passed delete merge including this cell,
* otherwise delete all merges
*--------------------------------------------------------------------*
  IF   ip_cell_column IS INITIAL
    OR ip_cell_row    IS INITIAL.
    CLEAR me->mt_merged_cells.
  ELSE.
    lv_column = zcl_excel_common=>convert_column2int( ip_cell_column ).

    LOOP AT me->mt_merged_cells TRANSPORTING NO FIELDS
    WHERE
        ( row_from <= ip_cell_row and row_to >= ip_cell_row )
    AND
        ( col_from <= lv_column and col_to >= lv_column ).

      DELETE me->mt_merged_cells.
      EXIT.
    ENDLOOP.
  ENDIF.

ENDMETHOD.


METHOD delete_row_outline.

  DELETE me->mt_row_outlines WHERE row_from = iv_row_from
                               AND row_to   = iv_row_to.
  IF sy-subrc <> 0.  " didn't find outline that was to be deleted
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Row outline to be deleted does not exist'.
  ENDIF.

ENDMETHOD.


  method DETERMINE_RANGE.

    IF NOT ( ( ip_column IS INITIAL AND ip_row IS INITIAL ) OR ip_cell IS INITIAL ).
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Please provide either row and column, or cell/range reference'.
    ENDIF.

    IF ip_cell IS NOT INITIAL.
      zcl_excel_common=>convert_range2column_a_row(
        EXPORTING
          i_range = ip_cell
        IMPORTING
          e_column_start = ev_column_a_start
          e_column_end   = ev_column_a_end
          e_row_start    = ev_row_start
          e_row_end      = ev_row_end  ).
    ELSE.
      ev_column_a_start = ip_column.
      ev_row_start      = ip_row.
      ev_column_a_end   = ip_column_to.
      ev_row_end        = ip_row_to.
    ENDIF.

    IF ev_column_a_end IS INITIAL.
      ev_column_a_end = ev_column_a_start.
    ENDIF.

    IF ev_row_end IS INITIAL.
      ev_row_end = ev_row_start.
    ENDIF.

    ev_column_i_start = zcl_excel_common=>convert_column2int( ev_column_a_start ).
    ev_column_i_end   = zcl_excel_common=>convert_column2int( ev_column_a_end ).

    IF ev_column_i_start > ev_column_i_end.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Start column may not be greater than end column'.
    ENDIF.

    IF ev_row_start > ev_row_end.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Start row may not be greater than end row'.
    ENDIF.

  endmethod.


method FREEZE_PANES.

  IF ip_num_columns IS NOT SUPPLIED AND ip_num_rows IS NOT SUPPLIED.
    RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = 'Pleas provide number of rows and/or columns to freeze'.
  ENDIF.

  IF ip_num_columns IS SUPPLIED AND ip_num_columns <= 0.
    RAISE EXCEPTION TYPE zcx_excel
              EXPORTING
                error = 'Number of columns to freeze should be positive'.
  ENDIF.

  IF ip_num_rows IS SUPPLIED AND ip_num_rows <= 0.
    RAISE EXCEPTION TYPE zcx_excel
              EXPORTING
                error = 'Number of rows to freeze should be positive'.
  ENDIF.

  freeze_pane_cell_column = ip_num_columns + 1.
  freeze_pane_cell_row = ip_num_rows + 1.
  endmethod.


method GENERATE_TITLE.
  DATA: lo_worksheets_iterator  TYPE REF TO cl_object_collection_iterator,
        lo_worksheet            TYPE REF TO zcl_excel_worksheet.

  DATA: t_titles                TYPE HASHED TABLE OF zexcel_sheet_title WITH UNIQUE KEY table_line,
        title                   TYPE zexcel_sheet_title,
        sheetnumber             TYPE i.

* Get list of currently used titles
  lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
  WHILE lo_worksheets_iterator->has_next( ) = abap_true.
    lo_worksheet ?= lo_worksheets_iterator->get_next( ).
    title = lo_worksheet->get_title( ).
    INSERT title INTO TABLE t_titles.
    ADD 1 TO sheetnumber.
  ENDWHILE.

* Now build sheetnumber.  Increase counter until we hit a number that is not used so far
  ADD 1 TO sheetnumber.  " Start counting with next number
  DO.
    title = sheetnumber.
    SHIFT title LEFT DELETING LEADING space.
    CONCATENATE 'Sheet'(001) title INTO ep_title.
    INSERT ep_title INTO TABLE t_titles.
    IF sy-subrc = 0.  " Title not used so far --> take it
      EXIT.
    ENDIF.

    ADD 1 TO sheetnumber.
  ENDDO.
  endmethod.


method GET_ACTIVE_CELL.

  DATA: lv_active_column TYPE zexcel_cell_column_alpha,
        lv_active_row    TYPE string.

  lv_active_column = zcl_excel_common=>convert_column2alpha( active_cell-cell_column ).
  lv_active_row    = active_cell-cell_row.
  SHIFT lv_active_row RIGHT DELETING TRAILING space.
  SHIFT lv_active_row LEFT DELETING LEADING space.
  CONCATENATE lv_active_column lv_active_row INTO ep_active_cell.

  endmethod.


METHOD get_cell.

  DATA: lv_column         TYPE zexcel_cell_column,
        lv_column_alpha   TYPE zexcel_cell_column_alpha,
        lv_row            TYPE zexcel_cell_row,
        ls_sheet_content  TYPE zexcel_s_cell_data.

" #524 accept Excel-style cell reference parameter IP_CELL
  determine_range(
    EXPORTING ip_cell           = ip_cell
              ip_column         = ip_column
              ip_row            = ip_row
    IMPORTING ev_column_i_start = lv_column
              ev_column_a_start = lv_column_alpha
              ev_row_start      = lv_row ).

  READ TABLE sheet_content INTO ls_sheet_content WITH TABLE KEY cell_row     = lv_row
                                                                cell_column  = lv_column.

  ep_rc = sy-subrc.
  ep_value    = ls_sheet_content-cell_value.
  ep_guid     = ls_sheet_content-cell_style.       " issue 139 - added this to be used for columnwidth calculation
  ep_formula  = ls_sheet_content-cell_formula.
  IF et_rtf IS SUPPLIED AND ls_sheet_content-rtf_tab IS BOUND. " #486 - rich text formating data
    et_rtf = ls_sheet_content-rtf_tab->*.
  ENDIF.

  " Addition to solve issue #120, contribution by Stefan Schmöcker
  DATA: style_iterator TYPE REF TO cl_object_collection_iterator,
        style          TYPE REF TO zcl_excel_style.
  IF ep_style IS REQUESTED.
    style_iterator = me->excel->get_styles_iterator( ).
    WHILE style_iterator->has_next( ) = 'X'.
      style ?= style_iterator->get_next( ).
      IF style->get_guid( ) = ls_sheet_content-cell_style.
        ep_style = style.
        EXIT.
      ENDIF.
    ENDWHILE.
  ENDIF.
ENDMETHOD.


METHOD get_column.

  DATA: lo_column_iterator TYPE REF TO cl_object_collection_iterator,
        lo_column          TYPE REF TO zcl_excel_column,
        lv_column          TYPE zexcel_cell_column.

  lv_column = zcl_excel_common=>convert_column2int( ip_column ).

  lo_column_iterator = me->get_columns_iterator( ).
  WHILE lo_column_iterator->has_next( ) = abap_true.
    lo_column ?= lo_column_iterator->get_next( ).
    IF lo_column->get_column_index( ) = lv_column.
      eo_column = lo_column.
      EXIT.
    ENDIF.
  ENDWHILE.

  IF eo_column IS NOT BOUND.
    eo_column = me->add_new_column( ip_column ).
  ENDIF.

ENDMETHOD.


METHOD get_columns.
  eo_columns = me->columns.
ENDMETHOD.


METHOD GET_COLUMNS_ITERATOR.

  eo_iterator = me->columns->get_iterator( ).

ENDMETHOD.


method GET_DATA_VALIDATIONS_ITERATOR.

  eo_iterator = me->data_validations->get_iterator( ).
  endmethod.


method GET_DATA_VALIDATIONS_SIZE.
  ep_size = me->data_validations->size( ).
  endmethod.


METHOD GET_DEFAULT_COLUMN.
  IF me->column_default IS NOT BOUND.
    CREATE OBJECT me->column_default
      EXPORTING
        ip_index     = 'A'         " ????
        ip_worksheet = me
        ip_excel     = me->excel.
  ENDIF.

  eo_column = me->column_default.
ENDMETHOD.


method GET_DEFAULT_EXCEL_DATE_FORMAT.
  CONSTANTS: c_lang_e TYPE lang VALUE 'E'.

  IF default_excel_date_format IS NOT INITIAL.
    ep_default_excel_date_format = default_excel_date_format.
    RETURN.
  ENDIF.

  "try to get defaults
  TRY.
      cl_abap_datfm=>get_date_format_des( EXPORTING im_langu = c_lang_e
                                          IMPORTING ex_dateformat = default_excel_date_format ).
    CATCH cx_abap_datfm_format_unknown.

  ENDTRY.

  " and fallback to fixed format
  IF default_excel_date_format IS INITIAL.
    default_excel_date_format = zcl_excel_style_number_format=>c_format_date_ddmmyyyydot.
  ENDIF.

  ep_default_excel_date_format = default_excel_date_format.
  endmethod.


method GET_DEFAULT_EXCEL_TIME_FORMAT.
  DATA: l_timefm TYPE xutimefm.

  IF default_excel_time_format IS NOT INITIAL.
    ep_default_excel_time_format = default_excel_time_format.
    RETURN.
  ENDIF.

* Let's get default
  l_timefm = cl_abap_timefm=>get_environment_timefm( ).
  CASE l_timefm.
    WHEN 0.
*0  24 Hour Format (Example: 12:05:10)
      default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
    WHEN 1.
*1  12 Hour Format (Example: 12:05:10 PM)
      default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
    WHEN 2.
*2  12 Hour Format (Example: 12:05:10 pm) for now all the same. no chnage upper lower
      default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
    WHEN 3.
*3  Hours from 0 to 11 (Example: 00:05:10 PM)  for now all the same. no chnage upper lower
      default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
    WHEN 4.
*4  Hours from 0 to 11 (Example: 00:05:10 pm)  for now all the same. no chnage upper lower
      default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
    WHEN OTHERS.
      " and fallback to fixed format
      default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
  ENDCASE.

  ep_default_excel_time_format = default_excel_time_format.
  endmethod.


METHOD get_default_row.
  IF me->row_default IS NOT BOUND.
    CREATE OBJECT me->row_default.
  ENDIF.

  eo_row = me->row_default.
ENDMETHOD.


method GET_DIMENSION_RANGE.

  me->update_dimension_range( ).
  IF upper_cell EQ lower_cell. "only one cell
    " Worksheet not filled
*    IF upper_cell-cell_coords = '0'.
    IF upper_cell-cell_coords IS INITIAL.
      ep_dimension_range = 'A1'.
    ELSE.
      ep_dimension_range = upper_cell-cell_coords.
    ENDIF.
  ELSE.
    CONCATENATE upper_cell-cell_coords ':' lower_cell-cell_coords INTO ep_dimension_range.
  ENDIF.

  endmethod.


method GET_DRAWINGS.

  DATA: lo_drawing                    TYPE REF TO zcl_excel_drawing,
        lo_iterator                   TYPE REF TO cl_object_collection_iterator.

  CASE ip_type.
    WHEN zcl_excel_drawing=>type_image.
      r_drawings = drawings.
    WHEN zcl_excel_drawing=>type_chart.
      r_drawings = charts.
    WHEN space.
      CREATE OBJECT r_drawings
        EXPORTING
          ip_type = ''.

      lo_iterator = drawings->get_iterator( ).
      WHILE lo_iterator->has_next( ) = abap_true.
        lo_drawing ?= lo_iterator->get_next( ).
        r_drawings->include( lo_drawing ).
      ENDWHILE.
      lo_iterator = charts->get_iterator( ).
      WHILE lo_iterator->has_next( ) = abap_true.
        lo_drawing ?= lo_iterator->get_next( ).
        r_drawings->include( lo_drawing ).
      ENDWHILE.
    WHEN OTHERS.
  ENDCASE.
  endmethod.


method GET_DRAWINGS_ITERATOR.
  CASE ip_type.
    WHEN zcl_excel_drawing=>type_image.
      eo_iterator = drawings->get_iterator( ).
    WHEN zcl_excel_drawing=>type_chart.
      eo_iterator = charts->get_iterator( ).
  ENDCASE.
  endmethod.


method GET_FREEZE_CELL.
  ep_row = me->freeze_pane_cell_row.
  ep_column = me->freeze_pane_cell_column.
  endmethod.


METHOD get_guid.

  ep_guid = me->guid.

ENDMETHOD.


method GET_HIGHEST_COLUMN.
  me->update_dimension_range( ).
  r_highest_column = me->lower_cell-cell_column.
  endmethod.


METHOD get_highest_row.
  me->update_dimension_range( ).
  r_highest_row = me->lower_cell-cell_row.
ENDMETHOD.


method GET_HYPERLINKS_ITERATOR.
  eo_iterator = hyperlinks->get_iterator( ).
  endmethod.


method GET_HYPERLINKS_SIZE.
  ep_size = hyperlinks->size( ).
  endmethod.


METHOD get_merge.

  FIELD-SYMBOLS: <ls_merged_cell> LIKE LINE OF me->mt_merged_cells.

  DATA: lv_col_from    TYPE string,
        lv_col_to      TYPE string,
        lv_row_from    TYPE string,
        lv_row_to      TYPE string,
        lv_merge_range TYPE string.

  LOOP AT me->mt_merged_cells ASSIGNING <ls_merged_cell>.

    lv_col_from = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_from ).
    lv_col_to   = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_to   ).
    lv_row_from = <ls_merged_cell>-row_from.
    lv_row_to   = <ls_merged_cell>-row_to  .
    CONCATENATE lv_col_from lv_row_from ':' lv_col_to lv_row_to
       INTO lv_merge_range.
    CONDENSE lv_merge_range NO-GAPS.
    APPEND lv_merge_range TO merge_range.

  ENDLOOP.

ENDMETHOD.


method GET_PAGEBREAKS.
  ro_pagebreaks = mo_pagebreaks.
endmethod.


method GET_RANGES_ITERATOR.

  eo_iterator = me->ranges->get_iterator( ).

  endmethod.


METHOD get_row.

  DATA: lo_row_iterator TYPE REF TO cl_object_collection_iterator,
        lo_row          TYPE REF TO zcl_excel_row.

  lo_row_iterator = me->get_rows_iterator( ).
  WHILE lo_row_iterator->has_next( ) = abap_true.
    lo_row ?= lo_row_iterator->get_next( ).
    IF lo_row->get_row_index( ) = ip_row.
      eo_row = lo_row.
      EXIT.
    ENDIF.
  ENDWHILE.

  IF eo_row IS NOT BOUND.
    eo_row = me->add_new_row( ip_row ).
  ENDIF.

ENDMETHOD.


METHOD GET_ROWS.
  eo_rows = me->rows.
ENDMETHOD.


METHOD get_rows_iterator.

  eo_iterator = me->rows->get_iterator( ).

ENDMETHOD.


METHOD get_row_outlines.

  rt_row_outlines = me->mt_row_outlines.

ENDMETHOD.


METHOD get_style_cond.

  DATA: lo_style_iterator TYPE REF TO cl_object_collection_iterator,
        lo_style_cond     TYPE REF TO zcl_excel_style_cond.

  lo_style_iterator = me->get_style_cond_iterator( ).
  WHILE lo_style_iterator->has_next( ) = abap_true.
    lo_style_cond ?= lo_style_iterator->get_next( ).
    IF lo_style_cond->get_guid( ) = ip_guid.
      eo_style_cond = lo_style_cond.
      EXIT.
    ENDIF.
  ENDWHILE.

ENDMETHOD.


METHOD get_style_cond_iterator.

  eo_iterator = styles_cond->get_iterator( ).
ENDMETHOD.


method GET_TABCOLOR.
  ev_tabcolor = me->tabcolor.
  endmethod.


   METHOD get_table.
*--------------------------------------------------------------------*
* Comment D. Rauchenstein
* With this method, we get a fully functional Excel Upload, which solves
* a few issues of the other excel upload tools
* ZBCABA_ALSM_EXCEL_UPLOAD_EXT: Reads only up to 50 signs per Cell, Limit
* in row-Numbers. Other have Limitations of Lines, or you are not able
* to ignore filters or choosing the right tab.
*
* To get a fully functional XLSX Upload, you can use it e.g. with method
* CL_EXCEL_READER_2007->ZIF_EXCEL_READER~LOAD_FILE()
*--------------------------------------------------------------------*

     FIELD-SYMBOLS: <ls_line> TYPE data.
     FIELD-SYMBOLS: <lv_value> TYPE data.

     DATA lv_actual_row TYPE int4.
     DATA lv_actual_col TYPE int4.
     DATA lv_errormessage TYPE string.
     DATA lv_max_col TYPE zexcel_cell_column.
     DATA lv_max_row TYPE int4.
     DATA lv_delta_col TYPE int4.
     DATA lv_value  TYPE zexcel_cell_value.
     DATA lv_rc	TYPE sysubrc.


     lv_max_col =  me->get_highest_column( ).
     lv_max_row =  me->get_highest_row( ).

*--------------------------------------------------------------------*
* The row counter begins with 1 and should be corrected with the skips
*--------------------------------------------------------------------*
     lv_actual_row =  iv_skipped_rows + 1.
     lv_actual_col =  iv_skipped_cols + 1.


     TRY.
*--------------------------------------------------------------------*
* Check if we the basic features are possible with given "any table"
*--------------------------------------------------------------------*
         APPEND INITIAL LINE TO et_table ASSIGNING <ls_line>.
         IF sy-subrc <> 0 OR <ls_line> IS NOT ASSIGNED.

           lv_errormessage = 'Error at inserting new Line to internal Table'(002).
           RAISE EXCEPTION TYPE zcx_excel
             EXPORTING
               error = lv_errormessage.

         ELSE.
           lv_delta_col = lv_max_col - iv_skipped_cols.
           ASSIGN COMPONENT lv_delta_col OF STRUCTURE <ls_line> TO <lv_value>.
           IF sy-subrc <> 0 OR <lv_value> IS NOT ASSIGNED.
             lv_errormessage = 'Internal table has less columns than excel'(003).
             RAISE EXCEPTION TYPE zcx_excel
               EXPORTING
                 error = lv_errormessage.
           ELSE.
*--------------------------------------------------------------------*
*now we are ready for handle the table data
*--------------------------------------------------------------------*
             REFRESH et_table.
*--------------------------------------------------------------------*
* Handle each Row until end on right side
*--------------------------------------------------------------------*
             WHILE lv_actual_row <= lv_max_row .

*--------------------------------------------------------------------*
* Handle each Column until end on bottom
* First step is to step back on first column
*--------------------------------------------------------------------*
               lv_actual_col =  iv_skipped_cols + 1.

               UNASSIGN <ls_line>.
               APPEND INITIAL LINE TO et_table ASSIGNING <ls_line>.
               IF sy-subrc <> 0 OR <ls_line> IS NOT ASSIGNED.
                 lv_errormessage = 'Error at inserting new Line to internal Table'(002).
                 RAISE EXCEPTION TYPE zcx_excel
                   EXPORTING
                     error = lv_errormessage.
               ENDIF.
               WHILE lv_actual_col <= lv_max_col.

                 lv_delta_col = lv_actual_col - iv_skipped_cols.
                 ASSIGN COMPONENT lv_delta_col OF STRUCTURE <ls_line> TO <lv_value>.
                 IF sy-subrc <> 0.
                   lv_errormessage = |{ 'Error at assigning field (Col:'(004) } { lv_actual_col } { ' Row:'(005) } { lv_actual_row }|.
                   RAISE EXCEPTION TYPE zcx_excel
                     EXPORTING
                       error = lv_errormessage.
                 ENDIF.

                 me->get_cell(
                   EXPORTING
                     ip_column  = lv_actual_col    " Cell Column
                     ip_row     = lv_actual_row    " Cell Row
                   IMPORTING
                     ep_value   = lv_value    " Cell Value
                     ep_rc      = lv_rc    " Return Value of ABAP Statements
                 ).
                 IF lv_rc <> 0
                   AND lv_rc <> 4.                                                   "No found error means, zero/no value in cell
                   lv_errormessage = |{ 'Error at reading field value (Col:'(007) } { lv_actual_col } { ' Row:'(005) } { lv_actual_row }|.
                   RAISE EXCEPTION TYPE zcx_excel
                     EXPORTING
                       error = lv_errormessage.
                 ENDIF.

                 <lv_value> = lv_value.
*  CATCH zcx_excel.    "
                 ADD 1 TO lv_actual_col.
               ENDWHILE.
               ADD 1 TO lv_actual_row.
             ENDWHILE.
           ENDIF.


         ENDIF.

       CATCH cx_sy_assign_cast_illegal_cast.
         lv_errormessage = |{ 'Error at assigning field (Col:'(004) } { lv_actual_col } { ' Row:'(005) } { lv_actual_row }|.
         RAISE EXCEPTION TYPE zcx_excel
           EXPORTING
             error = lv_errormessage.
       CATCH cx_sy_assign_cast_unknown_type.
         lv_errormessage = |{ 'Error at assigning field (Col:'(004) } { lv_actual_col } { ' Row:'(005) } { lv_actual_row }|.
         RAISE EXCEPTION TYPE zcx_excel
           EXPORTING
             error = lv_errormessage.
       CATCH cx_sy_assign_out_of_range.
         lv_errormessage = 'Internal table has less columns than excel'(003).
         RAISE EXCEPTION TYPE zcx_excel
           EXPORTING
             error = lv_errormessage.
       CATCH cx_sy_conversion_error.
         lv_errormessage = |{ 'Error at converting field value (Col:'(006) } { lv_actual_col } { ' Row:'(005) } { lv_actual_row }|.
         RAISE EXCEPTION TYPE zcx_excel
           EXPORTING
             error = lv_errormessage.

     ENDTRY.
   ENDMETHOD.


method GET_TABLES_ITERATOR.
  eo_iterator = tables->if_object_collection~get_iterator( ).
  endmethod.


method GET_TABLES_SIZE.
  ep_size = tables->if_object_collection~size( ).
  endmethod.


method GET_TITLE.
  DATA lv_value TYPE string.
  IF ip_escaped EQ abap_true.
    lv_value = me->title.
    ep_title = zcl_excel_common=>escape_string( lv_value ).
  ELSE.
    ep_title = me->title.
  ENDIF.
  endmethod.


METHOD get_value_type.
  DATA:  lo_addit    TYPE REF TO cl_abap_elemdescr,
         ls_dfies    TYPE dfies,
         l_function  TYPE funcname,
         l_value(50) TYPE c.

  ep_value = ip_value.
  ep_value_type = cl_abap_typedescr=>typekind_string. " Thats our default if something goes wrong.

  TRY.
      lo_addit            ?= cl_abap_typedescr=>describe_by_data( ip_value ).
    CATCH cx_sy_move_cast_error.
      CLEAR lo_addit.
  ENDTRY.
  IF lo_addit IS BOUND.
    lo_addit->get_ddic_field( RECEIVING  p_flddescr   = ls_dfies
                              EXCEPTIONS not_found    = 1
                                         no_ddic_type = 2
                                         OTHERS       = 3 ) .
    IF sy-subrc = 0.
      ep_value_type = ls_dfies-inttype.

      IF ls_dfies-convexit IS NOT INITIAL.
* We need to convert with output conversion function
        CONCATENATE 'CONVERSION_EXIT_' ls_dfies-convexit '_OUTPUT' INTO l_function.
        SELECT SINGLE funcname INTO l_function
              FROM tfdir
              WHERE funcname = l_function.
        IF sy-subrc = 0.
          CALL FUNCTION l_function
            EXPORTING
              input  = ip_value
            IMPORTING
*             LONG_TEXT  =
              output = l_value
*             SHORT_TEXT =
            EXCEPTIONS
              OTHERS = 1.
          IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
          ELSE.
            TRY.
                ep_value = l_value.
              CATCH cx_root.
                ep_value = ip_value.
            ENDTRY.
          ENDIF.
        ENDIF.
      ENDIF.
    ELSE.
      ep_value_type = lo_addit->get_data_type_kind( ip_value ).
    ENDIF.
  ENDIF.

ENDMETHOD.


METHOD is_cell_merged.

  DATA: lv_column TYPE i.

  FIELD-SYMBOLS: <ls_merged_cell> LIKE LINE OF me->mt_merged_cells.

  lv_column = zcl_excel_common=>convert_column2int( ip_column ).

  rp_is_merged = abap_false.                                        " Assume not in merged area

  LOOP AT me->mt_merged_cells ASSIGNING <ls_merged_cell>.

    IF    <ls_merged_cell>-col_from <= lv_column
      AND <ls_merged_cell>-col_to   >= lv_column
      AND <ls_merged_cell>-row_from <= ip_row
      AND <ls_merged_cell>-row_to   >= ip_row.
      rp_is_merged = abap_true.                                     " until we are proven different
      RETURN.
    ENDIF.

  ENDLOOP.

ENDMETHOD.


method PRINT_TITLE_SET_RANGE.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                            2012-12-02
*--------------------------------------------------------------------*


  DATA:     lo_range_iterator               TYPE REF TO cl_object_collection_iterator,
            lo_range                        TYPE REF TO zcl_excel_range,
            lv_repeat_range_sheetname       TYPE string,
            lv_repeat_range_col             TYPE string,
            lv_row_char_from                TYPE char10,
            lv_row_char_to                  TYPE char10,
            lv_repeat_range_row             TYPE string,
            lv_repeat_range                 TYPE string.


*--------------------------------------------------------------------*
* Get range that represents printarea
* if non-existant, create it
*--------------------------------------------------------------------*
  lo_range_iterator = me->get_ranges_iterator( ).
  WHILE lo_range_iterator->has_next( ) = abap_true.

    lo_range ?= lo_range_iterator->get_next( ).
    IF lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
      EXIT.  " Found it
    ENDIF.
    CLEAR lo_range.

  ENDWHILE.


  IF me->print_title_col_from IS INITIAL AND
     me->print_title_row_from IS INITIAL.
*--------------------------------------------------------------------*
* No print titles are present,
*--------------------------------------------------------------------*
    IF lo_range IS BOUND.
      me->ranges->remove( lo_range ).
    ENDIF.
  ELSE.
*--------------------------------------------------------------------*
* Print titles are present,
*--------------------------------------------------------------------*
    IF lo_range IS NOT BOUND.
      lo_range =  me->add_new_range( ).
      lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
    ENDIF.

    lv_repeat_range_sheetname = me->get_title( ).
    lv_repeat_range_sheetname = zcl_excel_common=>escape_string( lv_repeat_range_sheetname ).

*--------------------------------------------------------------------*
* Repeat-columns
*--------------------------------------------------------------------*
    IF me->print_title_col_from IS NOT INITIAL.
      CONCATENATE lv_repeat_range_sheetname
                  '!$' me->print_title_col_from
                  ':$' me->print_title_col_to
          INTO lv_repeat_range_col.
    ENDIF.

*--------------------------------------------------------------------*
* Repeat-rows
*--------------------------------------------------------------------*
    IF me->print_title_row_from IS NOT INITIAL.
      lv_row_char_from = me->print_title_row_from.
      lv_row_char_to   = me->print_title_row_to.
      CONCATENATE '!$' lv_row_char_from
                  ':$' lv_row_char_to
          INTO lv_repeat_range_row.
      CONDENSE lv_repeat_range_row NO-GAPS.
      CONCATENATE lv_repeat_range_sheetname
                  lv_repeat_range_row
          INTO lv_repeat_range_row.
    ENDIF.

*--------------------------------------------------------------------*
* Concatenate repeat-rows and columns
*--------------------------------------------------------------------*
    IF lv_repeat_range_col IS INITIAL.
      lv_repeat_range = lv_repeat_range_row.
    ELSEIF lv_repeat_range_row IS INITIAL.
      lv_repeat_range = lv_repeat_range_col.
    ELSE.
      CONCATENATE lv_repeat_range_col lv_repeat_range_row
          INTO lv_repeat_range SEPARATED BY ','.
    ENDIF.


    lo_range->set_range_value( lv_repeat_range ).
  ENDIF.



  endmethod.


method SET_CELL.

  DATA: ls_sheet_content          TYPE zexcel_s_cell_data,
        lv_value                  TYPE zexcel_cell_value,
        lv_data_type              TYPE zexcel_cell_data_type,
        lv_value_type             TYPE abap_typekind,
        lv_style_guid             TYPE zexcel_cell_style,
        lo_addit                  TYPE REF TO cl_abap_elemdescr,
        lo_value                  TYPE REF TO data,
        lo_value_new              TYPE REF TO data.

  DATA: lv_provided_style_guid    TYPE zexcel_cell_style,
        lv_first_cell             TYPE abap_bool,
        lv_column                 TYPE zexcel_cell_column_alpha,
        lv_column_int             TYPE zexcel_cell_column,
        lv_row                    TYPE zexcel_cell_row,
        lv_column_start           TYPE zexcel_cell_column_alpha,
        lv_column_end             TYPE zexcel_cell_column_alpha,
        lv_column_int_start       TYPE zexcel_cell_column,
        lv_column_int_end         TYPE zexcel_cell_column,
        lv_row_start              TYPE zexcel_cell_row,
        lv_row_end                TYPE zexcel_cell_row.

  FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data,
                 <fs_numeric>       TYPE numeric,
                 <fs_date>          TYPE d,
                 <fs_time>          TYPE t,
                 <fs_value>         TYPE simple.

  IF ip_value  IS NOT SUPPLIED AND ip_formula IS NOT SUPPLIED.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Please provide the value or formula'.
  ENDIF.

  " #524 accept dynamic IP_STYLE paramter
  lv_provided_style_guid = zcl_excel_common=>get_guid_of( ip_style ).

  " #524 accept Excel-style cell reference parameter IP_CELL
  determine_range(
    EXPORTING ip_cell           = ip_cell
              ip_column         = ip_column
              ip_row            = ip_row
              ip_column_to      = ip_column_to
              ip_row_to         = ip_row_to
    IMPORTING ev_column_i_start = lv_column_int_start
              ev_column_a_start = lv_column_start
              ev_row_start      = lv_row_start
              ev_column_i_end   = lv_column_int_end
              ev_column_a_end   = lv_column_end
              ev_row_end        = lv_row_end ).

  lv_column_int = lv_column_int_start.
  lv_row        = lv_row_start.
  lv_first_cell = ABAP_TRUE.

  " Loop over each cell in range
  " lv_column, lv_column_int, lv_row contain current loop cell address
  " only assign value/formula to first (top left) cell
  " set style to all cells in range
  DO.
    lv_column = zcl_excel_common=>convert_column2alpha( lv_column_int ).

*     Begin of change issue #152 - don't touch exisiting style if only value is passed
*      lv_style_guid = ip_style.
    UNASSIGN <fs_sheet_content>.
    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH TABLE KEY cell_row    = lv_row      " Changed to access via table key , Stefan Schmöcker, 2013-08-03
                                                                         cell_column = lv_column_int.
    IF sy-subrc = 0.
      IF lv_provided_style_guid IS INITIAL.
        " If no style is provided as method-parameter and cell is found use cell's current style
        lv_style_guid = <fs_sheet_content>-cell_style.
      ELSE.
        " Style provided as method-parameter --> use this
        lv_style_guid = lv_provided_style_guid.
      ENDIF.
    ELSE.
      " No cell found --> use supplied style even if empty
      lv_style_guid = lv_provided_style_guid.
    ENDIF.
*   End of change issue #152 - don't touch exisiting style if only value is passed

    IF ip_value IS SUPPLIED AND lv_first_cell = ABAP_TRUE.
      "if data type is passed just write the value. Otherwise map abap type to excel and perform conversion
      "IP_DATA_TYPE is passed by excel reader so source types are preserved
*  First we get reference into local var.
      CREATE DATA lo_value LIKE ip_value.
      ASSIGN lo_value->* TO <fs_value>.
      <fs_value> = ip_value.
      IF ip_data_type IS SUPPLIED.
        IF ip_abap_type IS NOT SUPPLIED.
          get_value_type( EXPORTING ip_value      = ip_value
                          IMPORTING ep_value      = <fs_value> ) .
        ENDIF.
        lv_value = <fs_value>.
        lv_data_type = ip_data_type.
      ELSE.
        IF ip_abap_type IS SUPPLIED.
          lv_value_type = ip_abap_type.
        ELSE.
          get_value_type( EXPORTING ip_value      = ip_value
                          IMPORTING ep_value      = <fs_value>
                                    ep_value_type = lv_value_type ).
        ENDIF.
        CASE lv_value_type.
          WHEN cl_abap_typedescr=>typekind_int OR cl_abap_typedescr=>typekind_int1 OR cl_abap_typedescr=>typekind_int2.
            lo_addit = cl_abap_elemdescr=>get_i( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_value>.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_float OR cl_abap_typedescr=>typekind_packed.
            lo_addit = cl_abap_elemdescr=>get_f( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_value>.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_char OR cl_abap_typedescr=>typekind_string OR cl_abap_typedescr=>typekind_num OR
               cl_abap_typedescr=>typekind_hex.
            lv_value = <fs_value>.
            lv_data_type = 's'.

          WHEN cl_abap_typedescr=>typekind_date.
            lo_addit = cl_abap_elemdescr=>get_d( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_date>.
            IF sy-subrc = 0.
              <fs_date> = <fs_value>.
              lv_value = zcl_excel_common=>date_to_excel_string( ip_value = <fs_date> ) .
            ENDIF.
*   Begin of change issue #152 - don't touch exisiting style if only value is passed
*   Moved to end of routine - apply date-format even if other styleinformation is passed
*            IF ip_style IS NOT SUPPLIED. "get default date format in case parameter is initial
*              lo_style = excel->add_new_style( ).
*              lo_style->number_format->format_code = get_default_excel_date_format( ).
*              lv_style_guid = lo_style->get_guid( ).
*            ENDIF.
*   End of change issue #152 - don't touch exisiting style if only value is passed

          WHEN cl_abap_typedescr=>typekind_time.
            lo_addit = cl_abap_elemdescr=>get_t( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_time>.
            IF sy-subrc = 0.
              <fs_time> = <fs_value>.
              lv_value = zcl_excel_common=>time_to_excel_string( ip_value = <fs_time> ).
            ENDIF.
*   Begin of change issue #152 - don't touch exisiting style if only value is passed
*   Moved to end of routine - apply time-format even if other styleinformation is passed
*            IF ip_style IS NOT SUPPLIED. "get default time format for user in case parameter is initial
*              lo_style = excel->add_new_style( ).
*              lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_time6.
*              lv_style_guid = lo_style->get_guid( ).
*            ENDIF.
*   End of change issue #152 - don't touch exisiting style if only value is passed

          WHEN OTHERS.
            RAISE EXCEPTION TYPE zcx_excel
              EXPORTING
                error = 'Invalid data type of input value'.
        ENDCASE.
      ENDIF.

    ENDIF.

    IF ip_hyperlink IS BOUND AND lv_first_cell = ABAP_TRUE.
      ip_hyperlink->set_cell_reference( ip_column = lv_column
                                        ip_row = lv_row ).
      me->hyperlinks->add( ip_hyperlink ).
    ENDIF.

*   Begin of change issue #152 - don't touch exisiting style if only value is passed
*   Read table moved up, so that current style may be evaluated
*    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

*    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH KEY cell_row    = ip_row
*                                                                   cell_column = lv_column.
*
*    IF sy-subrc EQ 0.
*    IF <fs_sheet_content> IS ASSIGNED.
* End of change issue #152 - don't touch exisiting style if only value is passed

    IF <fs_sheet_content> IS NOT ASSIGNED.
      CLEAR: ls_sheet_content.
      ls_sheet_content-cell_row    = lv_row.
      ls_sheet_content-cell_column = lv_column_int.
      ls_sheet_content-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( ip_column = lv_column ip_row = lv_row ).
      INSERT ls_sheet_content INTO TABLE sheet_content ASSIGNING <fs_sheet_content>. "ins #152 - Now <fs_sheet_content> always holds the data
    ENDIF.

    IF lv_first_cell = ABAP_TRUE.
      <fs_sheet_content>-cell_value   = lv_value.
      <fs_sheet_content>-cell_formula = ip_formula.
      IF ip_formula IS INITIAL AND lv_value IS NOT INITIAL AND it_rtf[] IS NOT INITIAL.
        check_rtf( ip_value = lv_value
                   it_rtf   = it_rtf[] ).
        CREATE DATA <fs_sheet_content>-RTF_TAB.
        <fs_sheet_content>-RTF_TAB->* = it_rtf[].
      ENDIF.
    ELSE.
      CLEAR <fs_sheet_content>-cell_value.
      CLEAR <fs_sheet_content>-cell_formula.
      CLEAR <fs_sheet_content>-rtf_tab.
    ENDIF.
    <fs_sheet_content>-data_type    = lv_data_type.
    <fs_sheet_content>-cell_style   = lv_style_guid.

*    " Deduplicating code by inserting new table line and assiging a field symbol to it
*    ELSE.
*      CLEAR: ls_sheet_content.
*      ls_sheet_content-cell_row     = lv_row.
*      ls_sheet_content-cell_column  = lv_column_int.
*      IF lv_first_cell = ABAP_TRUE.
*        ls_sheet_content-cell_value   = lv_value.
*        ls_sheet_content-cell_formula = ip_formula.
*        IF ip_formula IS INITIAL AND lv_value IS NOT INITIAL AND it_rtf[] IS NOT INITIAL.
*          CREATE DATA ls_sheet_content-RTF_TAB.
*          ls_sheet_content-RTF_TAB->* = it_rtf[].
*        ENDIF.
*      ENDIF.
*      ls_sheet_content-data_type   = lv_data_type.
*      ls_sheet_content-cell_style  = lv_style_guid.
*      ls_sheet_content-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( ip_column = lv_column ip_row = lv_row ).
*      INSERT ls_sheet_content INTO TABLE sheet_content ASSIGNING <fs_sheet_content>. "ins #152 - Now <fs_sheet_content> always holds the data
**      APPEND ls_sheet_content TO sheet_content.
**      SORT sheet_content BY cell_row cell_column.
*      " me->update_dimension_range( ).
*
*    ENDIF.

    IF lv_first_cell = ABAP_TRUE.
*     Begin of change issue #152 - don't touch exisiting style if only value is passed
*     For Date- or Timefields change the formatcode if nothing is set yet
*     Enhancement option:  Check if existing formatcode is a date/ or timeformat
*                          If not, use default
      DATA: lo_format_code_datetime TYPE zexcel_number_format.
      DATA: stylemapping    TYPE zexcel_s_stylemapping.
      CASE lv_value_type.
        WHEN cl_abap_typedescr=>typekind_date.
          TRY.
              stylemapping = me->excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
            CATCH zcx_excel .
          ENDTRY.
          IF stylemapping-complete_stylex-number_format-format_code IS INITIAL OR
             stylemapping-complete_style-number_format-format_code IS INITIAL.
            lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_std.
          ELSE.
            lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
          ENDIF.
          me->change_cell_style( ip_column                      = lv_column
                                 ip_row                         = lv_row
                                 ip_number_format_format_code   = lo_format_code_datetime ).

        WHEN cl_abap_typedescr=>typekind_time.
          TRY.
              stylemapping = me->excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
            CATCH zcx_excel .
          ENDTRY.
          IF stylemapping-complete_stylex-number_format-format_code IS INITIAL OR
             stylemapping-complete_style-number_format-format_code IS INITIAL.
            lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_time6.
          ELSE.
            lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
          ENDIF.
          me->change_cell_style( ip_column                      = lv_column
                                 ip_row                         = lv_row
                                 ip_number_format_format_code   = lo_format_code_datetime ).

      ENDCASE.
*     End of change issue #152 - don't touch exisiting style if only value is passed

*     Fix issue #162
      lv_value = ip_value.
      IF lv_value CS cl_abap_char_utilities=>cr_lf.
        me->change_cell_style( ip_column               = lv_column
                               ip_row                  = lv_row
                               ip_alignment_wraptext   = abap_true ).
      ENDIF.
*     End of Fix issue #162

    ENDIF.

    " move to next column
    lv_column_int = lv_column_int + 1.
    " move to new row if column is out of range
    IF lv_column_int > lv_column_int_end.
      lv_column_int = lv_column_int_start.
      lv_row = lv_row + 1.
    ENDIF.
    " exit if row is out of range
    IF lv_row > lv_row_end.
      EXIT. " DO
    ENDIF.
    CLEAR: lv_first_cell.
  ENDDO.

  IF NOT ( lv_column_end = lv_column_start AND lv_row_end = lv_row_start ).
    set_merge(
        ip_column_start = lv_column_start
        ip_column_end   = lv_column_end
        ip_row          = lv_row_start
        ip_row_to       = lv_row_end ).
  ENDIF.

endmethod.


method SET_CELL_FORMULA.
  DATA:
              lv_column                       TYPE zexcel_cell_column,
              lv_row                          TYPE zexcel_cell_row,
              ls_sheet_content                LIKE LINE OF me->sheet_content.

  FIELD-SYMBOLS:
              <sheet_content>                 LIKE LINE OF me->sheet_content.

  determine_range(
    EXPORTING ip_cell           = ip_cell
              ip_column         = ip_column
              ip_row            = ip_row
    IMPORTING ev_column_i_start = lv_column
              ev_row_start      = lv_row ).

*--------------------------------------------------------------------*
* Get cell to set formula into
*--------------------------------------------------------------------*
  READ TABLE me->sheet_content ASSIGNING <sheet_content> WITH TABLE KEY cell_row    = lv_row
                                                                        cell_column = lv_column.
  IF sy-subrc <> 0.                                                                           " Create new entry in sheet_content if necessary
    CHECK ip_formula IS INITIAL.                                                              " no need to create new entry in sheet_content when no formula is passed
    ls_sheet_content-cell_row    = lv_row.
    ls_sheet_content-cell_column = lv_column.
    INSERT ls_sheet_content INTO TABLE me->sheet_content ASSIGNING <sheet_content>.
  ENDIF.

*--------------------------------------------------------------------*
* Fieldsymbol now holds the relevant cell
*--------------------------------------------------------------------*
  <sheet_content>-cell_formula = ip_formula.


  endmethod.


METHOD set_cell_style.

  DATA: lv_column                 TYPE zexcel_cell_column,
        lv_row                    TYPE zexcel_cell_row,
        lv_style_guid             TYPE zexcel_cell_style.

  FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data.

  determine_range(
    EXPORTING ip_cell           = ip_cell
              ip_column         = ip_column
              ip_row            = ip_row
    IMPORTING ev_column_i_start = lv_column
              ev_row_start      = lv_row ).

  lv_style_guid = zcl_excel_common=>get_guid_of( ip_style ).

  READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH KEY cell_row    = lv_row
                                                                 cell_column = lv_column.

  IF sy-subrc EQ 0.
    <fs_sheet_content>-cell_style   = lv_style_guid.
  ELSE.
    set_cell( ip_column = lv_column ip_row = lv_row ip_value = '' ip_style = lv_style_guid ).
  ENDIF.

ENDMETHOD.


method SET_COLUMN_WIDTH.
  DATA: lo_column  TYPE REF TO zcl_excel_column.
  DATA: width             TYPE float.

  lo_column = me->get_column( ip_column ).

* if a fix size is supplied use this
  IF ip_width_fix IS SUPPLIED.
    TRY.
        width = ip_width_fix.
        IF width <= 0.
          RAISE EXCEPTION TYPE zcx_excel
            EXPORTING
              error = 'Please supply a positive number as column-width'.
        ENDIF.
        lo_column->set_width( width ).
        EXIT.
      CATCH cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = 'Unable to interpret supplied input as number'.
    ENDTRY.
  ENDIF.

* If we get down to here, we have to use whatever is found in autosize.
  lo_column->set_auto_size( ip_width_autosize ).


  endmethod.


method SET_DEFAULT_EXCEL_DATE_FORMAT.

  IF ip_default_excel_date_format IS INITIAL.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Default date format cannot be blank'.
  ENDIF.

  default_excel_date_format = ip_default_excel_date_format.
  endmethod.


METHOD set_merge.

  DATA: ls_merge        TYPE mty_merge,
        lv_errormessage TYPE string.

  DATA: lv_column_start	TYPE zexcel_cell_column_alpha,
        lv_column_end	  TYPE zexcel_cell_column_alpha,
        lv_row         	TYPE zexcel_cell_row,
        lv_row_to	      TYPE zexcel_cell_row.

  " #524 accept Excel-style cell reference parameter IP_CELL
  determine_range(
    EXPORTING ip_cell           = ip_cell
              ip_column         = ip_column_start
              ip_row            = ip_row
              ip_column_to      = ip_column_end
              ip_row_to         = ip_row_to
    IMPORTING ev_column_a_start = lv_column_start
              ev_row_start      = lv_row
              ev_column_a_end   = lv_column_end
              ev_row_end        = lv_row_to ).

*--------------------------------------------------------------------*
* Build new range area to insert into range table
*--------------------------------------------------------------------*
  ls_merge-row_from = lv_row.
  IF lv_row IS NOT INITIAL AND lv_row_to IS INITIAL.
    ls_merge-row_to   = ls_merge-row_from.
  ELSE.
    ls_merge-row_to   = lv_row_to.
  ENDIF.
  IF ls_merge-row_from > ls_merge-row_to.
    lv_errormessage = 'Merge: First row larger then last row'(405).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

  ls_merge-col_from = zcl_excel_common=>convert_column2int( lv_column_start ).
  IF lv_column_start IS NOT INITIAL AND lv_column_end IS INITIAL.
    ls_merge-col_to   = ls_merge-col_from.
  ELSE.
    ls_merge-col_to   = zcl_excel_common=>convert_column2int( lv_column_end ).
  ENDIF.
  IF ls_merge-col_from > ls_merge-col_to.
    lv_errormessage = 'Merge: First column larger then last column'(406).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

*--------------------------------------------------------------------*
* Check merge not overlapping with existing merges
*--------------------------------------------------------------------*
  LOOP AT me->mt_merged_cells TRANSPORTING NO FIELDS WHERE NOT (    row_from > ls_merge-row_to
                                                                 OR row_to   < ls_merge-row_from
                                                                 OR col_from > ls_merge-col_to
                                                                 OR col_to   < ls_merge-col_from ).
    lv_errormessage = 'Overlapping merges'(404).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.

  ENDLOOP.

*--------------------------------------------------------------------*
* Everything seems ok --> add to merge table
*--------------------------------------------------------------------*
  INSERT ls_merge INTO TABLE me->mt_merged_cells.

ENDMETHOD.


method SET_PRINT_GRIDLINES.
  me->print_gridlines = i_print_gridlines.
  endmethod.


METHOD set_row_height.
  DATA: lo_row  TYPE REF TO zcl_excel_row.
  DATA: height  TYPE float.

  lo_row = me->get_row( ip_row ).

* if a fix size is supplied use this
  TRY.
      height = ip_height_fix.
      IF height <= 0.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = 'Please supply a positive number as row-height'.
      ENDIF.
      lo_row->set_row_height( height ).
      EXIT.
    CATCH cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Unable to interpret supplied input as number'.
  ENDTRY.

ENDMETHOD.


METHOD set_row_outline.

  DATA: ls_row_outline LIKE LINE OF me->mt_row_outlines.
  FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF me->mt_row_outlines.

  READ TABLE me->mt_row_outlines ASSIGNING <ls_row_outline> WITH TABLE KEY row_from = iv_row_from
                                                                           row_to   = iv_row_to.
  IF sy-subrc <> 0.
    IF iv_row_from <= 0.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'First row of outline must be a positive number'.
    ENDIF.
    IF iv_row_to < iv_row_from.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Last row of outline may not be less than first line of outline'.
    ENDIF.
    ls_row_outline-row_from = iv_row_from.
    ls_row_outline-row_to   = iv_row_to.
    INSERT ls_row_outline INTO TABLE me->mt_row_outlines ASSIGNING <ls_row_outline>.
  ENDIF.

  CASE iv_collapsed.

    WHEN abap_true
      OR abap_false.
      <ls_row_outline>-collapsed = iv_collapsed.

    WHEN OTHERS.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Unknown collapse state'.

  ENDCASE.
ENDMETHOD.


method SET_SHOW_GRIDLINES.
  me->show_gridlines = i_show_gridlines.
  endmethod.


method SET_SHOW_ROWCOLHEADERS.
  me->show_rowcolheaders = i_show_rowcolheaders.
  endmethod.


method SET_TABCOLOR.
  me->tabcolor = iv_tabcolor.
  endmethod.


method SET_TABLE.

  DATA: lo_tabdescr     TYPE REF TO cl_abap_structdescr,
        lr_data         TYPE REF TO data,
        ls_header       TYPE x030l,
        lt_dfies        TYPE ddfields,
        lv_row_int      TYPE zexcel_cell_row,
        lv_column_int   TYPE zexcel_cell_column,
        lv_column_alpha TYPE zexcel_cell_column_alpha,
        lv_cell_value   TYPE zexcel_cell_value.


  FIELD-SYMBOLS: <fs_table_line>  TYPE ANY,
                 <fs_fldval>      TYPE ANY,
                 <fs_dfies>       TYPE dfies.

  lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
  lv_row_int    = ip_top_left_row.

  CREATE DATA lr_data LIKE LINE OF ip_table.

  lo_tabdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

  ls_header = lo_tabdescr->get_ddic_header( ).

  lt_dfies = lo_tabdescr->get_ddic_field_list( ).

* It is better to loop column by column
  LOOP AT lt_dfies ASSIGNING <fs_dfies>.
    lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

    IF ip_no_header = abap_false.
      " First of all write column header
      lv_cell_value = <fs_dfies>-scrtext_m.
      me->set_cell( ip_column = lv_column_alpha
                    ip_row    = lv_row_int
                    ip_value  = lv_cell_value
                    ip_style  = ip_hdr_style ).
      IF ip_transpose = abap_true.
        ADD 1 TO lv_column_int.
      ELSE.
        ADD 1 TO lv_row_int.
      ENDIF.
    ENDIF.

    LOOP AT ip_table ASSIGNING <fs_table_line>.
      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).
      ASSIGN COMPONENT <fs_dfies>-fieldname OF STRUCTURE <fs_table_line> TO <fs_fldval>.
      MOVE <fs_fldval> TO lv_cell_value.
      me->set_cell( ip_column = lv_column_alpha
                    ip_row    = lv_row_int
                    ip_value  = <fs_fldval>   "lv_cell_value
                    ip_style  = ip_body_style ).
      IF ip_transpose = abap_true.
        ADD 1 TO lv_column_int.
      ELSE.
        ADD 1 TO lv_row_int.
      ENDIF.
    ENDLOOP.
    IF ip_transpose = abap_true.
      lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
      ADD 1 TO lv_row_int.
    ELSE.
      lv_row_int = ip_top_left_row.
      ADD 1 TO lv_column_int.
    ENDIF.
  ENDLOOP.

  endmethod.


method SET_TITLE.
*--------------------------------------------------------------------*
* ToDos:
*        2do §1  The current coding for replacing a named ranges name
*                after renaming a sheet should be checked if it is
*                really working if sheetname should be escaped
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (wip )              2012-12-08
*              - ...
* changes: aligning code
*          message made to support multilinguality
*--------------------------------------------------------------------*
* issue#243 - ' is not allowed as first character in sheet title
*              - Stefan Schmoecker,                          2012-12-02
* changes: added additional check for ' as first character
*--------------------------------------------------------------------*
  DATA:       lo_worksheets_iterator          TYPE REF TO cl_object_collection_iterator,
              lo_worksheet                    TYPE REF TO zcl_excel_worksheet,
              errormessage                    TYPE string,
              lv_rangesheetname_old           TYPE string,
              lv_rangesheetname_new           TYPE string,
              lo_ranges_iterator              TYPE REF TO cl_object_collection_iterator,
              lo_range                        TYPE REF TO zcl_excel_range,
              lv_range_value                  TYPE zexcel_range_value,
              lv_errormessage                 TYPE string.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Check whether title consists only of allowed characters
* Illegal characters are: /	\	[	]	*	?	:  --> http://msdn.microsoft.com/en-us/library/ff837411.aspx
* Illegal characters not in documentation:   ' as first character
*--------------------------------------------------------------------*
  IF ip_title CA '/\[]*?:'.
    lv_errormessage = 'Found illegal character in sheetname. List of forbidden characters: /\[]*?:'(402).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

  IF ip_title IS NOT INITIAL AND ip_title(1) = `'`.
    lv_errormessage = 'Sheetname may not start with &'(403).   " & used instead of ' to allow fallbacklanguage
    REPLACE '&' IN lv_errormessage WITH `'`.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.


*--------------------------------------------------------------------*
* Check whether title is unique in workbook
*--------------------------------------------------------------------*
  lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
  WHILE lo_worksheets_iterator->has_next( ) = 'X'.

    lo_worksheet ?= lo_worksheets_iterator->get_next( ).
    CHECK me->guid <> lo_worksheet->get_guid( ).  " Don't check against itself
    IF ip_title = lo_worksheet->get_title( ).  " Not unique --> raise exception
      errormessage = 'Duplicate sheetname &'.
      REPLACE '&' IN errormessage WITH ip_title.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = errormessage.
    ENDIF.

  ENDWHILE.

*--------------------------------------------------------------------*
* Remember old sheetname and rename sheet to desired name
*--------------------------------------------------------------------*
  CONCATENATE me->title '!' INTO lv_rangesheetname_old.
  me->title = ip_title.

*--------------------------------------------------------------------*
* After changing this worksheet's title we have to adjust
* all ranges that are referring to this worksheet.
*--------------------------------------------------------------------*
* 2do §1  -  Check if the following quickfix is solid
*           I fear it isn't - but this implementation is better then
*           nothing at all since it handles a supposed majority of cases
*--------------------------------------------------------------------*
  CONCATENATE me->title '!' INTO lv_rangesheetname_new.

  lo_ranges_iterator = me->excel->get_ranges_iterator( ).
  WHILE lo_ranges_iterator->has_next( ) = 'X'.

    lo_range ?= lo_ranges_iterator->get_next( ).
    lv_range_value = lo_range->get_value( ).
    REPLACE ALL OCCURRENCES OF lv_rangesheetname_old IN lv_range_value WITH lv_rangesheetname_new.
    IF sy-subrc = 0.
      lo_range->set_range_value( lv_range_value ).
    ENDIF.

  ENDWHILE.


  endmethod.


METHOD update_dimension_range.

  DATA: ls_sheet_content TYPE zexcel_s_cell_data,
        lt_sheet_content TYPE zexcel_t_cell_data_unsorted,
        lv_row_alpha     TYPE string,
        lv_column_alpha  TYPE zexcel_cell_column_alpha.

  CHECK sheet_content IS NOT INITIAL.

* update dimension range
  lt_sheet_content = sheet_content.
  "upper left corner
  SORT lt_sheet_content BY cell_row.
  READ TABLE lt_sheet_content INDEX 1 INTO ls_sheet_content.
  upper_cell-cell_row = ls_sheet_content-cell_row.
  SORT lt_sheet_content BY cell_column.
  READ TABLE lt_sheet_content INDEX 1 INTO ls_sheet_content.
  upper_cell-cell_column = ls_sheet_content-cell_column.

  lv_row_alpha = upper_cell-cell_row.
  lv_column_alpha = zcl_excel_common=>convert_column2alpha( upper_cell-cell_column ).
  SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
  SHIFT lv_row_alpha LEFT DELETING LEADING space.
  CONCATENATE lv_column_alpha lv_row_alpha INTO upper_cell-cell_coords.

  "bottom right corner
  SORT lt_sheet_content BY cell_row DESCENDING.
  READ TABLE lt_sheet_content INDEX 1 INTO ls_sheet_content.
  lower_cell-cell_row = ls_sheet_content-cell_row.
  SORT lt_sheet_content BY cell_column DESCENDING.
  READ TABLE lt_sheet_content INDEX 1 INTO ls_sheet_content.
  lower_cell-cell_column = ls_sheet_content-cell_column.

  lv_row_alpha = lower_cell-cell_row.
  lv_column_alpha = zcl_excel_common=>convert_column2alpha( lower_cell-cell_column ).
  SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
  SHIFT lv_row_alpha LEFT DELETING LEADING space.
  CONCATENATE lv_column_alpha lv_row_alpha INTO lower_cell-cell_coords.

ENDMETHOD.


method ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_COLUMNS.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
  CLEAR:  me->print_title_col_from,
          me->print_title_col_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
  me->print_title_set_range( ).


  endmethod.


method ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_ROWS.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
  CLEAR:  me->print_title_row_from,
          me->print_title_row_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
  me->print_title_set_range( ).


  endmethod.


method ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_COLUMNS.
  ev_columns_from = me->print_title_col_from.
  ev_columns_to   = me->print_title_col_to.
  endmethod.


method ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_ROWS.
  ev_rows_from = me->print_title_row_from.
  ev_rows_to   = me->print_title_row_to.
  endmethod.


method ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_COLUMNS.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

  DATA:     lv_col_from_int                 TYPE i,
            lv_col_to_int                   TYPE i,
            lv_errormessage                 TYPE string.


  lv_col_from_int = zcl_excel_common=>convert_column2int( iv_columns_from ).
  lv_col_to_int   = zcl_excel_common=>convert_column2int( iv_columns_to ).

*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
  IF lv_col_from_int < 1.
    lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

  IF  lv_col_from_int > lv_col_to_int.
    lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
  me->print_title_col_from = iv_columns_from.
  me->print_title_col_to   = iv_columns_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
  me->print_title_set_range( ).

  endmethod.


method ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_ROWS.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

  DATA:     lv_errormessage                 TYPE string.


*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
  IF iv_rows_from < 1.
    lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

  IF  iv_rows_from > iv_rows_to.
    lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
  me->print_title_row_from = iv_rows_from.
  me->print_title_row_to   = iv_rows_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
  me->print_title_set_range( ).


  endmethod.


method ZIF_EXCEL_SHEET_PROPERTIES~GET_STYLE.
  IF zif_excel_sheet_properties~style IS NOT INITIAL.
    ep_style = zif_excel_sheet_properties~style.
  ELSE.
    ep_style = me->excel->get_default_style( ).
  ENDIF.
  endmethod.


method ZIF_EXCEL_SHEET_PROPERTIES~INITIALIZE.

  zif_excel_sheet_properties~show_zeros   = zif_excel_sheet_properties=>c_showzero.
  zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_on.
  zif_excel_sheet_properties~summaryright = zif_excel_sheet_properties=>c_right_on.

* inizialize zoomscale values
  ZIF_EXCEL_SHEET_PROPERTIES~zoomscale = 100.
  ZIF_EXCEL_SHEET_PROPERTIES~zoomscale_normal = 100.
  ZIF_EXCEL_SHEET_PROPERTIES~zoomscale_pagelayoutview = 100 .
  ZIF_EXCEL_SHEET_PROPERTIES~zoomscale_sheetlayoutview = 100 .
  endmethod.


method ZIF_EXCEL_SHEET_PROPERTIES~SET_STYLE.
  zif_excel_sheet_properties~style = ip_style.
  endmethod.


method ZIF_EXCEL_SHEET_PROTECTION~INITIALIZE.

  me->zif_excel_sheet_protection~protected = zif_excel_sheet_protection=>c_unprotected.
  CLEAR me->zif_excel_sheet_protection~password.
  me->zif_excel_sheet_protection~auto_filter            = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~delete_columns         = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~delete_rows            = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~format_cells           = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~format_columns         = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~format_rows            = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~insert_columns         = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~insert_hyperlinks      = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~insert_rows            = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~objects                = zif_excel_sheet_protection=>c_noactive.
*  me->zif_excel_sheet_protection~password               = zif_excel_sheet_protection=>c_noactive. "issue #68
  me->zif_excel_sheet_protection~pivot_tables           = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~protected              = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~scenarios              = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~select_locked_cells    = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~select_unlocked_cells  = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~sheet                  = zif_excel_sheet_protection=>c_noactive.
  me->zif_excel_sheet_protection~sort                   = zif_excel_sheet_protection=>c_noactive.

  endmethod.


method ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME.
  me->zif_excel_sheet_vba_project~codename = ip_codename.
  endmethod.


method ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME_PR.
  me->zif_excel_sheet_vba_project~codename_pr = ip_codename_pr.
  endmethod.
ENDCLASS.
