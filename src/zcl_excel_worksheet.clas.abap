class ZCL_EXCEL_WORKSHEET definition
  public
  create public .

  public section.
*"* public components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
    type-pools ABAP .

    interfaces ZIF_EXCEL_SHEET_PRINTSETTINGS .
    interfaces ZIF_EXCEL_SHEET_PROPERTIES .
    interfaces ZIF_EXCEL_SHEET_PROTECTION .
    interfaces ZIF_EXCEL_SHEET_VBA_PROJECT .

    types:
      begin of  MTY_S_OUTLINE_ROW,
        ROW_FROM  type I,
        ROW_TO    type I,
        COLLAPSED type ABAP_BOOL,
      end of MTY_S_OUTLINE_ROW .
    types:
      MTY_TS_OUTLINES_ROW type sorted table of MTY_S_OUTLINE_ROW with unique key ROW_FROM ROW_TO .

    constants C_BREAK_COLUMN type ZEXCEL_BREAK value 2.     "#EC NOTEXT
    constants C_BREAK_NONE type ZEXCEL_BREAK value 0.       "#EC NOTEXT
    constants C_BREAK_ROW type ZEXCEL_BREAK value 1.        "#EC NOTEXT
    data EXCEL type ref to ZCL_EXCEL read-only .
  data PRINT_GRIDLINES type ZEXCEL_PRINT_GRIDLINES read-only value ABAP_FALSE. "#EC NOTEXT . " .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .
    data SHEET_CONTENT type ZEXCEL_T_CELL_DATA .
    data SHEET_SETUP type ref to ZCL_EXCEL_SHEET_SETUP .
   data SHOW_GRIDLINES type ZEXCEL_SHOW_GRIDLINES read-only value ABAP_TRUE. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .
    data SHOW_ROWCOLHEADERS type ZEXCEL_SHOW_GRIDLINES read-only value ABAP_TRUE. "#EC NOTEXT . " .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .
    data STYLES type ZEXCEL_T_SHEET_STYLE .
    data TABCOLOR type ZEXCEL_S_TABCOLOR read-only .

    methods ADD_DRAWING
      importing
        !IP_DRAWING type ref to ZCL_EXCEL_DRAWING .
    methods ADD_NEW_COLUMN
      importing
        !IP_COLUMN       type SIMPLE
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
        !IP_ROW       type SIMPLE
      returning
        value(EO_ROW) type ref to ZCL_EXCEL_ROW .
    methods BIND_ALV
      importing
        !IO_ALV      type ref to OBJECT
        !IT_TABLE    type standard table
        !I_TOP       type I default 1
        !I_LEFT      type I default 1
        !TABLE_STYLE type ZEXCEL_TABLE_STYLE optional
        !I_TABLE     type ABAP_BOOL default ABAP_TRUE
      raising
        ZCX_EXCEL .
    type-pools SLIS .
    type-pools SOI .
    methods BIND_ALV_OLE2
      importing
        !I_DOCUMENT_URL      type CHAR255 default SPACE
        !I_XLS               type C default SPACE
        !I_SAVE_PATH         type STRING
        !IO_ALV              type ref to CL_GUI_ALV_GRID
        !IT_LISTHEADER       type SLIS_T_LISTHEADER optional
        !I_TOP               type I default 1
        !I_LEFT              type I default 1
        !I_COLUMNS_HEADER    type C default 'X'
        !I_COLUMNS_AUTOFIT   type C default 'X'
        !I_FORMAT_COL_HEADER type SOI_FORMAT_ITEM optional
        !I_FORMAT_SUBTOTAL   type SOI_FORMAT_ITEM optional
        !I_FORMAT_TOTAL      type SOI_FORMAT_ITEM optional
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
        !IP_TABLE               type standard table
        !IT_FIELD_CATALOG       type ZEXCEL_T_FIELDCATALOG optional
        !IS_TABLE_SETTINGS      type ZEXCEL_S_TABLE_SETTINGS optional
        value(IV_DEFAULT_DESCR) type C optional
      exporting
        !ES_TABLE_SETTINGS      type ZEXCEL_S_TABLE_SETTINGS
      raising
        ZCX_EXCEL .
    methods CALCULATE_COLUMN_WIDTHS
      raising
        ZCX_EXCEL .
    methods CHANGE_CELL_STYLE
      importing
        !IP_COLUMN                      type SIMPLE
        !IP_ROW                         type ZEXCEL_CELL_ROW
        !IP_COMPLETE                    type ZEXCEL_S_CSTYLE_COMPLETE optional
        !IP_XCOMPLETE                   type ZEXCEL_S_CSTYLEX_COMPLETE optional
        !IP_FONT                        type ZEXCEL_S_CSTYLE_FONT optional
        !IP_XFONT                       type ZEXCEL_S_CSTYLEX_FONT optional
        !IP_FILL                        type ZEXCEL_S_CSTYLE_FILL optional
        !IP_XFILL                       type ZEXCEL_S_CSTYLEX_FILL optional
        !IP_BORDERS                     type ZEXCEL_S_CSTYLE_BORDERS optional
        !IP_XBORDERS                    type ZEXCEL_S_CSTYLEX_BORDERS optional
        !IP_ALIGNMENT                   type ZEXCEL_S_CSTYLE_ALIGNMENT optional
        !IP_XALIGNMENT                  type ZEXCEL_S_CSTYLEX_ALIGNMENT optional
        !IP_NUMBER_FORMAT_FORMAT_CODE   type ZEXCEL_NUMBER_FORMAT optional
        !IP_PROTECTION                  type ZEXCEL_S_CSTYLE_PROTECTION optional
        !IP_XPROTECTION                 type ZEXCEL_S_CSTYLEX_PROTECTION optional
        !IP_FONT_BOLD                   type FLAG optional
        !IP_FONT_COLOR                  type ZEXCEL_S_STYLE_COLOR optional
        !IP_FONT_COLOR_RGB              type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_FONT_COLOR_INDEXED          type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_FONT_COLOR_THEME            type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_FONT_COLOR_TINT             type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_FONT_FAMILY                 type ZEXCEL_STYLE_FONT_FAMILY optional
        !IP_FONT_ITALIC                 type FLAG optional
        !IP_FONT_NAME                   type ZEXCEL_STYLE_FONT_NAME optional
        !IP_FONT_SCHEME                 type ZEXCEL_STYLE_FONT_SCHEME optional
        !IP_FONT_SIZE                   type ZEXCEL_STYLE_FONT_SIZE optional
        !IP_FONT_STRIKETHROUGH          type FLAG optional
        !IP_FONT_UNDERLINE              type FLAG optional
        !IP_FONT_UNDERLINE_MODE         type ZEXCEL_STYLE_FONT_UNDERLINE optional
        !IP_FILL_FILLTYPE               type ZEXCEL_FILL_TYPE optional
        !IP_FILL_ROTATION               type ZEXCEL_ROTATION optional
        !IP_FILL_FGCOLOR                type ZEXCEL_S_STYLE_COLOR optional
        !IP_FILL_FGCOLOR_RGB            type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_FILL_FGCOLOR_INDEXED        type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_FILL_FGCOLOR_THEME          type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_FILL_FGCOLOR_TINT           type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_FILL_BGCOLOR                type ZEXCEL_S_STYLE_COLOR optional
        !IP_FILL_BGCOLOR_RGB            type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_FILL_BGCOLOR_INDEXED        type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_FILL_BGCOLOR_THEME          type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_FILL_BGCOLOR_TINT           type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_BORDERS_ALLBORDERS          type ZEXCEL_S_CSTYLE_BORDER optional
        !IP_FILL_GRADTYPE_TYPE          type ZEXCEL_S_GRADIENT_TYPE-TYPE optional
        !IP_FILL_GRADTYPE_DEGREE        type ZEXCEL_S_GRADIENT_TYPE-DEGREE optional
        !IP_XBORDERS_ALLBORDERS         type ZEXCEL_S_CSTYLEX_BORDER optional
        !IP_BORDERS_DIAGONAL            type ZEXCEL_S_CSTYLE_BORDER optional
        !IP_FILL_GRADTYPE_BOTTOM        type ZEXCEL_S_GRADIENT_TYPE-BOTTOM optional
        !IP_FILL_GRADTYPE_TOP           type ZEXCEL_S_GRADIENT_TYPE-TOP optional
        !IP_XBORDERS_DIAGONAL           type ZEXCEL_S_CSTYLEX_BORDER optional
        !IP_BORDERS_DIAGONAL_MODE       type ZEXCEL_DIAGONAL optional
        !IP_FILL_GRADTYPE_RIGHT         type ZEXCEL_S_GRADIENT_TYPE-RIGHT optional
        !IP_BORDERS_DOWN                type ZEXCEL_S_CSTYLE_BORDER optional
        !IP_FILL_GRADTYPE_LEFT          type ZEXCEL_S_GRADIENT_TYPE-LEFT optional
        !IP_FILL_GRADTYPE_POSITION1     type ZEXCEL_S_GRADIENT_TYPE-POSITION1 optional
        !IP_XBORDERS_DOWN               type ZEXCEL_S_CSTYLEX_BORDER optional
        !IP_BORDERS_LEFT                type ZEXCEL_S_CSTYLE_BORDER optional
        !IP_FILL_GRADTYPE_POSITION2     type ZEXCEL_S_GRADIENT_TYPE-POSITION2 optional
        !IP_FILL_GRADTYPE_POSITION3     type ZEXCEL_S_GRADIENT_TYPE-POSITION3 optional
        !IP_XBORDERS_LEFT               type ZEXCEL_S_CSTYLEX_BORDER optional
        !IP_BORDERS_RIGHT               type ZEXCEL_S_CSTYLE_BORDER optional
        !IP_XBORDERS_RIGHT              type ZEXCEL_S_CSTYLEX_BORDER optional
        !IP_BORDERS_TOP                 type ZEXCEL_S_CSTYLE_BORDER optional
        !IP_XBORDERS_TOP                type ZEXCEL_S_CSTYLEX_BORDER optional
        !IP_ALIGNMENT_HORIZONTAL        type ZEXCEL_ALIGNMENT optional
        !IP_ALIGNMENT_VERTICAL          type ZEXCEL_ALIGNMENT optional
        !IP_ALIGNMENT_TEXTROTATION      type ZEXCEL_TEXT_ROTATION optional
        !IP_ALIGNMENT_WRAPTEXT          type FLAG optional
        !IP_ALIGNMENT_SHRINKTOFIT       type FLAG optional
        !IP_ALIGNMENT_INDENT            type ZEXCEL_INDENT optional
        !IP_PROTECTION_HIDDEN           type ZEXCEL_CELL_PROTECTION optional
        !IP_PROTECTION_LOCKED           type ZEXCEL_CELL_PROTECTION optional
        !IP_BORDERS_ALLBORDERS_STYLE    type ZEXCEL_BORDER optional
        !IP_BORDERS_ALLBORDERS_COLOR    type ZEXCEL_S_STYLE_COLOR optional
        !IP_BORDERS_ALLBO_COLOR_RGB     type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_BORDERS_ALLBO_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_BORDERS_ALLBO_COLOR_THEME   type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_BORDERS_ALLBO_COLOR_TINT    type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_BORDERS_DIAGONAL_STYLE      type ZEXCEL_BORDER optional
        !IP_BORDERS_DIAGONAL_COLOR      type ZEXCEL_S_STYLE_COLOR optional
        !IP_BORDERS_DIAGONAL_COLOR_RGB  type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_BORDERS_DIAGONAL_COLOR_INDE type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_BORDERS_DIAGONAL_COLOR_THEM type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_BORDERS_DIAGONAL_COLOR_TINT type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_BORDERS_DOWN_STYLE          type ZEXCEL_BORDER optional
        !IP_BORDERS_DOWN_COLOR          type ZEXCEL_S_STYLE_COLOR optional
        !IP_BORDERS_DOWN_COLOR_RGB      type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_BORDERS_DOWN_COLOR_INDEXED  type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_BORDERS_DOWN_COLOR_THEME    type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_BORDERS_DOWN_COLOR_TINT     type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_BORDERS_LEFT_STYLE          type ZEXCEL_BORDER optional
        !IP_BORDERS_LEFT_COLOR          type ZEXCEL_S_STYLE_COLOR optional
        !IP_BORDERS_LEFT_COLOR_RGB      type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_BORDERS_LEFT_COLOR_INDEXED  type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_BORDERS_LEFT_COLOR_THEME    type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_BORDERS_LEFT_COLOR_TINT     type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_BORDERS_RIGHT_STYLE         type ZEXCEL_BORDER optional
        !IP_BORDERS_RIGHT_COLOR         type ZEXCEL_S_STYLE_COLOR optional
        !IP_BORDERS_RIGHT_COLOR_RGB     type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_BORDERS_RIGHT_COLOR_INDEXED type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_BORDERS_RIGHT_COLOR_THEME   type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_BORDERS_RIGHT_COLOR_TINT    type ZEXCEL_STYLE_COLOR_TINT optional
        !IP_BORDERS_TOP_STYLE           type ZEXCEL_BORDER optional
        !IP_BORDERS_TOP_COLOR           type ZEXCEL_S_STYLE_COLOR optional
        !IP_BORDERS_TOP_COLOR_RGB       type ZEXCEL_STYLE_COLOR_ARGB optional
        !IP_BORDERS_TOP_COLOR_INDEXED   type ZEXCEL_STYLE_COLOR_INDEXED optional
        !IP_BORDERS_TOP_COLOR_THEME     type ZEXCEL_STYLE_COLOR_THEME optional
        !IP_BORDERS_TOP_COLOR_TINT      type ZEXCEL_STYLE_COLOR_TINT optional
      returning
        value(EP_GUID)                  type ZEXCEL_CELL_STYLE
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
        !IP_CELL_ROW    type ZEXCEL_CELL_ROW optional
      raising
        ZCX_EXCEL .
    methods DELETE_ROW_OUTLINE
      importing
        !IV_ROW_FROM type I
        !IV_ROW_TO   type I
      raising
        ZCX_EXCEL .
    methods FREEZE_PANES
      importing
        !IP_NUM_COLUMNS type I optional
        !IP_NUM_ROWS    type I optional
      raising
        ZCX_EXCEL .
    methods GET_ACTIVE_CELL
      returning
        value(EP_ACTIVE_CELL) type STRING
      raising
        ZCX_EXCEL .
    methods GET_CELL
      importing
        !IP_COLUMN  type SIMPLE
        !IP_ROW     type ZEXCEL_CELL_ROW
      exporting
        !EP_VALUE   type ZEXCEL_CELL_VALUE
        !EP_RC      type SYSUBRC
        !EP_STYLE   type ref to ZCL_EXCEL_STYLE
        !EP_GUID    type ZEXCEL_CELL_STYLE
        !EP_FORMULA type ZEXCEL_CELL_FORMULA
      raising
        ZCX_EXCEL .
    methods GET_COLUMN
      importing
        !IP_COLUMN       type SIMPLE
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
        !IP_TYPE          type ZEXCEL_DRAWING_TYPE optional
      returning
        value(R_DRAWINGS) type ref to ZCL_EXCEL_DRAWINGS .
    methods GET_DRAWINGS_ITERATOR
      importing
        !IP_TYPE           type ZEXCEL_DRAWING_TYPE
      returning
        value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
    methods GET_FREEZE_CELL
      exporting
        !EP_ROW    type ZEXCEL_CELL_ROW
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
        !IP_ROW       type INT4
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
        !IP_GUID             type ZEXCEL_CELL_STYLE
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
        !IP_ESCAPED     type FLAG default ''
      returning
        value(EP_TITLE) type ZEXCEL_SHEET_TITLE .
    methods IS_CELL_MERGED
      importing
        !IP_COLUMN          type SIMPLE
        !IP_ROW             type ZEXCEL_CELL_ROW
      returning
        value(RP_IS_MERGED) type ABAP_BOOL
      raising
        ZCX_EXCEL .
    methods SET_CELL
      importing
        !IP_COLUMN    type SIMPLE
        !IP_ROW       type ZEXCEL_CELL_ROW
        !IP_VALUE     type SIMPLE optional
        !IP_FORMULA   type ZEXCEL_CELL_FORMULA optional
        !IP_STYLE     type ZEXCEL_CELL_STYLE optional
        !IP_HYPERLINK type ref to ZCL_EXCEL_HYPERLINK optional
        !IP_DATA_TYPE type ZEXCEL_CELL_DATA_TYPE optional
        !IP_ABAP_TYPE type ABAP_TYPEKIND optional
      raising
        ZCX_EXCEL .
    methods SET_CELL_FORMULA
      importing
        !IP_COLUMN  type SIMPLE
        !IP_ROW     type ZEXCEL_CELL_ROW
        !IP_FORMULA type ZEXCEL_CELL_FORMULA
      raising
        ZCX_EXCEL .
    methods SET_CELL_STYLE
      importing
        !IP_COLUMN type SIMPLE
        !IP_ROW    type ZEXCEL_CELL_ROW
        !IP_STYLE  type ZEXCEL_CELL_STYLE
      raising
        ZCX_EXCEL .
    methods SET_COLUMN_WIDTH
      importing
        !IP_COLUMN         type SIMPLE
        !IP_WIDTH_FIX      type SIMPLE default 0
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
        !IP_COLUMN_START type SIMPLE default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MIN_COL
        !IP_COLUMN_END   type SIMPLE default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MAX_COL
        !IP_ROW          type ZEXCEL_CELL_ROW default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MIN_ROW
        !IP_ROW_TO       type ZEXCEL_CELL_ROW default ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MAX_ROW
      !IP_STYLE type ZEXCEL_CELL_STYLE optional "added parameter
      !IP_VALUE type SIMPLE optional "added parameter
      !IP_FORMULA type ZEXCEL_CELL_FORMULA optional "added parameter
      raising
        ZCX_EXCEL .
    methods SET_PRINT_GRIDLINES
      importing
        !I_PRINT_GRIDLINES type ZEXCEL_PRINT_GRIDLINES .
    methods SET_ROW_HEIGHT
      importing
        !IP_ROW        type SIMPLE
        !IP_HEIGHT_FIX type SIMPLE
      raising
        ZCX_EXCEL .
    methods SET_ROW_OUTLINE
      importing
        !IV_ROW_FROM  type I
        !IV_ROW_TO    type I
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
        !IP_TABLE           type standard table
        !IP_HDR_STYLE       type ZEXCEL_CELL_STYLE optional
        !IP_BODY_STYLE      type ZEXCEL_CELL_STYLE optional
        !IP_TABLE_TITLE     type STRING
        !IP_TOP_LEFT_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA default 'B'
        !IP_TOP_LEFT_ROW    type ZEXCEL_CELL_ROW default 3
        !IP_TRANSPOSE       type XFELD optional
        !IP_NO_HEADER       type XFELD optional
      raising
        ZCX_EXCEL .
    methods SET_TITLE
      importing
        !IP_TITLE type ZEXCEL_SHEET_TITLE
      raising
        ZCX_EXCEL .
    methods GET_TABLE
      importing
        IV_SKIPPED_ROWS type INT4  default 0
        IV_SKIPPED_COLS type INT4  default 0
      exporting
        ET_TABLE        type standard table
      raising
        ZCX_EXCEL.

methods SET_MERGE_STYLE
    importing
      !IP_COLUMN_START type simple optional
      !IP_COLUMN_END type simple optional
      !IP_ROW type ZEXCEL_CELL_ROW optional
      !IP_ROW_TO type ZEXCEL_CELL_ROW optional
      !IP_STYLE type ZEXCEL_CELL_STYLE optional .

methods SET_AREA_FORMULA
    importing
      !IP_COLUMN_START type simple
      !IP_COLUMN_END type simple optional
      !IP_ROW type ZEXCEL_CELL_ROW
      !IP_ROW_TO type ZEXCEL_CELL_ROW optional
      !IP_FORMULA type ZEXCEL_CELL_FORMULA
      !IP_MERGE type ABAP_BOOL optional
    raising
      ZCX_EXCEL .

methods SET_AREA_STYLE
    importing
      !IP_COLUMN_START type simple
      !IP_COLUMN_END type simple optional
      !IP_ROW type ZEXCEL_CELL_ROW
      !IP_ROW_TO type ZEXCEL_CELL_ROW optional
      !IP_STYLE type ZEXCEL_CELL_STYLE
      !IP_MERGE type ABAP_BOOL optional .

 methods SET_AREA
    importing
      !IP_COLUMN_START type simple
      !IP_COLUMN_END type simple optional
      !IP_ROW type ZEXCEL_CELL_ROW
      !IP_ROW_TO type ZEXCEL_CELL_ROW optional
      !IP_VALUE type SIMPLE optional
      !IP_FORMULA type ZEXCEL_CELL_FORMULA optional
      !IP_STYLE type ZEXCEL_CELL_STYLE optional
      !IP_HYPERLINK type ref to ZCL_EXCEL_HYPERLINK optional
      !IP_DATA_TYPE type ZEXCEL_CELL_DATA_TYPE optional
      !IP_ABAP_TYPE type ABAP_TYPEKIND optional
      !IP_MERGE type ABAP_BOOL optional
    raising
      ZCX_EXCEL .

*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
  protected section.
  private section.

    types:
      begin of MTY_S_FONT_METRIC,
        CHAR       type C length 1,
        CHAR_WIDTH type TDCWIDTHS,
      end of MTY_S_FONT_METRIC .
    types:
      MTY_TH_FONT_METRICS
             type hashed table of MTY_S_FONT_METRIC
             with unique key CHAR .
    types:
      begin of MTY_S_FONT_CACHE,
        FONT_NAME       type ZEXCEL_STYLE_FONT_NAME,
        FONT_HEIGHT     type TDFONTSIZE,
        FLAG_BOLD       type ABAP_BOOL,
        FLAG_ITALIC     type ABAP_BOOL,
        TH_FONT_METRICS type MTY_TH_FONT_METRICS,
      end of MTY_S_FONT_CACHE .
    types:
      MTY_TH_FONT_CACHE
             type hashed table of MTY_S_FONT_CACHE
             with unique key FONT_NAME FONT_HEIGHT FLAG_BOLD FLAG_ITALIC .
    types:
*  types:
*    mty_ts_row_dimension TYPE SORTED TABLE OF zexcel_s_worksheet_rowdimensio WITH UNIQUE KEY row .
      begin of MTY_MERGE,
        ROW_FROM type I,
        ROW_TO   type I,
        COL_FROM type I,
        COL_TO   type I,
      end of MTY_MERGE .
    types:
      MTY_TS_MERGE type sorted table of MTY_MERGE with unique key TABLE_LINE .

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
    data TITLE type ZEXCEL_SHEET_TITLE value 'Worksheet'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    data UPPER_CELL type ZEXCEL_S_CELL_DATA .

    methods CALCULATE_CELL_WIDTH
      importing
        !IP_COLUMN      type SIMPLE
        !IP_ROW         type ZEXCEL_CELL_ROW
      returning
        value(EP_WIDTH) type FLOAT
      raising
        ZCX_EXCEL .
    methods GENERATE_TITLE
      returning
        value(EP_TITLE) type ZEXCEL_SHEET_TITLE .
    methods GET_VALUE_TYPE
      importing
        !IP_VALUE      type SIMPLE
      exporting
        !EP_VALUE      type SIMPLE
        !EP_VALUE_TYPE type ABAP_TYPEKIND .
    methods PRINT_TITLE_SET_RANGE .
    methods UPDATE_DIMENSION_RANGE
      raising
        ZCX_EXCEL .
ENDCLASS.



CLASS ZCL_EXCEL_WORKSHEET IMPLEMENTATION.


  method ADD_DRAWING.
    case IP_DRAWING->GET_TYPE( ).
      when ZCL_EXCEL_DRAWING=>TYPE_IMAGE.
        DRAWINGS->INCLUDE( IP_DRAWING ).
      when ZCL_EXCEL_DRAWING=>TYPE_CHART.
        CHARTS->INCLUDE( IP_DRAWING ).
    endcase.
  endmethod.


  method ADD_NEW_COLUMN.
    data: LV_COLUMN_ALPHA type ZEXCEL_CELL_COLUMN_ALPHA.

    LV_COLUMN_ALPHA = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( IP_COLUMN ).

    create object EO_COLUMN
      exporting
        IP_INDEX     = LV_COLUMN_ALPHA
        IP_EXCEL     = ME->EXCEL
        IP_WORKSHEET = ME.
    COLUMNS->ADD( EO_COLUMN ).
  endmethod.


  method ADD_NEW_DATA_VALIDATION.

    create object EO_DATA_VALIDATION.
    DATA_VALIDATIONS->ADD( EO_DATA_VALIDATION ).
  endmethod.


  method ADD_NEW_RANGE.
* Create default blank range
    create object EO_RANGE.
    RANGES->ADD( EO_RANGE ).
  endmethod.


  method ADD_NEW_ROW.
    create object EO_ROW
      exporting
        IP_INDEX = IP_ROW.
    ROWS->ADD( EO_ROW ).
  endmethod.


  method ADD_NEW_STYLE_COND.
    create object EO_STYLE_COND.
    STYLES_COND->ADD( EO_STYLE_COND ).
  endmethod.


  method BIND_ALV.
    data: LO_CONVERTER type ref to ZCL_EXCEL_CONVERTER.

    create object LO_CONVERTER.

    try.
        LO_CONVERTER->CONVERT(
          exporting
            IO_ALV         = IO_ALV
            IT_TABLE       = IT_TABLE
            I_ROW_INT      = I_TOP
            I_COLUMN_INT   = I_LEFT
            I_TABLE        = I_TABLE
            I_STYLE_TABLE  = TABLE_STYLE
            IO_WORKSHEET   = ME
          changing
            CO_EXCEL       = EXCEL ).
      catch ZCX_EXCEL .
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

    data: LO_CONTROL  type ref to I_OI_CONTAINER_CONTROL.
    data: LO_PROXY    type ref to I_OI_DOCUMENT_PROXY.
    data: LO_SPREADSHEET type ref to I_OI_SPREADSHEET.
    data: LO_ERROR    type ref to I_OI_ERROR.
    data: LC_RETCODE  type SOI_RET_STRING.
    data: LI_HAS      type I. "Proxy has spreadsheet interface?
    data: L_IS_CLOSED type I.

* Data for session 1: Get LVC data from ALV object
* ------------------------------------------

    data: L_HAS_ACTIVEX,
          L_DOCTYPE_EXCEL_SHEET(11) type C.

* LVC
    data: LT_FIELDCAT_LVC       type LVC_T_FCAT.
    data: WA_FIELDCAT_LVC       type LVC_S_FCAT.
    data: LT_SORT_LVC           type LVC_T_SORT.
    data: LT_FILTER_IDX_LVC     type LVC_T_FIDX.
    data: LT_GROUPLEVELS_LVC    type LVC_T_GRPL.

* KKBLO
    data: LT_FIELDCAT_KKBLO     type  KKBLO_T_FIELDCAT.
    data: LT_SORT_KKBLO         type  KKBLO_T_SORTINFO.
    data: LT_GROUPLEVELS_KKBLO  type  KKBLO_T_GROUPLEVELS.
    data: LT_FILTER_IDX_KKBLO   type  KKBLO_T_SFINFO.
    data: WA_LISTHEADER         like line of IT_LISTHEADER.

* Subtotal
    data: LT_COLLECT00          type ref to DATA.
    data: LT_COLLECT01          type ref to DATA.
    data: LT_COLLECT02          type ref to DATA.
    data: LT_COLLECT03          type ref to DATA.
    data: LT_COLLECT04          type ref to DATA.
    data: LT_COLLECT05          type ref to DATA.
    data: LT_COLLECT06          type ref to DATA.
    data: LT_COLLECT07          type ref to DATA.
    data: LT_COLLECT08          type ref to DATA.
    data: LT_COLLECT09          type ref to DATA.

* data table name
    data: L_TABNAME             type  KKBLO_TABNAME.

* local object
    data: LO_GRID               type ref to LCL_GUI_ALV_GRID.

* data table get from ALV
    data: LT_ALV                  type ref to DATA.

* total / subtotal data
    field-symbols: <F_COLLECT00>  type standard table.
    field-symbols: <F_COLLECT01>  type standard table.
    field-symbols: <F_COLLECT02>  type standard table.
    field-symbols: <F_COLLECT03>  type standard table.
    field-symbols: <F_COLLECT04>  type standard table.
    field-symbols: <F_COLLECT05>  type standard table.
    field-symbols: <F_COLLECT06>  type standard table.
    field-symbols: <F_COLLECT07>  type standard table.
    field-symbols: <F_COLLECT08>  type standard table.
    field-symbols: <F_COLLECT09>  type standard table.

* table before append subtotal lines
    field-symbols: <F_ALV_TAB>    type standard table.

* data for session 2: sort, filter and calculate total/subtotal
* ------------------------------------------

* table to save index of subotal / total line in excel tanle
* this ideal to control index of subtotal / total line later
* for ex, when get subtotal / total line to format
    types: begin of ST_SUBTOT_INDEXS,
             INDEX type I,
           end of ST_SUBTOT_INDEXS.
    data: LT_SUBTOT_INDEXS type table of ST_SUBTOT_INDEXS.
    data: WA_SUBTOT_INDEXS like line of LT_SUBTOT_INDEXS.

* data table after append subtotal
    data: LT_EXCEL                type ref to DATA.

    data: L_TABIX                 type I.
    data: L_SAVE_INDEX            type I.

* dyn subtotal table name
    data: L_COLLECT               type STRING.

* subtotal range, to format subtotal (and total)
    data: SUBRANGES               type SOI_RANGE_LIST.
    data: SUBRANGEITEM            type SOI_RANGE_ITEM.
    data: L_SUB_INDEX             type I.


* table after append subtotal lines
    field-symbols: <F_EXCEL_TAB>  type standard table.
    field-symbols: <F_EXCEL_LINE> type ANY.

* dyn subtotal tables
    field-symbols: <F_COLLECT_TAB>      type standard table.
    field-symbols: <F_COLLECT_LINE>     type ANY.

    field-symbols: <F_FILTER_IDX_LINE>  like line of LT_FILTER_IDX_KKBLO.
    field-symbols: <F_FIELDCAT_LINE>    like line of LT_FIELDCAT_KKBLO.
    field-symbols: <F_GROUPLEVELS_LINE> like line of LT_GROUPLEVELS_KKBLO.
    field-symbols: <F_LINE>             type ANY.

* Data for session 3: map data to semantic table
* ------------------------------------------

    types: begin of ST_COLUMN_INDEX,
             FIELDNAME type KKBLO_FIELDNAME,
             TABNAME   type KKBLO_TABNAME,
             COL       like SY-INDEX,
           end of ST_COLUMN_INDEX.

* columns index
    data: LT_COLUMN_INDEX   type table of ST_COLUMN_INDEX.
    data: WA_COLUMN_INDEX   like line of LT_COLUMN_INDEX.

* table of dependent field ( currency and quantity unit field)
    data: LT_FIELDCAT_DEPF  type KKBLO_T_FIELDCAT.
    data: WA_FIELDCAT_DEPF  type KKBLO_FIELDCAT.

* XXL interface:
* -XXL: contain exporting columns characteristic
    data: LT_SEMA type table of GXXLT_S initial size 0.
    data: WA_SEMA like line of LT_SEMA.

* -XXL interface: header
    data: LT_HKEY type table of GXXLT_H initial size 0.
    data: WA_HKEY like line of LT_HKEY.

* -XXL interface: header keys
    data: LT_VKEY type table of GXXLT_V initial size 0.
    data: WA_VKEY like line of LT_VKEY.

* Number of H Keys: number of key columns
    data: L_N_HRZ_KEYS      type  I.
* Number of data columns in the list object: non-key columns no
    data: L_N_ATT_COLS      type  I.
* Number of V Keys: number of header row
    data: L_N_VRT_KEYS      type  I.

* curency to format amount
    data: LT_TCURX          type table of TCURX.
    data: WA_TCURX          like line of LT_TCURX.
    data: L_DEF             type FLAG. " currency / quantity flag
    data: WA_T006           type T006. " decimal place of unit

    data: L_NUM             type I. " table columns number
    data: L_TYP             type C. " table type
    data: WA                type ref to DATA.
    data: L_INT             type I.
    data: L_COUNTER         type I.

    field-symbols: <F_EXCEL_COLUMN>     type ANY.
    field-symbols: <F_FCAT_COLUMN>      type ANY.

* Data for session 4: write to excel
* ------------------------------------------

    data: SEMA_TYPE         type  C.

    data L_ERROR           type ref to C_OI_PROXY_ERROR.
    data COUNT              type I.
    data DATAC              type I.
    data DATAREAL           type I. " exporting column number
    data VKEYCOUNT          type I.
    data ALL type I.
    data MIT type I         value 1.  " index of recent row?
    data LI_COL_POS type I  value 1.  " column position
    data LI_COL_NUM type I.           " table columns number
    field-symbols: <LINE>   type ANY.
    field-symbols: <ITEM>   type ANY.

    data TD                 type SYDES_DESC.

    data: TYP.
    data: RANGES             type SOI_RANGE_LIST.
    data: RANGEITEM          type SOI_RANGE_ITEM.
    data: CONTENTS           type SOI_GENERIC_TABLE.
    data: CONTENTSITEM       type SOI_GENERIC_ITEM.
    data: SEMAITEM           type GXXLT_S.
    data: HKEYITEM           type GXXLT_H.
    data: VKEYITEM           type GXXLT_V.
    data: LI_COMMENTARY_ROWS type I.  "row number of title lines + 1
    data: LO_ERROR_W         type ref to  I_OI_ERROR.
    data: L_RETCODE          type SOI_RET_STRING.
    data: NO_FLUSH           type C value 'X'.
    data: LI_HEAD_TOP        type I. "header rows position

* Data for session 5: Save and clode document
* ------------------------------------------

    data: LI_DOCUMENT_SIZE   type I.
    data: LS_PATH            type RLGRAP-FILENAME.

* MACRO: Close_document
*-------------------------------------------

    define CLOSE_DOCUMENT.
      clear: L_IS_CLOSED.
      if LO_PROXY is not initial.

* check proxy detroyed adi

        call method LO_PROXY->IS_DESTROYED
          importing
            RET_VALUE = L_IS_CLOSED.

* if dun detroyed yet: close -> release proxy

        if L_IS_CLOSED is initial.
          call method LO_PROXY->CLOSE_DOCUMENT
*        EXPORTING
*          do_save = do_save
            importing
              ERROR       = LO_ERROR
              RETCODE     = LC_RETCODE.
        endif.

        call method LO_PROXY->RELEASE_DOCUMENT
          importing
            ERROR   = LO_ERROR
            RETCODE = LC_RETCODE.

      else.
        LC_RETCODE = C_OI_ERRORS=>RET_DOCUMENT_NOT_OPEN.
      endif.

* Detroy control container

      if LO_CONTROL is not initial.
        call method LO_CONTROL->DESTROY_CONTROL.
      endif.

      clear:
        LO_SPREADSHEET,
        LO_PROXY,
        LO_CONTROL.

* free local

      clear: L_IS_CLOSED.

    end-of-definition.

* Macro to catch DOI error
*-------------------------------------------

    define ERROR_DOI.
      if LC_RETCODE ne C_OI_ERRORS=>RET_OK.
        CLOSE_DOCUMENT.
        call method LO_ERROR->RAISE_MESSAGE
          exporting
            TYPE = 'E'.
        clear: LO_ERROR.
      endif.
    end-of-definition.

*--------------------------------------------------------------------*
* SESSION 0: DOI CONSTRUCTOR
*--------------------------------------------------------------------*

* check active windown

    call function 'GUI_HAS_ACTIVEX'
      importing
        RETURN = L_HAS_ACTIVEX.

    if L_HAS_ACTIVEX is initial.
      raise MISS_GUIDE.
    endif.

*   Get Container Object of Screen

    call method C_OI_CONTAINER_CONTROL_CREATOR=>GET_CONTAINER_CONTROL
      importing
        CONTROL = LO_CONTROL
        RETCODE = LC_RETCODE.

    ERROR_DOI.

* Initialize Container control

    call method LO_CONTROL->INIT_CONTROL
      exporting
        PARENT                   = CL_GUI_CONTAINER=>DEFAULT_SCREEN
        R3_APPLICATION_NAME      = ''
        INPLACE_ENABLED          = 'X'
        NO_FLUSH                 = 'X'
        REGISTER_ON_CLOSE_EVENT  = 'X'
        REGISTER_ON_CUSTOM_EVENT = 'X'
      importing
        ERROR                    = LO_ERROR
        RETCODE                  = LC_RETCODE.

    ERROR_DOI.

* Get Proxy Document:
* check exist of document proxy, if exist -> close first

    if not LO_PROXY is initial.
      CLOSE_DOCUMENT.
    endif.

    if I_XLS is not initial.
* xls format, doctype = soi_doctype_excel97_sheet
      L_DOCTYPE_EXCEL_SHEET = 'Excel.Sheet.8'.
    else.
* xlsx format, doctype = soi_doctype_excel_sheet
      L_DOCTYPE_EXCEL_SHEET = 'Excel.Sheet'.
    endif.

    call method LO_CONTROL->GET_DOCUMENT_PROXY
      exporting
        DOCUMENT_TYPE      = L_DOCTYPE_EXCEL_SHEET
        REGISTER_CONTAINER = 'X'
      importing
        DOCUMENT_PROXY     = LO_PROXY
        ERROR              = LO_ERROR
        RETCODE            = LC_RETCODE.

    ERROR_DOI.

    if I_DOCUMENT_URL is initial.

* create new excel document

      call method LO_PROXY->CREATE_DOCUMENT
        exporting
          CREATE_VIEW_DATA = 'X'
          OPEN_INPLACE     = 'X'
          NO_FLUSH         = 'X'
        importing
          ERROR            = LO_ERROR
          RETCODE          = LC_RETCODE.

      ERROR_DOI.

    else.

* Read excel template for i_DOCUMENT_URL
* this excel template can be store in local or server

      call method LO_PROXY->OPEN_DOCUMENT
        exporting
          DOCUMENT_URL = I_DOCUMENT_URL
          OPEN_INPLACE = 'X'
          NO_FLUSH     = 'X'
        importing
          ERROR        = LO_ERROR
          RETCODE      = LC_RETCODE.

      ERROR_DOI.

    endif.

* Check Spreadsheet Interface of Document Proxy

    call method LO_PROXY->HAS_SPREADSHEET_INTERFACE
      importing
        IS_AVAILABLE = LI_HAS
        ERROR        = LO_ERROR
        RETCODE      = LC_RETCODE.

    ERROR_DOI.

* create Spreadsheet object

    check LI_HAS is not initial.

    call method LO_PROXY->GET_SPREADSHEET_INTERFACE
      importing
        SHEET_INTERFACE = LO_SPREADSHEET
        ERROR           = LO_ERROR
        RETCODE         = LC_RETCODE.

    ERROR_DOI.

*--------------------------------------------------------------------*
* SESSION 1: GET LVC DATA FROM ALV OBJECT
*--------------------------------------------------------------------*

* data table

    create object LO_GRID
      exporting
        I_PARENT = CL_GUI_CONTAINER=>SCREEN0.

    call method LO_GRID->GET_ALV_ATTRIBUTES
      exporting
        IO_GRID  = IO_ALV
      importing
        ET_TABLE = LT_ALV.

    assign LT_ALV->* to <F_ALV_TAB>.

* fieldcat

    call method IO_ALV->GET_FRONTEND_FIELDCATALOG
      importing
        ET_FIELDCATALOG = LT_FIELDCAT_LVC.

* table name

    loop at LT_FIELDCAT_LVC into WA_FIELDCAT_LVC
    where not TABNAME is initial.
      L_TABNAME = WA_FIELDCAT_LVC-TABNAME.
      exit.
    endloop.

    if SY-SUBRC ne 0.
      L_TABNAME = '1'.
    endif.
    clear: WA_FIELDCAT_LVC.

* sort table

    call method IO_ALV->GET_SORT_CRITERIA
      importing
        ET_SORT = LT_SORT_LVC.


* filter index

    call method IO_ALV->GET_FILTERED_ENTRIES
      importing
        ET_FILTERED_ENTRIES = LT_FILTER_IDX_LVC.

* group level + subtotal

    call method IO_ALV->GET_SUBTOTALS
      importing
        EP_COLLECT00   = LT_COLLECT00
        EP_COLLECT01   = LT_COLLECT01
        EP_COLLECT02   = LT_COLLECT02
        EP_COLLECT03   = LT_COLLECT03
        EP_COLLECT04   = LT_COLLECT04
        EP_COLLECT05   = LT_COLLECT05
        EP_COLLECT06   = LT_COLLECT06
        EP_COLLECT07   = LT_COLLECT07
        EP_COLLECT08   = LT_COLLECT08
        EP_COLLECT09   = LT_COLLECT09
        ET_GROUPLEVELS = LT_GROUPLEVELS_LVC.

    assign LT_COLLECT00->* to <F_COLLECT00>.
    assign LT_COLLECT01->* to <F_COLLECT01>.
    assign LT_COLLECT02->* to <F_COLLECT02>.
    assign LT_COLLECT03->* to <F_COLLECT03>.
    assign LT_COLLECT04->* to <F_COLLECT04>.
    assign LT_COLLECT05->* to <F_COLLECT05>.
    assign LT_COLLECT06->* to <F_COLLECT06>.
    assign LT_COLLECT07->* to <F_COLLECT07>.
    assign LT_COLLECT08->* to <F_COLLECT08>.
    assign LT_COLLECT09->* to <F_COLLECT09>.

* transfer to KKBLO struct

    call function 'LVC_TRANSFER_TO_KKBLO'
      exporting
        IT_FIELDCAT_LVC           = LT_FIELDCAT_LVC
        IT_SORT_LVC               = LT_SORT_LVC
        IT_FILTER_INDEX_LVC       = LT_FILTER_IDX_LVC
        IT_GROUPLEVELS_LVC        = LT_GROUPLEVELS_LVC
      importing
        ET_FIELDCAT_KKBLO         = LT_FIELDCAT_KKBLO
        ET_SORT_KKBLO             = LT_SORT_KKBLO
        ET_FILTERED_ENTRIES_KKBLO = LT_FILTER_IDX_KKBLO
        ET_GROUPLEVELS_KKBLO      = LT_GROUPLEVELS_KKBLO
      tables
        IT_DATA                   = <F_ALV_TAB>
      exceptions
        IT_DATA_MISSING           = 1
        IT_FIELDCAT_LVC_MISSING   = 2
        others                    = 3.
    if SY-SUBRC <> 0.
      raise EX_TRANSFER_KKBLO_ERROR.
    endif.

    clear:
      WA_FIELDCAT_LVC,
      LT_FIELDCAT_LVC,
      LT_SORT_LVC,
      LT_FILTER_IDX_LVC,
      LT_GROUPLEVELS_LVC.

    clear:
      LO_GRID.


*--------------------------------------------------------------------*
* SESSION 2: SORT, FILTER AND CALCULATE TOTAL / SUBTOTAL
*--------------------------------------------------------------------*

* append subtotal & total line

    create data LT_EXCEL like <F_ALV_TAB>.
    assign LT_EXCEL->* to <F_EXCEL_TAB>.

    loop at <F_ALV_TAB> assigning <F_LINE>.
      L_SAVE_INDEX = SY-TABIX.

* filter base on filter index table

      read table LT_FILTER_IDX_KKBLO assigning <F_FILTER_IDX_LINE>
      with key INDEX = L_SAVE_INDEX
      binary search.
      if SY-SUBRC ne 0.
        append <F_LINE> to <F_EXCEL_TAB>.
      endif.

* append subtotal lines

      read table LT_GROUPLEVELS_KKBLO assigning <F_GROUPLEVELS_LINE>
      with key INDEX_TO = L_SAVE_INDEX
      binary search.
      if SY-SUBRC = 0.
        L_TABIX = SY-TABIX.
        do.
          if <F_GROUPLEVELS_LINE>-SUBTOT eq 'X' and
             <F_GROUPLEVELS_LINE>-HIDE_LEVEL is initial and
             <F_GROUPLEVELS_LINE>-CINDEX_FROM ne 0.

* dynamic append subtotal line to excel table base on grouplevel table
* ex <f_GROUPLEVELS_line>-level = 1
* then <f_collect_tab> = '<F_COLLECT01>'

            L_COLLECT = <F_GROUPLEVELS_LINE>-LEVEL.
            condense L_COLLECT.
            concatenate '<F_COLLECT0'
                        L_COLLECT '>'
*                      '->*'
                        into L_COLLECT.

            assign (L_COLLECT) to <F_COLLECT_TAB>.

* incase there're more than 1 total line of group, at the same level
* for example: subtotal of multi currency

            loop at <F_COLLECT_TAB> assigning <F_COLLECT_LINE>.
              if  SY-TABIX between <F_GROUPLEVELS_LINE>-CINDEX_FROM
                              and  <F_GROUPLEVELS_LINE>-CINDEX_TO.


                append <F_COLLECT_LINE> to <F_EXCEL_TAB>.

* save subtotal lines index

                WA_SUBTOT_INDEXS-INDEX = SY-TABIX.
                append WA_SUBTOT_INDEXS to LT_SUBTOT_INDEXS.

* append sub total ranges table for format later

                add 1 to L_SUB_INDEX.
                SUBRANGEITEM-NAME     =  L_SUB_INDEX.
                condense SUBRANGEITEM-NAME.
                concatenate 'SUBTOT'
                            SUBRANGEITEM-NAME
                            into SUBRANGEITEM-NAME.

                SUBRANGEITEM-ROWS     = WA_SUBTOT_INDEXS-INDEX.
                SUBRANGEITEM-COLUMNS  = 1.            " start col
                append SUBRANGEITEM to SUBRANGES.
                clear: SUBRANGEITEM.

              endif.
            endloop.
            unassign: <F_COLLECT_TAB>.
            unassign: <F_COLLECT_LINE>.
            clear: L_COLLECT.
          endif.

* check next subtotal level of group

          unassign: <F_GROUPLEVELS_LINE>.
          add 1 to L_TABIX.

          read table LT_GROUPLEVELS_KKBLO assigning <F_GROUPLEVELS_LINE>
          index L_TABIX.
          if SY-SUBRC ne 0
          or <F_GROUPLEVELS_LINE>-INDEX_TO ne L_SAVE_INDEX.
            exit.
          endif.

          unassign:
            <F_COLLECT_TAB>,
            <F_COLLECT_LINE>.

        enddo.
      endif.

      clear:
        L_TABIX,
        L_SAVE_INDEX.

      unassign:
        <F_FILTER_IDX_LINE>,
        <F_GROUPLEVELS_LINE>.

    endloop.

* free local data

    unassign:
      <F_LINE>,
      <F_COLLECT_TAB>,
      <F_COLLECT_LINE>,
      <F_FIELDCAT_LINE>.

* append grand total line

    if <F_COLLECT00> is assigned.
      assign <F_COLLECT00> to <F_COLLECT_TAB>.
      if <F_COLLECT_TAB> is not initial.
        loop at <F_COLLECT_TAB> assigning <F_COLLECT_LINE>.

          append <F_COLLECT_LINE> to <F_EXCEL_TAB>.

* save total line index

          WA_SUBTOT_INDEXS-INDEX = SY-TABIX.
          append WA_SUBTOT_INDEXS to LT_SUBTOT_INDEXS.

* append grand total range (to format)

          add 1 to L_SUB_INDEX.
          SUBRANGEITEM-NAME     =  L_SUB_INDEX.
          condense SUBRANGEITEM-NAME.
          concatenate 'TOTAL'
                      SUBRANGEITEM-NAME
                      into SUBRANGEITEM-NAME.

          SUBRANGEITEM-ROWS     = WA_SUBTOT_INDEXS-INDEX.
          SUBRANGEITEM-COLUMNS  = 1.            " start col
          append SUBRANGEITEM to SUBRANGES.
        endloop.
      endif.
    endif.

    clear:
      SUBRANGEITEM,
      LT_SORT_KKBLO,
      <F_COLLECT00>,
      <F_COLLECT01>,
      <F_COLLECT02>,
      <F_COLLECT03>,
      <F_COLLECT04>,
      <F_COLLECT05>,
      <F_COLLECT06>,
      <F_COLLECT07>,
      <F_COLLECT08>,
      <F_COLLECT09>.

    unassign:
      <F_COLLECT00>,
      <F_COLLECT01>,
      <F_COLLECT02>,
      <F_COLLECT03>,
      <F_COLLECT04>,
      <F_COLLECT05>,
      <F_COLLECT06>,
      <F_COLLECT07>,
      <F_COLLECT08>,
      <F_COLLECT09>,
      <F_COLLECT_TAB>,
      <F_COLLECT_LINE>.

*--------------------------------------------------------------------*
* SESSION 3: MAP DATA TO SEMANTIC TABLE
*--------------------------------------------------------------------*

* get dependent field field: currency and quantity

    create data WA like line of <F_EXCEL_TAB>.
    assign WA->* to <F_EXCEL_LINE>.

    describe field <F_EXCEL_LINE> type L_TYP components L_NUM.

    do L_NUM times.
      L_SAVE_INDEX = SY-INDEX.
      assign component L_SAVE_INDEX of structure <F_EXCEL_LINE>
      to <F_EXCEL_COLUMN>.
      if SY-SUBRC ne 0.
        message E059(0K) with 'FATAL ERROR' raising FATAL_ERROR.
      endif.

      loop at LT_FIELDCAT_KKBLO assigning <F_FIELDCAT_LINE>
      where TABNAME = L_TABNAME.
        assign component <F_FIELDCAT_LINE>-FIELDNAME
        of structure <F_EXCEL_LINE> to <F_FCAT_COLUMN>.

        describe distance between <F_EXCEL_COLUMN> and <F_FCAT_COLUMN>
        into L_INT in byte mode.

* append column index
* this columns index is of table, not fieldcat

        if L_INT = 0.
          WA_COLUMN_INDEX-FIELDNAME = <F_FIELDCAT_LINE>-FIELDNAME.
          WA_COLUMN_INDEX-TABNAME   = <F_FIELDCAT_LINE>-TABNAME.
          WA_COLUMN_INDEX-COL       = L_SAVE_INDEX.
          append WA_COLUMN_INDEX to LT_COLUMN_INDEX.
        endif.

* append dependent fields (currency and quantity unit)

        if <F_FIELDCAT_LINE>-CFIELDNAME is not initial.
          clear WA_FIELDCAT_DEPF.
          WA_FIELDCAT_DEPF-FIELDNAME = <F_FIELDCAT_LINE>-CFIELDNAME.
          WA_FIELDCAT_DEPF-TABNAME   = <F_FIELDCAT_LINE>-CTABNAME.
          collect WA_FIELDCAT_DEPF into LT_FIELDCAT_DEPF.
        endif.

        if <F_FIELDCAT_LINE>-QFIELDNAME is not initial.
          clear WA_FIELDCAT_DEPF.
          WA_FIELDCAT_DEPF-FIELDNAME = <F_FIELDCAT_LINE>-QFIELDNAME.
          WA_FIELDCAT_DEPF-TABNAME   = <F_FIELDCAT_LINE>-QTABNAME.
          collect WA_FIELDCAT_DEPF into LT_FIELDCAT_DEPF.
        endif.

* rewrite field data type

        if <F_FIELDCAT_LINE>-INTTYPE = 'X'
        and <F_FIELDCAT_LINE>-DATATYPE(3) = 'INT'.
          <F_FIELDCAT_LINE>-INTTYPE = 'I'.
        endif.

      endloop.

      clear: L_SAVE_INDEX.
      unassign: <F_FIELDCAT_LINE>.

    enddo.

* build semantic tables

    L_N_HRZ_KEYS = 1.

*   Get keyfigures

    loop at LT_FIELDCAT_KKBLO assigning <F_FIELDCAT_LINE>
    where TABNAME = L_TABNAME
    and TECH ne 'X'
    and NO_OUT ne 'X'.

      clear WA_SEMA.
      clear WA_HKEY.

*   Units belong to keyfigures -> display as str

      read table LT_FIELDCAT_DEPF into WA_FIELDCAT_DEPF with key
      FIELDNAME = <F_FIELDCAT_LINE>-FIELDNAME
      TABNAME   = <F_FIELDCAT_LINE>-TABNAME.

      if SY-SUBRC = 0.
        WA_SEMA-COL_TYP = 'STR'.
        WA_SEMA-COL_OPS = 'DFT'.

*   Keyfigures

      else.
        case <F_FIELDCAT_LINE>-DATATYPE.
          when 'QUAN'.
            WA_SEMA-COL_TYP = 'N03'.

            if <F_FIELDCAT_LINE>-NO_SUM ne 'X'.
              WA_SEMA-COL_OPS = 'ADD'.
            else.
              WA_SEMA-COL_OPS = 'NOP'. " no dependent field
            endif.

          when 'DATS'.
            WA_SEMA-COL_TYP = 'DAT'.
            WA_SEMA-COL_OPS = 'NOP'.

          when 'CHAR' or 'UNIT' or 'CUKY'. " Added fieldformats UNIT and CUKY - dd. 26-10-2012 Wouter Heuvelmans
            WA_SEMA-COL_TYP = 'STR'.
            WA_SEMA-COL_OPS = 'DFT'.   " dependent field

*   incase numeric, ex '00120' -> display as '12'

          when 'NUMC'.
            WA_SEMA-COL_TYP = 'STR'.
            WA_SEMA-COL_OPS = 'DFT'.

          when others.
            WA_SEMA-COL_TYP = 'NUM'.

            if <F_FIELDCAT_LINE>-NO_SUM ne 'X'.
              WA_SEMA-COL_OPS = 'ADD'.
            else.
              WA_SEMA-COL_OPS = 'NOP'.
            endif.
        endcase.
      endif.

      L_COUNTER = L_COUNTER + 1.
      L_N_ATT_COLS = L_N_ATT_COLS + 1.

      WA_SEMA-COL_NO = L_COUNTER.

      read table LT_COLUMN_INDEX into WA_COLUMN_INDEX with key
      FIELDNAME = <F_FIELDCAT_LINE>-FIELDNAME
      TABNAME   = <F_FIELDCAT_LINE>-TABNAME.

      if SY-SUBRC = 0.
        WA_SEMA-COL_SRC = WA_COLUMN_INDEX-COL.
      else.
        raise FATAL_ERROR.
      endif.

* columns index of ref currency field in table

      if not <F_FIELDCAT_LINE>-CFIELDNAME is initial.
        read table LT_COLUMN_INDEX into WA_COLUMN_INDEX with key
        FIELDNAME = <F_FIELDCAT_LINE>-CFIELDNAME
        TABNAME   = <F_FIELDCAT_LINE>-CTABNAME.

        if SY-SUBRC = 0.
          WA_SEMA-COL_CUR = WA_COLUMN_INDEX-COL.
        endif.

* quantities fields
* treat as currency when display on excel

      elseif not <F_FIELDCAT_LINE>-QFIELDNAME is initial.
        read table LT_COLUMN_INDEX into WA_COLUMN_INDEX with key
        FIELDNAME = <F_FIELDCAT_LINE>-QFIELDNAME
        TABNAME   = <F_FIELDCAT_LINE>-QTABNAME.
        if SY-SUBRC = 0.
          WA_SEMA-COL_CUR = WA_COLUMN_INDEX-COL.
        endif.

      endif.

*   Treat of fixed currency in the fieldcatalog for column

      data: L_NUM_HELP(2) type N.

      if not <F_FIELDCAT_LINE>-CURRENCY is initial.

        select * from TCURX into table LT_TCURX.
        sort LT_TCURX.
        read table LT_TCURX into WA_TCURX
                   with key CURRKEY = <F_FIELDCAT_LINE>-CURRENCY.
        if SY-SUBRC = 0.
          L_NUM_HELP = WA_TCURX-CURRDEC.
          concatenate 'N' L_NUM_HELP into WA_SEMA-COL_TYP.
          WA_SEMA-COL_CUR = SY-TABIX * ( -1 ).
        endif.

      endif.

      WA_HKEY-COL_NO    = L_N_ATT_COLS.
      WA_HKEY-ROW_NO    = L_N_HRZ_KEYS.
      WA_HKEY-COL_NAME  = <F_FIELDCAT_LINE>-REPTEXT.
      append WA_HKEY to LT_HKEY.
      append WA_SEMA to LT_SEMA.

    endloop.

* free local data

    clear:
      LT_COLUMN_INDEX,
      WA_COLUMN_INDEX,
      LT_FIELDCAT_DEPF,
      WA_FIELDCAT_DEPF,
      LT_TCURX,
      WA_TCURX,
      L_NUM,
      L_TYP,
      WA,
      L_INT,
      L_COUNTER.

    unassign:
      <F_FIELDCAT_LINE>,
      <F_EXCEL_LINE>,
      <F_EXCEL_COLUMN>,
      <F_FCAT_COLUMN>.

*--------------------------------------------------------------------*
* SESSION 4: WRITE TO EXCEL
*--------------------------------------------------------------------*

    clear: WA_TCURX.
    refresh: LT_TCURX.

*   if spreadsheet dun have proxy yet

    if LI_HAS is initial.
      L_RETCODE = C_OI_ERRORS=>RET_INTERFACE_NOT_SUPPORTED.
      call method C_OI_ERRORS=>CREATE_ERROR_FOR_RETCODE
        exporting
          RETCODE  = L_RETCODE
          NO_FLUSH = NO_FLUSH
        importing
          ERROR    = LO_ERROR_W.
      exit.
    endif.

    create object L_ERROR
      exporting
        OBJECT_NAME = 'OLE_DOCUMENT_PROXY'
        METHOD_NAME = 'get_ranges_names'.

    call method C_OI_ERRORS=>ADD_ERROR
      exporting
        ERROR = L_ERROR.


    describe table LT_SEMA lines DATAREAL.
    describe table <F_EXCEL_TAB> lines DATAC.
    describe table LT_VKEY lines VKEYCOUNT.

    if DATAC = 0.
      raise INV_DATA_RANGE.
    endif.


    if VKEYCOUNT ne L_N_VRT_KEYS.
      raise DIM_MISMATCH_VKEY.
    endif.

    ALL = L_N_VRT_KEYS + L_N_ATT_COLS.

    if DATAREAL ne ALL.
      raise DIM_MISMATCH_SEMA.
    endif.

    data: DECIMAL type C.

* get decimal separator format ('.', ',', ...) in Office config

    call method LO_PROXY->GET_APPLICATION_PROPERTY
      exporting
        PROPERTY_NAME    = 'INTERNATIONAL'
        SUBPROPERTY_NAME = 'DECIMAL_SEPARATOR'
      changing
        RETVALUE         = DECIMAL.

    data: WA_USR type USR01.
    select * from USR01 into WA_USR where BNAME = SY-UNAME.
    endselect.

    data: COMMA_ELIM(4) type C.
    field-symbols <G> type ANY.
    data SEARCH_ITEM(4) value '   #'.

    concatenate ',' DECIMAL '.' DECIMAL into COMMA_ELIM.

    data HELP type I. " table (with subtotal) line number

    HELP = DATAC.

    data: ROWMAX type I value 1.    " header row number
    data: COLUMNMAX type I value 0. " header columns number

    loop at LT_HKEY into HKEYITEM.
      if HKEYITEM-COL_NO > COLUMNMAX.
        COLUMNMAX = HKEYITEM-COL_NO.
      endif.

      if HKEYITEM-ROW_NO > ROWMAX.
        ROWMAX = HKEYITEM-ROW_NO.
      endif.
    endloop.

    data: HKEYCOLUMNS type I. " header columns no

    HKEYCOLUMNS = COLUMNMAX.

    if HKEYCOLUMNS <   L_N_ATT_COLS.
      HKEYCOLUMNS = L_N_ATT_COLS.
    endif.

    COLUMNMAX = 0.

    loop at LT_VKEY into VKEYITEM.
      if VKEYITEM-COL_NO > COLUMNMAX.
        COLUMNMAX = VKEYITEM-COL_NO.
      endif.
    endloop.

    data OVERFLOW type I value 1.
    data TESTNAME(10) type C.
    data TEMP2 type I.                " 1st item row position in excel
    data REALMIT type I value 1.
    data REALOVERFLOW type I value 1. " row index in content

    call method LO_SPREADSHEET->SCREEN_UPDATE
      exporting
        UPDATING = ''.

    call method LO_SPREADSHEET->LOAD_LIB.

    data: STR(40) type C. " range names of columns range (w/o col header)
    data: ROWS type I.    " row postion of 1st item line in ecxel

* calculate row position of data table

    describe table IT_LISTHEADER lines LI_COMMENTARY_ROWS.

* if grid had title, add 1 empy line between title and table

    if LI_COMMENTARY_ROWS ne 0.
      add 1 to LI_COMMENTARY_ROWS.
    endif.

* add top position of block data

    LI_COMMENTARY_ROWS = LI_COMMENTARY_ROWS + I_TOP - 1.

* write header (commentary rows)

    data: LI_COMMENTARY_ROW_INDEX type I value 1.
    data: LI_CONTENT_INDEX type I value 1.
    data: LS_INDEX(10) type C.
    data  LS_COMMENTARY_RANGE(40) type C value 'TITLE'.
    data: LI_FONT_BOLD    type I.
    data: LI_FONT_ITALIC  type I.
    data: LI_FONT_SIZE    type I.

    loop at IT_LISTHEADER into WA_LISTHEADER.
      LI_COMMENTARY_ROW_INDEX = I_TOP + LI_CONTENT_INDEX - 1.
      LS_INDEX = LI_CONTENT_INDEX.
      condense LS_INDEX.
      concatenate LS_COMMENTARY_RANGE(5) LS_INDEX
                  into LS_COMMENTARY_RANGE.
      condense LS_COMMENTARY_RANGE.

* insert title range

      call method LO_SPREADSHEET->INSERT_RANGE_DIM
        exporting
          NAME     = LS_COMMENTARY_RANGE
          TOP      = LI_COMMENTARY_ROW_INDEX
          LEFT     = I_LEFT
          ROWS     = 1
          COLUMNS  = 1
          NO_FLUSH = NO_FLUSH.

* format range

      case WA_LISTHEADER-TYP.
        when 'H'. "title
          LI_FONT_SIZE    = 16.
          LI_FONT_BOLD    = 1.
          LI_FONT_ITALIC  = -1.
        when 'S'. "subtile
          LI_FONT_SIZE = -1.
          LI_FONT_BOLD    = 1.
          LI_FONT_ITALIC  = -1.
        when others. "'A' comment
          LI_FONT_SIZE = -1.
          LI_FONT_BOLD    = -1.
          LI_FONT_ITALIC  = 1.
      endcase.

      call method LO_SPREADSHEET->SET_FONT
        exporting
          RANGENAME = LS_COMMENTARY_RANGE
          FAMILY    = ''
          SIZE      = LI_FONT_SIZE
          BOLD      = LI_FONT_BOLD
          ITALIC    = LI_FONT_ITALIC
          ALIGN     = 0
          NO_FLUSH  = NO_FLUSH.

* title: range content

      RANGEITEM-NAME = LS_COMMENTARY_RANGE.
      RANGEITEM-COLUMNS = 1.
      RANGEITEM-ROWS = 1.
      append RANGEITEM to RANGES.

      CONTENTSITEM-ROW    = LI_CONTENT_INDEX.
      CONTENTSITEM-COLUMN = 1.
      concatenate WA_LISTHEADER-KEY
                  WA_LISTHEADER-INFO
                  into CONTENTSITEM-VALUE
                  separated by SPACE.
      condense CONTENTSITEM-VALUE.
      append CONTENTSITEM to CONTENTS.

      add 1 to LI_CONTENT_INDEX.

      clear:
        RANGEITEM,
        CONTENTSITEM,
        LS_INDEX.

    endloop.

* set range data title

    call method LO_SPREADSHEET->SET_RANGES_DATA
      exporting
        RANGES   = RANGES
        CONTENTS = CONTENTS
        NO_FLUSH = NO_FLUSH.

    refresh:
       RANGES,
       CONTENTS.

    ROWS = ROWMAX + LI_COMMENTARY_ROWS + 1.

    ALL = WA_USR-DATFM.
    ALL = ALL + 3.

    loop at LT_SEMA into SEMAITEM.
      if SEMAITEM-COL_TYP = 'DAT' or SEMAITEM-COL_TYP = 'MON' or
         SEMAITEM-COL_TYP = 'N00' or SEMAITEM-COL_TYP = 'N01' or
         SEMAITEM-COL_TYP = 'N01' or SEMAITEM-COL_TYP = 'N02' or
         SEMAITEM-COL_TYP = 'N03' or SEMAITEM-COL_TYP = 'PCT' or
         SEMAITEM-COL_TYP = 'STR' or SEMAITEM-COL_TYP = 'NUM'.
        clear STR.
        STR = SEMAITEM-COL_NO.
        condense STR.
        concatenate 'DATA' STR into STR.
        MIT = SEMAITEM-COL_NO.
        LI_COL_POS = SEMAITEM-COL_NO + I_LEFT - 1.

* range from data1 to data(n), for each columns of table

        call method LO_SPREADSHEET->INSERT_RANGE_DIM
          exporting
            NAME     = STR
            TOP      = ROWS
            LEFT     = LI_COL_POS
            ROWS     = HELP
            COLUMNS  = 1
            NO_FLUSH = NO_FLUSH.

        data DEC type I value -1.
        data TYPEINFO type SYDES_TYPEINFO.
        loop at <F_EXCEL_TAB> assigning <LINE>.
          assign component SEMAITEM-COL_NO of structure <LINE> to <ITEM>.
          describe field <ITEM> into TD.
          read table TD-TYPES index 1 into TYPEINFO.
          if TYPEINFO-TYPE = 'P'.
            DEC = TYPEINFO-DECIMALS.
          elseif TYPEINFO-TYPE = 'I'.
            DEC = 0.
          endif.

          describe field <LINE> type TYP components COUNT.
          MIT = 1.
          do COUNT times.
            if MIT = SEMAITEM-COL_SRC.
              assign component SY-INDEX of structure <LINE> to <ITEM>.
              describe field <ITEM> into TD.
              read table TD-TYPES index 1 into TYPEINFO.
              if TYPEINFO-TYPE = 'P'.
                DEC = TYPEINFO-DECIMALS.
              endif.
              exit.
            endif.
            MIT = MIT + 1.
          enddo.
          exit.
        endloop.

* format for each columns of table (w/o columns headers)

        if SEMAITEM-COL_TYP = 'DAT'.
          if SEMAITEM-COL_NO > VKEYCOUNT.
            call method LO_SPREADSHEET->SET_FORMAT
              exporting
                RANGENAME = STR
                CURRENCY  = ''
                TYP       = ALL
                NO_FLUSH  = NO_FLUSH.
          else.
            call method LO_SPREADSHEET->SET_FORMAT
              exporting
                RANGENAME = STR
                CURRENCY  = ''
                TYP       = 0
                NO_FLUSH  = NO_FLUSH.
          endif.
        elseif SEMAITEM-COL_TYP = 'STR'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 0
              NO_FLUSH  = NO_FLUSH.
        elseif SEMAITEM-COL_TYP = 'MON'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 10
              NO_FLUSH  = NO_FLUSH.
        elseif SEMAITEM-COL_TYP = 'N00'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 1
              DECIMALS  = 0
              NO_FLUSH  = NO_FLUSH.
        elseif SEMAITEM-COL_TYP = 'N01'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 1
              DECIMALS  = 1
              NO_FLUSH  = NO_FLUSH.
        elseif SEMAITEM-COL_TYP = 'N02'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 1
              DECIMALS  = 2
              NO_FLUSH  = NO_FLUSH.
        elseif SEMAITEM-COL_TYP = 'N03'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 1
              DECIMALS  = 3
              NO_FLUSH  = NO_FLUSH.
        elseif SEMAITEM-COL_TYP = 'N04'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 1
              DECIMALS  = 4
              NO_FLUSH  = NO_FLUSH.
        elseif SEMAITEM-COL_TYP = 'NUM'.
          if DEC eq -1.
            call method LO_SPREADSHEET->SET_FORMAT
              exporting
                RANGENAME = STR
                CURRENCY  = ''
                TYP       = 1
                DECIMALS  = 2
                NO_FLUSH  = NO_FLUSH.
          else.
            call method LO_SPREADSHEET->SET_FORMAT
              exporting
                RANGENAME = STR
                CURRENCY  = ''
                TYP       = 1
                DECIMALS  = DEC
                NO_FLUSH  = NO_FLUSH.
          endif.
        elseif SEMAITEM-COL_TYP = 'PCT'.
          call method LO_SPREADSHEET->SET_FORMAT
            exporting
              RANGENAME = STR
              CURRENCY  = ''
              TYP       = 3
              DECIMALS  = 0
              NO_FLUSH  = NO_FLUSH.
        endif.

      endif.
    endloop.

* get item contents for set_range_data method
* get currency cell also

    MIT = 1.

    data: CURRCELLS type SOI_CELL_TABLE.
    data: CURRITEM  type SOI_CELL_ITEM.

    CURRITEM-ROWS = 1.
    CURRITEM-COLUMNS = 1.
    CURRITEM-FRONT = -1.
    CURRITEM-BACK = -1.
    CURRITEM-FONT = ''.
    CURRITEM-SIZE = -1.
    CURRITEM-BOLD = -1.
    CURRITEM-ITALIC = -1.
    CURRITEM-ALIGN = -1.
    CURRITEM-FRAMETYP = -1.
    CURRITEM-FRAMECOLOR = -1.
    CURRITEM-CURRENCY = ''.
    CURRITEM-NUMBER = 1.
    CURRITEM-INPUT = -1.

    data: CONST type I.

*   Change for Correction request
*    Initial 10000 lines are missing in Excel Export
*    if there are only 2 columns in exported List object.

    if DATAREAL gt 2.
      CONST = 20000 / DATAREAL.
    else.
      CONST = 20000 / ( DATAREAL + 2 ).
    endif.

    data: LINES type I.
    data: INNERLINES type I.
    data: COUNTER type I.
    data: CURRITEM2 like CURRITEM.
    data: CURRITEM3 like CURRITEM.
    data: LENGTH type I.
    data: FOUND.

* append content table (for method set_range_content)

    loop at <F_EXCEL_TAB> assigning <LINE>.

* save line index to compare with lt_subtot_indexs,
* to discover line is a subtotal / totale line or not
* ex use to set 'dun display zero in subtotal / total line'

      L_SAVE_INDEX = SY-TABIX.

      do DATAREAL times.
        read table LT_SEMA into SEMAITEM with key COL_NO = SY-INDEX.
        if SEMAITEM-COL_SRC ne 0.
          assign component SEMAITEM-COL_SRC
                 of structure <LINE> to <ITEM>.
        else.
          assign component SY-INDEX
                 of structure <LINE> to <ITEM>.
        endif.

        CONTENTSITEM-ROW = REALOVERFLOW.

        if SY-SUBRC = 0.
          move SEMAITEM-COL_OPS to SEARCH_ITEM(3).
          search 'ADD#CNT#MIN#MAX#AVG#NOP#DFT#'
                            for SEARCH_ITEM.
          if SY-SUBRC ne 0.
            raise ERROR_IN_SEMA.
          endif.
          move SEMAITEM-COL_TYP to SEARCH_ITEM(3).
          search 'NUM#N00#N01#N02#N03#N04#PCT#DAT#MON#STR#'
                            for SEARCH_ITEM.
          if SY-SUBRC ne 0.
            raise ERROR_IN_SEMA.
          endif.
          CONTENTSITEM-COLUMN = SY-INDEX.
          if SEMAITEM-COL_TYP eq 'DAT' or SEMAITEM-COL_TYP eq 'MON'.
            if SEMAITEM-COL_NO > VKEYCOUNT.

              " Hinweis 512418
              " EXCEL bezieht Datumsangaben
              " auf den 31.12.1899, behandelt
              " aber 1900 als ein Schaltjahr
              " d.h. ab 1.3.1900 korrekt
              " 1.3.1900 als Zahl = 61

              data: GENESIS type D value '18991230'.
              data: NUMBER_OF_DAYS type P.
* change for date in char format & sema_type = X
              data: TEMP_DATE type D.

              if not <ITEM> is initial and not <ITEM> co ' ' and not
              <ITEM> co '0'.
* change for date in char format & sema_type = X starts
                if SEMA_TYPE = 'X'.
                  describe field <ITEM> type TYP.
                  if TYP = 'C'.
                    TEMP_DATE = <ITEM>.
                    NUMBER_OF_DAYS = TEMP_DATE - GENESIS.
                  else.
                    NUMBER_OF_DAYS = <ITEM> - GENESIS.
                  endif.
                else.
                  NUMBER_OF_DAYS = <ITEM> - GENESIS.
                endif.
* change for date in char format & sema_type = X ends
                if NUMBER_OF_DAYS < 61.
                  NUMBER_OF_DAYS = NUMBER_OF_DAYS - 1.
                endif.

                set country 'DE'.
                write NUMBER_OF_DAYS to CONTENTSITEM-VALUE
                no-grouping
                                          left-justified.
                set country SPACE.
                translate CONTENTSITEM-VALUE using COMMA_ELIM.
              else.
                clear CONTENTSITEM-VALUE.
              endif.
            else.
              move <ITEM> to CONTENTSITEM-VALUE.
            endif.
          elseif SEMAITEM-COL_TYP eq 'NUM' or
                 SEMAITEM-COL_TYP eq 'N00' or
                 SEMAITEM-COL_TYP eq 'N01' or
                 SEMAITEM-COL_TYP eq 'N02' or
                 SEMAITEM-COL_TYP eq 'N03' or
                 SEMAITEM-COL_TYP eq 'N04' or
                 SEMAITEM-COL_TYP eq 'PCT'.
            set country 'DE'.
            describe field <ITEM> type TYP.

            if SEMAITEM-COL_CUR is initial.
              if TYP ne 'F'.
                write <ITEM> to CONTENTSITEM-VALUE no-grouping
                                                   no-sign decimals 14.
              else.
                write <ITEM> to CONTENTSITEM-VALUE no-grouping
                                                   no-sign.
              endif.
            else.
* Treat of fixed curreny for column >>Y9CK007319
              if SEMAITEM-COL_CUR < 0.
                SEMAITEM-COL_CUR = SEMAITEM-COL_CUR * ( -1 ).
                select * from TCURX into table LT_TCURX.
                sort LT_TCURX.
                read table LT_TCURX into
                                    WA_TCURX index SEMAITEM-COL_CUR.
                if SY-SUBRC = 0.
                  if TYP ne 'F'.
                    write <ITEM> to CONTENTSITEM-VALUE no-grouping
                     currency WA_TCURX-CURRKEY no-sign decimals 14.
                  else.
                    write <ITEM> to CONTENTSITEM-VALUE no-grouping
                     currency WA_TCURX-CURRKEY no-sign.
                  endif.
                endif.
              else.
                assign component SEMAITEM-COL_CUR
                     of structure <LINE> to <G>.
* mit = index of recent row
                CURRITEM-TOP  = ROWMAX + MIT + LI_COMMENTARY_ROWS.

                LI_COL_POS =  SY-INDEX + I_LEFT - 1.
                CURRITEM-LEFT = LI_COL_POS.

* if filed is quantity field (qfieldname ne space)
* or amount field (cfieldname ne space), then format decimal place
* corresponding with config

                clear: L_DEF.
                read table LT_FIELDCAT_KKBLO assigning <F_FIELDCAT_LINE>
                with key  TABNAME = L_TABNAME
                          TECH    = SPACE
                          NO_OUT  = SPACE
                          COL_POS = SEMAITEM-COL_NO.
                if SY-SUBRC = 0.
                  if <F_FIELDCAT_LINE>-CFIELDNAME is not initial.
                    L_DEF = 'C'.
                  else."if <f_fieldcat_line>-qfieldname is not initial.
                    L_DEF = 'Q'.
                  endif.
                endif.

* if field is amount field
* exporting of amount field base on currency decimal table: TCURX
                if L_DEF = 'C'. "field is amount field
                  select single * from TCURX into WA_TCURX
                    where CURRKEY = <G>.
* if amount ref to un-know currency -> default decimal  = 2
                  if SY-SUBRC eq 0.
                    CURRITEM-DECIMALS = WA_TCURX-CURRDEC.
                  else.
                    CURRITEM-DECIMALS = 2.
                  endif.

                  append CURRITEM to CURRCELLS.
                  if TYP ne 'F'.
                    write <ITEM> to CONTENTSITEM-VALUE
                                        currency <G>
                       no-sign no-grouping.
                  else.
                    write <ITEM> to CONTENTSITEM-VALUE
                       decimals 14      currency <G>
                       no-sign no-grouping.
                  endif.

* if field is quantity field
* exporting of quantity field base on quantity decimal table: T006

                else."if l_def = 'Q'. " field is quantity field
                  clear: WA_T006.
                  select single * from T006 into WA_T006
                    where MSEHI = <G>.
* if quantity ref to un-know unit-> default decimal  = 2
                  if SY-SUBRC eq 0.
                    CURRITEM-DECIMALS = WA_T006-DECAN.
                  else.
                    CURRITEM-DECIMALS = 2.
                  endif.
                  append CURRITEM to CURRCELLS.

                  write <ITEM> to CONTENTSITEM-VALUE
                                      unit <G>
                     no-sign no-grouping.
                  condense CONTENTSITEM-VALUE.

                endif.

              endif.                                        "Y9CK007319
            endif.
            condense CONTENTSITEM-VALUE.

* add function fieldcat-no zero display

            loop at LT_FIELDCAT_KKBLO assigning <F_FIELDCAT_LINE>
            where TABNAME = L_TABNAME
            and   TECH ne 'X'
            and   NO_OUT ne 'X'.
              if <F_FIELDCAT_LINE>-COL_POS = SEMAITEM-COL_NO.
                if <F_FIELDCAT_LINE>-NO_ZERO = 'X'.
                  if <ITEM> = '0'.
                    clear: CONTENTSITEM-VALUE.
                  endif.

* dun display zero in total/subtotal line too

                else.
                  clear: WA_SUBTOT_INDEXS.
                  read table LT_SUBTOT_INDEXS into WA_SUBTOT_INDEXS
                  with key INDEX = L_SAVE_INDEX.
                  if SY-SUBRC = 0 and <ITEM> = '0'.
                    clear: CONTENTSITEM-VALUE.
                  endif.
                endif.
              endif.
            endloop.
            unassign: <F_FIELDCAT_LINE>.

            if <ITEM> lt 0.
              search CONTENTSITEM-VALUE for 'E'.
              if SY-FDPOS eq 0.

* use prefix notation for signed numbers

                translate CONTENTSITEM-VALUE using '- '.
                condense CONTENTSITEM-VALUE no-gaps.
                concatenate '-' CONTENTSITEM-VALUE
                           into CONTENTSITEM-VALUE.
              else.
                concatenate '-' CONTENTSITEM-VALUE
                           into CONTENTSITEM-VALUE.
              endif.
            endif.
            set country SPACE.
* Hier wird nur die korrekte Kommaseparatierung gemacht, wenn die
* Zeichen einer
* Zahl enthalten sind. Das ist für Timestamps, die auch ":" enthalten.
* Für die
* darf keine Kommaseparierung stattfinden.
* Changing for correction request - Y6BK041073
            if CONTENTSITEM-VALUE co '0123456789.,-+E '.
              translate CONTENTSITEM-VALUE using COMMA_ELIM.
            endif.
          else.
            clear CONTENTSITEM-VALUE.

* if type is not numeric -> dun display with zero

            write <ITEM> to CONTENTSITEM-VALUE no-zero.

            shift CONTENTSITEM-VALUE left deleting leading SPACE.

          endif.
          append CONTENTSITEM to CONTENTS.
        endif.
      enddo.

      REALMIT = REALMIT + 1.
      REALOVERFLOW = REALOVERFLOW + 1.

      MIT = MIT + 1.
*   overflow = current row index in content table
      OVERFLOW = OVERFLOW + 1.
    endloop.

    unassign: <F_FIELDCAT_LINE>.

* set item range for set_range_data method

    TESTNAME = MIT / CONST.
    condense TESTNAME.

    concatenate 'TEST' TESTNAME into TESTNAME.

    REALOVERFLOW = REALOVERFLOW - 1.
    REALMIT = REALMIT - 1.
    HELP = REALOVERFLOW.

    RANGEITEM-NAME = TESTNAME.
    RANGEITEM-COLUMNS = DATAREAL.
    RANGEITEM-ROWS = HELP.
    append RANGEITEM to RANGES.

* insert item range dim

    TEMP2 = ROWMAX + 1 + LI_COMMENTARY_ROWS + REALMIT - REALOVERFLOW.

* items data

    call method LO_SPREADSHEET->INSERT_RANGE_DIM
      exporting
        NAME     = TESTNAME
        TOP      = TEMP2
        LEFT     = I_LEFT
        ROWS     = HELP
        COLUMNS  = DATAREAL
        NO_FLUSH = NO_FLUSH.

* get columns header contents for set_range_data method
* export columns header only if no columns header option = space

    data: ROWCOUNT type I.
    data: COLUMNCOUNT type I.

    if I_COLUMNS_HEADER = 'X'.

* append columns header to contents: hkey

      ROWCOUNT = 1.
      do ROWMAX times.
        COLUMNCOUNT = 1.
        do HKEYCOLUMNS times.
          loop at LT_HKEY into HKEYITEM where COL_NO = COLUMNCOUNT
                                           and ROW_NO   = ROWCOUNT.
          endloop.
          if SY-SUBRC = 0.
            STR = HKEYITEM-COL_NAME.
            CONTENTSITEM-VALUE = HKEYITEM-COL_NAME.
          else.
            CONTENTSITEM-VALUE = STR.
          endif.
          CONTENTSITEM-COLUMN = COLUMNCOUNT.
          CONTENTSITEM-ROW = ROWCOUNT.
          append CONTENTSITEM to CONTENTS.
          COLUMNCOUNT = COLUMNCOUNT + 1.
        enddo.
        ROWCOUNT = ROWCOUNT + 1.
      enddo.

* incase columns header in multiline

      data: ROWMAXTEMP type I.
      if ROWMAX > 1.
        ROWMAXTEMP = ROWMAX - 1.
        ROWCOUNT = 1.
        do ROWMAXTEMP times.
          COLUMNCOUNT = 1.
          do COLUMNMAX times.
            CONTENTSITEM-COLUMN = COLUMNCOUNT.
            CONTENTSITEM-ROW    = ROWCOUNT.
            CONTENTSITEM-VALUE  = ''.
            append CONTENTSITEM to CONTENTS.
            COLUMNCOUNT = COLUMNCOUNT + 1.
          enddo.
          ROWCOUNT = ROWCOUNT + 1.
        enddo.
      endif.

* append columns header to contents: vkey

      COLUMNCOUNT = 1.
      do COLUMNMAX times.
        loop at LT_VKEY into VKEYITEM where COL_NO = COLUMNCOUNT.
        endloop.
        CONTENTSITEM-VALUE = VKEYITEM-COL_NAME.
        CONTENTSITEM-ROW = ROWMAX.
        CONTENTSITEM-COLUMN = COLUMNCOUNT.
        append CONTENTSITEM to CONTENTS.
        COLUMNCOUNT = COLUMNCOUNT + 1.
      enddo.
*--------------------------------------------------------------------*
* set header range for method set_range_data
* insert header keys range dim

      LI_HEAD_TOP = LI_COMMENTARY_ROWS + 1.
      LI_COL_POS = I_LEFT.

* insert range headers

      if HKEYCOLUMNS ne 0.
        RANGEITEM-NAME = 'TESTHKEY'.
        RANGEITEM-ROWS = ROWMAX.
        RANGEITEM-COLUMNS = HKEYCOLUMNS.
        append RANGEITEM to RANGES.
        clear: RANGEITEM.

        call method LO_SPREADSHEET->INSERT_RANGE_DIM
          exporting
            NAME     = 'TESTHKEY'
            TOP      = LI_HEAD_TOP
            LEFT     = LI_COL_POS
            ROWS     = ROWMAX
            COLUMNS  = HKEYCOLUMNS
            NO_FLUSH = NO_FLUSH.
      endif.
    endif.

* format for columns header + total + subtotal
* ------------------------------------------

    HELP = ROWMAX + REALMIT. " table + header lines

    data: LT_FORMAT     type SOI_FORMAT_TABLE.
    data: WA_FORMAT     like line of LT_FORMAT.
    data: WA_FORMAT_TEMP like line of LT_FORMAT.

    field-symbols: <F_SOURCE> type ANY.
    field-symbols: <F_DES>    type ANY.

* columns header format

    WA_FORMAT-FRONT       = -1.
    WA_FORMAT-BACK        = 15. "grey
    WA_FORMAT-FONT        = SPACE.
    WA_FORMAT-SIZE        = -1.
    WA_FORMAT-BOLD        = 1.
    WA_FORMAT-ALIGN       = 0.
    WA_FORMAT-FRAMETYP    = -1.
    WA_FORMAT-FRAMECOLOR  = -1.

* get column header format from input record
* -> map input format

    if I_COLUMNS_HEADER = 'X'.
      WA_FORMAT-NAME        = 'TESTHKEY'.
      if I_FORMAT_COL_HEADER is not initial.
        describe field I_FORMAT_COL_HEADER type L_TYP components
        LI_COL_NUM.
        do LI_COL_NUM times.
          if SY-INDEX ne 1. " dun map range name
            assign component SY-INDEX of structure I_FORMAT_COL_HEADER
            to <F_SOURCE>.
            if <F_SOURCE> is not initial.
              assign component SY-INDEX of structure WA_FORMAT to <F_DES>.
              <F_DES> = <F_SOURCE>.
              unassign: <F_DES>.
            endif.
            unassign: <F_SOURCE>.
          endif.
        enddo.

        clear: LI_COL_NUM.
      endif.

      append WA_FORMAT to LT_FORMAT.
    endif.

* Zusammenfassen der Spalten mit gleicher Nachkommastellenzahl
* collect vertical cells (col)  with the same number of decimal places
* to increase perfomance in currency cell format

    describe table CURRCELLS lines LINES.
    LINES = LINES - 1.
    do LINES times.
      describe table CURRCELLS lines INNERLINES.
      INNERLINES = INNERLINES - 1.
      sort CURRCELLS by LEFT TOP.
      clear FOUND.
      do INNERLINES times.
        read table CURRCELLS index SY-INDEX into CURRITEM.
        COUNTER = SY-INDEX + 1.
        read table CURRCELLS index COUNTER into CURRITEM2.
        if CURRITEM-LEFT eq CURRITEM2-LEFT.
          LENGTH = CURRITEM-TOP + CURRITEM-ROWS.
          if LENGTH eq CURRITEM2-TOP and CURRITEM-DECIMALS eq CURRITEM2-DECIMALS.
            move CURRITEM to CURRITEM3.
            CURRITEM3-ROWS = CURRITEM3-ROWS + CURRITEM2-ROWS.
            CURRITEM-LEFT = -1.
            modify CURRCELLS index SY-INDEX from CURRITEM.
            CURRITEM2-LEFT = -1.
            modify CURRCELLS index COUNTER from CURRITEM2.
            append CURRITEM3 to CURRCELLS.
            FOUND = 'X'.
          endif.
        endif.
      enddo.
      if FOUND is initial.
        exit.
      endif.
      delete CURRCELLS where LEFT = -1.
    enddo.

* Zusammenfassen der Zeilen mit gleicher Nachkommastellenzahl
* collect horizontal cells (row) with the same number of decimal places
* to increase perfomance in currency cell format

    describe table CURRCELLS lines LINES.
    LINES = LINES - 1.
    do LINES times.
      describe table CURRCELLS lines INNERLINES.
      INNERLINES = INNERLINES - 1.
      sort CURRCELLS by TOP LEFT.
      clear FOUND.
      do INNERLINES times.
        read table CURRCELLS index SY-INDEX into CURRITEM.
        COUNTER = SY-INDEX + 1.
        read table CURRCELLS index COUNTER into CURRITEM2.
        if CURRITEM-TOP eq CURRITEM2-TOP and CURRITEM-ROWS eq
        CURRITEM2-ROWS.
          LENGTH = CURRITEM-LEFT + CURRITEM-COLUMNS.
          if LENGTH eq CURRITEM2-LEFT and CURRITEM-DECIMALS eq CURRITEM2-DECIMALS.
            move CURRITEM to CURRITEM3.
            CURRITEM3-COLUMNS = CURRITEM3-COLUMNS + CURRITEM2-COLUMNS.
            CURRITEM-LEFT = -1.
            modify CURRCELLS index SY-INDEX from CURRITEM.
            CURRITEM2-LEFT = -1.
            modify CURRCELLS index COUNTER from CURRITEM2.
            append CURRITEM3 to CURRCELLS.
            FOUND = 'X'.
          endif.
        endif.
      enddo.
      if FOUND is initial.
        exit.
      endif.
      delete CURRCELLS where LEFT = -1.
    enddo.
* Ende der Zusammenfassung


* item data: format for currency cell, corresponding with currency

    call method LO_SPREADSHEET->CELL_FORMAT
      exporting
        CELLS    = CURRCELLS
        NO_FLUSH = NO_FLUSH.

* item data: write item table content

    call method LO_SPREADSHEET->SET_RANGES_DATA
      exporting
        RANGES   = RANGES
        CONTENTS = CONTENTS
        NO_FLUSH = NO_FLUSH.

* whole table range to format all table

    if I_COLUMNS_HEADER = 'X'.
      LI_HEAD_TOP = LI_COMMENTARY_ROWS + 1.
    else.
      LI_HEAD_TOP = LI_COMMENTARY_ROWS + 2.
      HELP = HELP - 1.
    endif.

    call method LO_SPREADSHEET->INSERT_RANGE_DIM
      exporting
        NAME     = 'WHOLE_TABLE'
        TOP      = LI_HEAD_TOP
        LEFT     = I_LEFT
        ROWS     = HELP
        COLUMNS  = DATAREAL
        NO_FLUSH = NO_FLUSH.

* columns width auto fix
* this parameter = space in case use with exist template

    if I_COLUMNS_AUTOFIT = 'X'.
      call method LO_SPREADSHEET->FIT_WIDEST
        exporting
          NAME     = 'WHOLE_TABLE'
          NO_FLUSH = NO_FLUSH.
    endif.

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

    call method LO_SPREADSHEET->SET_FRAME
      exporting
        RANGENAME = 'WHOLE_TABLE'
        TYP       = 127
        COLOR     = 1
        NO_FLUSH  = SPACE
      importing
        ERROR     = LO_ERROR
        RETCODE   = LC_RETCODE.

    ERROR_DOI.

* reformat subtotal / total line after format wholw table

    loop at SUBRANGES into SUBRANGEITEM.
      L_SUB_INDEX = SUBRANGEITEM-ROWS + LI_COMMENTARY_ROWS + ROWMAX.

      call method LO_SPREADSHEET->INSERT_RANGE_DIM
        exporting
          NAME     = SUBRANGEITEM-NAME
          LEFT     = I_LEFT
          TOP      = L_SUB_INDEX
          ROWS     = 1
          COLUMNS  = DATAREAL
          NO_FLUSH = NO_FLUSH.

      WA_FORMAT-NAME    = SUBRANGEITEM-NAME.

*   default format:
*     - clolor: subtotal = light yellow, subtotal = yellow
*     - frame: box

      if  SUBRANGEITEM-NAME(3) = 'SUB'.
        WA_FORMAT-BACK = 36. "subtotal line
        WA_FORMAT_TEMP = I_FORMAT_SUBTOTAL.
      else.
        WA_FORMAT-BACK = 27. "total line
        WA_FORMAT_TEMP = I_FORMAT_TOTAL.
      endif.
      WA_FORMAT-FRAMETYP = 79.
      WA_FORMAT-FRAMECOLOR = 1.
      WA_FORMAT-NUMBER  = -1.
      WA_FORMAT-ALIGN   = -1.

*   get subtoal + total format from intput parameter
*   overwrite default format

      if WA_FORMAT_TEMP is not initial.
        describe field WA_FORMAT_TEMP type L_TYP components LI_COL_NUM.
        do LI_COL_NUM times.
          if SY-INDEX ne 1. " dun map range name
            assign component SY-INDEX of structure WA_FORMAT_TEMP
            to <F_SOURCE>.
            if <F_SOURCE> is not initial.
              assign component SY-INDEX of structure WA_FORMAT to <F_DES>.
              <F_DES> = <F_SOURCE>.
              unassign: <F_DES>.
            endif.
            unassign: <F_SOURCE>.
          endif.
        enddo.

        clear: LI_COL_NUM.
      endif.

      append WA_FORMAT to LT_FORMAT.
      clear: WA_FORMAT-NAME.
      clear: L_SUB_INDEX.
      clear: WA_FORMAT_TEMP.

    endloop.

    if LT_FORMAT[] is not initial.
      call method LO_SPREADSHEET->SET_RANGES_FORMAT
        exporting
          FORMATTABLE = LT_FORMAT
          NO_FLUSH    = NO_FLUSH.
      refresh: LT_FORMAT.
    endif.
*--------------------------------------------------------------------*
    call method LO_SPREADSHEET->SCREEN_UPDATE
      exporting
        UPDATING = 'X'.

    call method C_OI_ERRORS=>FLUSH_ERRORS.

    LO_ERROR_W = L_ERROR.
    LC_RETCODE = LO_ERROR_W->ERROR_CODE.

** catch no_flush -> led to dump ( optional )
*    go_error = l_error.
*    gc_retcode = go_error->error_code.
*    error_doi.

    clear:
      LT_SEMA,
      WA_SEMA,
      LT_HKEY,
      WA_HKEY,
      LT_VKEY,
      WA_VKEY,
      L_N_HRZ_KEYS,
      L_N_ATT_COLS,
      L_N_VRT_KEYS,
      COUNT,
      DATAC,
      DATAREAL,
      VKEYCOUNT,
      ALL,
      MIT,
      LI_COL_POS,
      LI_COL_NUM,
      RANGES,
      RANGEITEM,
      CONTENTS,
      CONTENTSITEM,
      SEMAITEM,
      HKEYITEM,
      VKEYITEM,
      LI_COMMENTARY_ROWS,
      L_RETCODE,
      LI_HEAD_TOP,
      <F_EXCEL_TAB>.

    clear:
       LO_ERROR_W.

    unassign:
    <LINE>,
    <ITEM>,
    <F_EXCEL_TAB>.

*--------------------------------------------------------------------*
* SESSION 5: SAVE AND CLOSE FILE
*--------------------------------------------------------------------*

* ex of save path: 'FILE://C:\temp\test.xlsx'
    concatenate 'FILE://' I_SAVE_PATH
                into LS_PATH.

    call method LO_PROXY->SAVE_DOCUMENT_TO_URL
      exporting
        NO_FLUSH      = 'X'
        URL           = LS_PATH
      importing
        ERROR         = LO_ERROR
        RETCODE       = LC_RETCODE
      changing
        DOCUMENT_SIZE = LI_DOCUMENT_SIZE.

    ERROR_DOI.

* if save successfully -> raise successful message
*  message i499(sy) with 'Document is Exported to ' p_path.
    message I499(SY) with 'Data has been exported successfully'.

    clear:
      LS_PATH,
      LI_DOCUMENT_SIZE.

    CLOSE_DOCUMENT.
  endmethod.


  method BIND_TABLE.
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

    constants:
      LC_TOP_LEFT_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA value 'A',
      LC_TOP_LEFT_ROW    type ZEXCEL_CELL_ROW value 1.

    data:
      LV_ROW_INT            type ZEXCEL_CELL_ROW,
      LV_FIRST_ROW          type ZEXCEL_CELL_ROW,
      LV_LAST_ROW           type ZEXCEL_CELL_ROW,
      LV_COLUMN_INT         type ZEXCEL_CELL_COLUMN,
      LV_COLUMN_ALPHA       type ZEXCEL_CELL_COLUMN_ALPHA,
      LT_FIELD_CATALOG      type ZEXCEL_T_FIELDCATALOG,
      LV_ID                 type I,
      LV_ROWS               type I,
      LV_FORMULA            type STRING,
      LS_SETTINGS           type ZEXCEL_S_TABLE_SETTINGS,
      LO_TABLE              type ref to ZCL_EXCEL_TABLE,
      LT_COLUMN_NAME_BUFFER type sorted table of STRING with unique key TABLE_LINE,
      LV_VALUE              type STRING,
      LV_VALUE_LOWERCASE    type STRING,
      LV_SYINDEX            type CHAR3,
      LV_ERRORMESSAGE       type STRING,                                            "ins issue #237

      LV_COLUMNS            type I,
      LT_COLUMNS            type ZEXCEL_T_FIELDCATALOG,
      LV_MAXCOL             type I,
      LV_MAXROW             type I,
      LO_ITERATOR           type ref to CL_OBJECT_COLLECTION_ITERATOR,
      LO_STYLE_COND         type ref to ZCL_EXCEL_STYLE_COND,
      LO_CURTABLE           type ref to ZCL_EXCEL_TABLE.

    field-symbols:
      <LS_FIELD_CATALOG>        type ZEXCEL_S_FIELDCATALOG,
      <LS_FIELD_CATALOG_CUSTOM> type ZEXCEL_S_FIELDCATALOG,
      <FS_TABLE_LINE>           type ANY,
      <FS_FLDVAL>               type ANY.

    LS_SETTINGS = IS_TABLE_SETTINGS.

    if LS_SETTINGS-TOP_LEFT_COLUMN is initial.
      LS_SETTINGS-TOP_LEFT_COLUMN = LC_TOP_LEFT_COLUMN.
    endif.

    if LS_SETTINGS-TABLE_STYLE is initial.
      LS_SETTINGS-TABLE_STYLE = ZCL_EXCEL_TABLE=>BUILTINSTYLE_MEDIUM2.
    endif.

    if LS_SETTINGS-TOP_LEFT_ROW is initial.
      LS_SETTINGS-TOP_LEFT_ROW = LC_TOP_LEFT_ROW.
    endif.

    if IT_FIELD_CATALOG is not supplied.
      LT_FIELD_CATALOG = ZCL_EXCEL_COMMON=>GET_FIELDCATALOG( IP_TABLE = IP_TABLE ).
    else.
      LT_FIELD_CATALOG = IT_FIELD_CATALOG.
    endif.

    sort LT_FIELD_CATALOG by POSITION.

*--------------------------------------------------------------------*
*  issue #237   Check if overlapping areas exist  Start
*--------------------------------------------------------------------*
    "Get the number of columns for the current table
    LT_COLUMNS = LT_FIELD_CATALOG.
    delete LT_COLUMNS where DYNPFLD ne ABAP_TRUE.
    describe table LT_COLUMNS lines LV_COLUMNS.

    "Calculate the top left row of the current table
    LV_COLUMN_INT = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( LS_SETTINGS-TOP_LEFT_COLUMN ).
    LV_ROW_INT    = LS_SETTINGS-TOP_LEFT_ROW.

    "Get number of row for the current table
    describe table IP_TABLE lines LV_ROWS.

    "Calculate the bottom right row for the current table
    LV_MAXCOL                       = LV_COLUMN_INT + LV_COLUMNS - 1.
    LV_MAXROW                       = LV_ROW_INT    + LV_ROWS - 1.
    LS_SETTINGS-BOTTOM_RIGHT_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( LV_MAXCOL ).
    LS_SETTINGS-BOTTOM_RIGHT_ROW    = LV_MAXROW.

    LV_COLUMN_INT                   = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( LS_SETTINGS-TOP_LEFT_COLUMN ).

    LO_ITERATOR = ME->TABLES->IF_OBJECT_COLLECTION~GET_ITERATOR( ).
    while LO_ITERATOR->IF_OBJECT_COLLECTION_ITERATOR~HAS_NEXT( ) eq ABAP_TRUE.

      LO_CURTABLE ?= LO_ITERATOR->IF_OBJECT_COLLECTION_ITERATOR~GET_NEXT( ).
      if  (    (  LS_SETTINGS-TOP_LEFT_ROW     ge LO_CURTABLE->SETTINGS-TOP_LEFT_ROW                             and LS_SETTINGS-TOP_LEFT_ROW     le LO_CURTABLE->SETTINGS-BOTTOM_RIGHT_ROW )
            or
               (  LS_SETTINGS-BOTTOM_RIGHT_ROW ge LO_CURTABLE->SETTINGS-TOP_LEFT_ROW                             and LS_SETTINGS-BOTTOM_RIGHT_ROW le LO_CURTABLE->SETTINGS-BOTTOM_RIGHT_ROW )
          )
        and
          (    (  LV_COLUMN_INT ge ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( LO_CURTABLE->SETTINGS-TOP_LEFT_COLUMN ) and LV_COLUMN_INT le ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( LO_CURTABLE->SETTINGS-BOTTOM_RIGHT_COLUMN ) )
            or
               (  LV_MAXCOL     ge ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( LO_CURTABLE->SETTINGS-TOP_LEFT_COLUMN ) and LV_MAXCOL     le ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( LO_CURTABLE->SETTINGS-BOTTOM_RIGHT_COLUMN ) )
          ).
        LV_ERRORMESSAGE = 'Table overlaps with previously bound table and will not be added to worksheet.'(400).
        zcx_excel=>raise_text( lv_errormessage ).
      endif.

    endwhile.
*--------------------------------------------------------------------*
*  issue #237   Check if overlapping areas exist  End
*--------------------------------------------------------------------*

    create object LO_TABLE.
    LO_TABLE->SETTINGS = LS_SETTINGS.
    LO_TABLE->SET_DATA( IR_DATA = IP_TABLE ).
    LV_ID = ME->EXCEL->GET_NEXT_TABLE_ID( ).
    LO_TABLE->SET_ID( IV_ID = LV_ID ).
*  lo_table->fieldcat = lt_field_catalog[].

    ME->TABLES->ADD( LO_TABLE ).

* It is better to loop column by column (only visible column)
    loop at LT_FIELD_CATALOG assigning <LS_FIELD_CATALOG> where DYNPFLD eq ABAP_TRUE.

      LV_COLUMN_ALPHA = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( LV_COLUMN_INT ).

      " Due restrinction of new table object we cannot have two column with the same name
      " Check if a column with the same name exists, if exists add a counter
      " If no medium description is provided we try to use small or long
*    lv_value = <ls_field_catalog>-scrtext_m.
      field-symbols: <SCRTXT1> type ANY,
                     <SCRTXT2> type ANY,
                     <SCRTXT3> type ANY.

      case IV_DEFAULT_DESCR.
        when 'M'.
          assign <LS_FIELD_CATALOG>-SCRTEXT_M to <SCRTXT1>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_S to <SCRTXT2>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_L to <SCRTXT3>.
        when 'S'.
          assign <LS_FIELD_CATALOG>-SCRTEXT_S to <SCRTXT1>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_M to <SCRTXT2>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_L to <SCRTXT3>.
        when 'L'.
          assign <LS_FIELD_CATALOG>-SCRTEXT_L to <SCRTXT1>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_M to <SCRTXT2>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_S to <SCRTXT3>.
        when others.
          assign <LS_FIELD_CATALOG>-SCRTEXT_M to <SCRTXT1>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_S to <SCRTXT2>.
          assign <LS_FIELD_CATALOG>-SCRTEXT_L to <SCRTXT3>.
      endcase.


      if <SCRTXT1> is not initial.
        LV_VALUE = <SCRTXT1>.
        <LS_FIELD_CATALOG>-SCRTEXT_L = LV_VALUE.
      elseif <SCRTXT2> is not initial.
        LV_VALUE = <SCRTXT2>.
        <LS_FIELD_CATALOG>-SCRTEXT_L = LV_VALUE.
      elseif <SCRTXT3> is not initial.
        LV_VALUE = <SCRTXT3>.
        <LS_FIELD_CATALOG>-SCRTEXT_L = LV_VALUE.
      else.
        LV_VALUE = 'Column'.  " default value as Excel does
        <LS_FIELD_CATALOG>-SCRTEXT_L = LV_VALUE.
      endif.
      while 1 = 1.
        LV_VALUE_LOWERCASE = LV_VALUE.
        translate LV_VALUE_LOWERCASE to lower case.
        read table LT_COLUMN_NAME_BUFFER transporting no fields with key TABLE_LINE = LV_VALUE_LOWERCASE binary search.
        if SY-SUBRC <> 0.
          <LS_FIELD_CATALOG>-SCRTEXT_L = LV_VALUE.
          insert LV_VALUE_LOWERCASE into table LT_COLUMN_NAME_BUFFER.
          exit.
        else.
          LV_SYINDEX = SY-INDEX.
          concatenate <LS_FIELD_CATALOG>-SCRTEXT_L LV_SYINDEX into LV_VALUE.
        endif.

      endwhile.
      " First of all write column header
      if <LS_FIELD_CATALOG>-STYLE_HEADER is not initial.
        ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                      IP_ROW    = LV_ROW_INT
                      IP_VALUE  = LV_VALUE
                      IP_STYLE  = <LS_FIELD_CATALOG>-STYLE_HEADER ).
      else.
        ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                      IP_ROW    = LV_ROW_INT
                      IP_VALUE  = LV_VALUE ).
      endif.

      add 1 to LV_ROW_INT.
      loop at IP_TABLE assigning <FS_TABLE_LINE>.

        assign component <LS_FIELD_CATALOG>-FIELDNAME of structure <FS_TABLE_LINE> to <FS_FLDVAL>.
        " issue #290 Add formula support in table
        if <LS_FIELD_CATALOG>-FORMULA eq ABAP_TRUE.
          if <LS_FIELD_CATALOG>-STYLE is not initial.
            if <LS_FIELD_CATALOG>-ABAP_TYPE is not initial.
              ME->SET_CELL( IP_COLUMN   = LV_COLUMN_ALPHA
                          IP_ROW      = LV_ROW_INT
                          IP_FORMULA  = <FS_FLDVAL>
                          IP_ABAP_TYPE = <LS_FIELD_CATALOG>-ABAP_TYPE
                          IP_STYLE    = <LS_FIELD_CATALOG>-STYLE ).
            else.
              ME->SET_CELL( IP_COLUMN   = LV_COLUMN_ALPHA
                            IP_ROW      = LV_ROW_INT
                            IP_FORMULA  = <FS_FLDVAL>
                            IP_STYLE    = <LS_FIELD_CATALOG>-STYLE ).
            endif.
          elseif <LS_FIELD_CATALOG>-ABAP_TYPE is not initial.
            ME->SET_CELL( IP_COLUMN   = LV_COLUMN_ALPHA
                          IP_ROW      = LV_ROW_INT
                          IP_FORMULA  = <FS_FLDVAL>
                          IP_ABAP_TYPE = <LS_FIELD_CATALOG>-ABAP_TYPE ).
          else.
            ME->SET_CELL( IP_COLUMN   = LV_COLUMN_ALPHA
                          IP_ROW      = LV_ROW_INT
                          IP_FORMULA  = <FS_FLDVAL> ).
          endif.
        else.
          if <LS_FIELD_CATALOG>-STYLE is not initial.
            if <LS_FIELD_CATALOG>-ABAP_TYPE is not initial.
              ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                          IP_ROW    = LV_ROW_INT
                          IP_VALUE  = <FS_FLDVAL>
                          IP_ABAP_TYPE = <LS_FIELD_CATALOG>-ABAP_TYPE
                          IP_STYLE  = <LS_FIELD_CATALOG>-STYLE ).
            else.
              ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                            IP_ROW    = LV_ROW_INT
                            IP_VALUE  = <FS_FLDVAL>
                            IP_STYLE  = <LS_FIELD_CATALOG>-STYLE ).
            endif.
          else.
            if <LS_FIELD_CATALOG>-ABAP_TYPE is not initial.
              ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                          IP_ROW    = LV_ROW_INT
                          IP_ABAP_TYPE = <LS_FIELD_CATALOG>-ABAP_TYPE
                          IP_VALUE  = <FS_FLDVAL> ).
            else.
              ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                            IP_ROW    = LV_ROW_INT
                            IP_VALUE  = <FS_FLDVAL> ).
            endif.
          endif.
        endif.
        add 1 to LV_ROW_INT.

      endloop.
      if SY-SUBRC <> 0. "create empty row if table has no data
        ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                      IP_ROW    = LV_ROW_INT
                      IP_VALUE  = SPACE ).
        add 1 to LV_ROW_INT.
      endif.

*--------------------------------------------------------------------*
      " totals
*--------------------------------------------------------------------*
      if <LS_FIELD_CATALOG>-TOTALS_FUNCTION is not initial.
        LV_FORMULA = LO_TABLE->GET_TOTALS_FORMULA( IP_COLUMN = <LS_FIELD_CATALOG>-SCRTEXT_L IP_FUNCTION = <LS_FIELD_CATALOG>-TOTALS_FUNCTION ).
        if <LS_FIELD_CATALOG>-STYLE_TOTAL is not initial.
          ME->SET_CELL( IP_COLUMN   = LV_COLUMN_ALPHA
                        IP_ROW      = LV_ROW_INT
                        IP_FORMULA  = LV_FORMULA
                        IP_STYLE    = <LS_FIELD_CATALOG>-STYLE_TOTAL ).
        else.
          ME->SET_CELL( IP_COLUMN   = LV_COLUMN_ALPHA
                        IP_ROW      = LV_ROW_INT
                        IP_FORMULA  = LV_FORMULA ).
        endif.
      endif.

      LV_ROW_INT = LS_SETTINGS-TOP_LEFT_ROW.
      add 1 to LV_COLUMN_INT.

*--------------------------------------------------------------------*
      " conditional formatting
*--------------------------------------------------------------------*
      if <LS_FIELD_CATALOG>-STYLE_COND is not initial.
        LV_FIRST_ROW    = LS_SETTINGS-TOP_LEFT_ROW + 1. " +1 to exclude header
        LV_LAST_ROW     = LS_SETTINGS-TOP_LEFT_ROW + LV_ROWS.
        LO_STYLE_COND = ME->GET_STYLE_COND( <LS_FIELD_CATALOG>-STYLE_COND ).
        LO_STYLE_COND->SET_RANGE( IP_START_COLUMN  = LV_COLUMN_ALPHA
                                  IP_START_ROW     = LV_FIRST_ROW
                                  IP_STOP_COLUMN   = LV_COLUMN_ALPHA
                                  IP_STOP_ROW      = LV_LAST_ROW ).
      endif.

    endloop.

*--------------------------------------------------------------------*
    " Set field catalog
*--------------------------------------------------------------------*
    LO_TABLE->FIELDCAT = LT_FIELD_CATALOG[].

    ES_TABLE_SETTINGS = LS_SETTINGS.
    ES_TABLE_SETTINGS-BOTTOM_RIGHT_COLUMN = LV_COLUMN_ALPHA.
    " >> Issue #291
    if IP_TABLE is initial.
      ES_TABLE_SETTINGS-BOTTOM_RIGHT_ROW    = LS_SETTINGS-TOP_LEFT_ROW + 2.           "Last rows
    else.
      ES_TABLE_SETTINGS-BOTTOM_RIGHT_ROW    = LS_SETTINGS-TOP_LEFT_ROW + LV_ROWS + 1. "Last rows
    endif.
    " << Issue #291

  endmethod.


  method CALCULATE_CELL_WIDTH.
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

    constants:
      LC_DEFAULT_FONT_NAME   type ZEXCEL_STYLE_FONT_NAME value 'Calibri', "#EC NOTEXT
      LC_DEFAULT_FONT_HEIGHT type TDFONTSIZE value '110',
      LC_EXCEL_CELL_PADDING  type FLOAT value '0.75'.

    data: LD_CELL_VALUE                type ZEXCEL_CELL_VALUE,
          LD_CURRENT_CHARACTER         type C length 1,
          LD_STYLE_GUID                type ZEXCEL_CELL_STYLE,
          LS_STYLEMAPPING              type ZEXCEL_S_STYLEMAPPING,
          LO_TABLE_OBJECT              type ref to OBJECT,
          LO_TABLE                     type ref to ZCL_EXCEL_TABLE,
          LD_TABLE_TOP_LEFT_COLUMN     type ZEXCEL_CELL_COLUMN,
          LD_TABLE_BOTTOM_RIGHT_COLUMN type ZEXCEL_CELL_COLUMN,
          LD_FLAG_CONTAINS_AUTO_FILTER type ABAP_BOOL value ABAP_FALSE,
          LD_FLAG_BOLD                 type ABAP_BOOL value ABAP_FALSE,
          LD_FLAG_ITALIC               type ABAP_BOOL value ABAP_FALSE,
          LD_DATE                      type D,
          LD_DATE_CHAR                 type C length 50,
          LD_FONT_HEIGHT               type TDFONTSIZE value LC_DEFAULT_FONT_HEIGHT,
          LT_ITCFC                     type standard table of ITCFC,
          LD_OFFSET                    type I,
          LD_LENGTH                    type I,
          LD_UCCP                      type I,
          LS_FONT_METRIC               type MTY_S_FONT_METRIC,
          LD_WIDTH_FROM_FONT_METRICS   type I,
          LD_FONT_FAMILY               type ITCFH-TDFAMILY,
          LD_FONT_NAME                 type ZEXCEL_STYLE_FONT_NAME value LC_DEFAULT_FONT_NAME,
          LT_FONT_FAMILIES             like standard table of LD_FONT_FAMILY,
          LS_FONT_CACHE                type MTY_S_FONT_CACHE.

    field-symbols: <LS_FONT_CACHE>  type MTY_S_FONT_CACHE,
                   <LS_FONT_METRIC> type MTY_S_FONT_METRIC,
                   <LS_ITCFC>       type ITCFC.

    " Determine cell content and cell style
    ME->GET_CELL( exporting IP_COLUMN = IP_COLUMN
                            IP_ROW    = IP_ROW
                  importing EP_VALUE  = LD_CELL_VALUE
                            EP_GUID   = LD_STYLE_GUID ).

    " ABAP2XLSX uses tables to define areas containing headers and
    " auto-filters. Find out if the current cell is in the header
    " of one of these tables.
    loop at ME->TABLES->COLLECTION into LO_TABLE_OBJECT.
      " Downcast: OBJECT -> ZCL_EXCEL_TABLE
      LO_TABLE ?= LO_TABLE_OBJECT.

      " Convert column letters to corresponding integer values
      LD_TABLE_TOP_LEFT_COLUMN =
        ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT(
          LO_TABLE->SETTINGS-TOP_LEFT_COLUMN ).

      LD_TABLE_BOTTOM_RIGHT_COLUMN =
        ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT(
          LO_TABLE->SETTINGS-BOTTOM_RIGHT_COLUMN ).

      " Is the current cell part of the table header?
      if IP_COLUMN between LD_TABLE_TOP_LEFT_COLUMN and
                           LD_TABLE_BOTTOM_RIGHT_COLUMN and
         IP_ROW    eq LO_TABLE->SETTINGS-TOP_LEFT_ROW.
        " Current cell is part of the table header
        " -> Assume that an auto filter is present and that the font is
        "    bold
        LD_FLAG_CONTAINS_AUTO_FILTER = ABAP_TRUE.
        LD_FLAG_BOLD = ABAP_TRUE.
      endif.
    endloop.

    " If a style GUID is present, read style attributes
    if LD_STYLE_GUID is not initial.
      try.
          " Read style attributes
          LS_STYLEMAPPING = ME->EXCEL->GET_STYLE_TO_GUID( LD_STYLE_GUID ).

          " If the current cell contains the default date format,
          " convert the cell value to a date and calculate its length
          if LS_STYLEMAPPING-COMPLETE_STYLE-NUMBER_FORMAT-FORMAT_CODE =
             ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_STD.

            " Convert excel date to ABAP date
            LD_DATE =
              ZCL_EXCEL_COMMON=>EXCEL_STRING_TO_DATE( LD_CELL_VALUE ).

            " Format ABAP date using user's formatting settings
            write LD_DATE to LD_DATE_CHAR.

            " Remember the formatted date to calculate the cell size
            LD_CELL_VALUE = LD_DATE_CHAR.

          endif.

          " Read the font size and convert it to the font height
          " used by SAPscript (multiplication by 10)
          if LS_STYLEMAPPING-COMPLETE_STYLEX-FONT-SIZE = ABAP_TRUE.
            LD_FONT_HEIGHT = LS_STYLEMAPPING-COMPLETE_STYLE-FONT-SIZE * 10.
          endif.

          " If set, remember the font name
          if LS_STYLEMAPPING-COMPLETE_STYLEX-FONT-NAME = ABAP_TRUE.
            LD_FONT_NAME = LS_STYLEMAPPING-COMPLETE_STYLE-FONT-NAME.
          endif.

          " If set, remember whether font is bold and italic.
          if LS_STYLEMAPPING-COMPLETE_STYLEX-FONT-BOLD = ABAP_TRUE.
            LD_FLAG_BOLD = LS_STYLEMAPPING-COMPLETE_STYLE-FONT-BOLD.
          endif.

          if LS_STYLEMAPPING-COMPLETE_STYLEX-FONT-ITALIC = ABAP_TRUE.
            LD_FLAG_ITALIC = LS_STYLEMAPPING-COMPLETE_STYLE-FONT-ITALIC.
          endif.

        catch ZCX_EXCEL.                                "#EC NO_HANDLER
          " Style GUID is present, but style was not found
          " Continue with default values

      endtry.
    endif.

    " Check if the same font (font name and font attributes) was already
    " used before
    read table MTH_FONT_CACHE
      with table key
        FONT_NAME   = LD_FONT_NAME
        FONT_HEIGHT = LD_FONT_HEIGHT
        FLAG_BOLD   = LD_FLAG_BOLD
        FLAG_ITALIC = LD_FLAG_ITALIC
      assigning <LS_FONT_CACHE>.

    if SY-SUBRC <> 0.
      " Font is used for the first time
      " Add the font to our local font cache
      LS_FONT_CACHE-FONT_NAME   = LD_FONT_NAME.
      LS_FONT_CACHE-FONT_HEIGHT = LD_FONT_HEIGHT.
      LS_FONT_CACHE-FLAG_BOLD   = LD_FLAG_BOLD.
      LS_FONT_CACHE-FLAG_ITALIC = LD_FLAG_ITALIC.
      insert LS_FONT_CACHE into table MTH_FONT_CACHE
        assigning <LS_FONT_CACHE>.

      " Determine the SAPscript font family name from the Excel
      " font name
      select TDFAMILY
        from TFO01
        into table LT_FONT_FAMILIES
        up to 1 rows
        where TDTEXT = LD_FONT_NAME
        order by primary key.

      " Check if a matching font family was found
      " Fonts can be uploaded from TTF files using transaction SE73
      if LINES( LT_FONT_FAMILIES ) > 0.
        read table LT_FONT_FAMILIES index 1 into LD_FONT_FAMILY.

        " Load font metrics (returns a table with the size of each letter
        " in the font)
        call function 'LOAD_FONT'
          exporting
            FAMILY      = LD_FONT_FAMILY
            HEIGHT      = LD_FONT_HEIGHT
            PRINTER     = 'SWIN'
            BOLD        = LD_FLAG_BOLD
            ITALIC      = LD_FLAG_ITALIC
          tables
            METRIC      = LT_ITCFC
          exceptions
            FONT_FAMILY = 1
            CODEPAGE    = 2
            DEVICE_TYPE = 3
            others      = 4.
        if SY-SUBRC <> 0.
          clear LT_ITCFC.
        endif.

        " For faster access, convert each character number to the actual
        " character, and store the characters and their sizes in a hash
        " table
        loop at LT_ITCFC assigning <LS_ITCFC>.
          LD_UCCP = <LS_ITCFC>-CPCHARNO.
          LS_FONT_METRIC-CHAR =
            CL_ABAP_CONV_IN_CE=>UCCPI( LD_UCCP ).
          LS_FONT_METRIC-CHAR_WIDTH = <LS_ITCFC>-TDCWIDTHS.
          insert LS_FONT_METRIC
            into table <LS_FONT_CACHE>-TH_FONT_METRICS.
        endloop.

      endif.
    endif.

    " Calculate the cell width
    " If available, use font metrics
    if LINES( <LS_FONT_CACHE>-TH_FONT_METRICS ) = 0.
      " Font metrics are not available
      " -> Calculate the cell width using only the font size
      LD_LENGTH = STRLEN( LD_CELL_VALUE ).
      EP_WIDTH = LD_LENGTH * LD_FONT_HEIGHT / LC_DEFAULT_FONT_HEIGHT + LC_EXCEL_CELL_PADDING.

    else.
      " Font metrics are available

      " Calculate the size of the text by adding the sizes of each
      " letter
      LD_LENGTH = STRLEN( LD_CELL_VALUE ).
      do LD_LENGTH times.
        " Subtract 1, because the first character is at offset 0
        LD_OFFSET = SY-INDEX - 1.

        " Read the current character from the cell value
        LD_CURRENT_CHARACTER = LD_CELL_VALUE+LD_OFFSET(1).

        " Look up the size of the current letter
        read table <LS_FONT_CACHE>-TH_FONT_METRICS
          with table key CHAR = LD_CURRENT_CHARACTER
          assigning <LS_FONT_METRIC>.
        if SY-SUBRC = 0.
          " The size of the letter is known
          " -> Add the actual size of the letter
          add <LS_FONT_METRIC>-CHAR_WIDTH to LD_WIDTH_FROM_FONT_METRICS.
        else.
          " The size of the letter is unknown
          " -> Add the font height as the default letter size
          add LD_FONT_HEIGHT to LD_WIDTH_FROM_FONT_METRICS.
        endif.
      enddo.

      " Add cell padding (Excel makes columns a bit wider than the space
      " that is needed for the text itself) and convert unit
      " (division by 100)
      EP_WIDTH = LD_WIDTH_FROM_FONT_METRICS / 100 + LC_EXCEL_CELL_PADDING.
    endif.

    " If the current cell contains an auto filter, make it a bit wider.
    " The size used by the auto filter button does not depend on the font
    " size.
    if LD_FLAG_CONTAINS_AUTO_FILTER = ABAP_TRUE.
      add 2 to EP_WIDTH.
    endif.

  endmethod.


  method CALCULATE_COLUMN_WIDTHS.
    types:
      begin of T_AUTO_SIZE,
        COL_INDEX type INT4,
        WIDTH     type FLOAT,
      end   of T_AUTO_SIZE.
    types: TT_AUTO_SIZE type table of T_AUTO_SIZE.

    data: LO_COLUMN_ITERATOR type ref to CL_OBJECT_COLLECTION_ITERATOR,
          LO_COLUMN          type ref to ZCL_EXCEL_COLUMN.

    data: AUTO_SIZE   type FLAG.
    data: AUTO_SIZES  type TT_AUTO_SIZE.
    data: COUNT       type INT4.
    data: HIGHEST_ROW type INT4.
    data: WIDTH       type FLOAT.

    field-symbols: <AUTO_SIZE>        like line of AUTO_SIZES.

    LO_COLUMN_ITERATOR = ME->GET_COLUMNS_ITERATOR( ).
    while LO_COLUMN_ITERATOR->HAS_NEXT( ) = ABAP_TRUE.
      LO_COLUMN ?= LO_COLUMN_ITERATOR->GET_NEXT( ).
      AUTO_SIZE = LO_COLUMN->GET_AUTO_SIZE( ).
      if AUTO_SIZE = ABAP_TRUE.
        append initial line to AUTO_SIZES assigning <AUTO_SIZE>.
        <AUTO_SIZE>-COL_INDEX = LO_COLUMN->GET_COLUMN_INDEX( ).
        <AUTO_SIZE>-WIDTH     = -1.
      endif.
    endwhile.

    " There is only something to do if there are some auto-size columns
    if not AUTO_SIZES is initial.
      HIGHEST_ROW = ME->GET_HIGHEST_ROW( ).
      loop at AUTO_SIZES assigning <AUTO_SIZE>.
        COUNT = 1.
        while COUNT <= HIGHEST_ROW.
* Do not check merged cells
          if IS_CELL_MERGED(
              IP_COLUMN    = <AUTO_SIZE>-COL_INDEX
              IP_ROW       = COUNT ) = ABAP_FALSE.
            WIDTH = CALCULATE_CELL_WIDTH( IP_COLUMN = <AUTO_SIZE>-COL_INDEX     " issue #155 - less restrictive typing for ip_column
                                          IP_ROW    = COUNT ).
            if WIDTH > <AUTO_SIZE>-WIDTH.
              <AUTO_SIZE>-WIDTH = WIDTH.
            endif.
          endif.
          COUNT = COUNT + 1.
        endwhile.
        LO_COLUMN = ME->GET_COLUMN( <AUTO_SIZE>-COL_INDEX ). " issue #155 - less restrictive typing for ip_column
        LO_COLUMN->SET_WIDTH( <AUTO_SIZE>-WIDTH ).
      endloop.
    endif.

  endmethod.


  method CHANGE_CELL_STYLE.
    " issue # 139
    data: STYLEMAPPING    type ZEXCEL_S_STYLEMAPPING,

          COMPLETE_STYLE  type ZEXCEL_S_CSTYLE_COMPLETE,
          COMPLETE_STYLEX type ZEXCEL_S_CSTYLEX_COMPLETE,

          BORDERX         type ZEXCEL_S_CSTYLEX_BORDER,
          L_GUID          type ZEXCEL_CELL_STYLE.   "issue # 177

* We have a lot of parameters.  Use some macros to make the coding more structured

    define CLEAR_INITIAL_COLORXFIELDS.
      if &1-RGB is initial.
        clear &2-RGB.
      endif.
      if &1-INDEXED is initial.
        clear &2-INDEXED.
      endif.
      if &1-THEME is initial.
        clear &2-THEME.
      endif.
      if &1-TINT is initial.
        clear &2-TINT.
      endif.
    end-of-definition.

    define MOVE_SUPPLIED_BORDERS.
      if IP_&1 is supplied.  " only act if parameter was supplied
        if IP_X&1 is supplied.  "
          BORDERX = IP_X&1.          " use supplied x-parameter
        else.
          clear BORDERX with 'X'.
* clear in a way that would be expected to work easily
          if IP_&1-BORDER_STYLE is  initial.
            clear BORDERX-BORDER_STYLE.
          endif.
          CLEAR_INITIAL_COLORXFIELDS IP_&1-BORDER_COLOR BORDERX-BORDER_COLOR.
        endif.
        move-corresponding IP_&1   to COMPLETE_STYLE-&2.
        move-corresponding BORDERX to COMPLETE_STYLEX-&2.
      endif.
    end-of-definition.

* First get current stylsettings
    try.
        ME->GET_CELL( exporting IP_COLUMN = IP_COLUMN  " Cell Column
                                IP_ROW    = IP_ROW      " Cell Row
                      importing EP_GUID   = L_GUID )." Cell Value ).  "issue # 177


        STYLEMAPPING = ME->EXCEL->GET_STYLE_TO_GUID( L_GUID ).        "issue # 177
        COMPLETE_STYLE  = STYLEMAPPING-COMPLETE_STYLE.
        COMPLETE_STYLEX = STYLEMAPPING-COMPLETE_STYLEX.
      catch ZCX_EXCEL.
* Error --> use submitted style
    endtry.

*  move_supplied_multistyles: complete.
    if IP_COMPLETE is supplied.
      if IP_XCOMPLETE is not supplied.
        zcx_excel=>raise_text( 'Complete styleinfo has to be supplied with corresponding X-field' ).
      endif.
      move-corresponding IP_COMPLETE  to COMPLETE_STYLE.
      move-corresponding IP_XCOMPLETE to COMPLETE_STYLEX.
    endif.



    if IP_FONT is supplied.
      data: FONTX like IP_XFONT.
      if IP_XFONT is supplied.
        FONTX = IP_XFONT.
      else.
* Only supplied values should be used - exception: Flags bold and italic strikethrough underline
        move 'X' to: FONTX-BOLD,
                     FONTX-ITALIC,
                     FONTX-STRIKETHROUGH,
                     FONTX-UNDERLINE_MODE.
        clear FONTX-COLOR with 'X'.
        CLEAR_INITIAL_COLORXFIELDS IP_FONT-COLOR FONTX-COLOR.
        if IP_FONT-FAMILY is not initial.
          FONTX-FAMILY = 'X'.
        endif.
        if IP_FONT-NAME is not initial.
          FONTX-NAME = 'X'.
        endif.
        if IP_FONT-SCHEME is not initial.
          FONTX-SCHEME = 'X'.
        endif.
        if IP_FONT-SIZE is not initial.
          FONTX-SIZE = 'X'.
        endif.
        if IP_FONT-UNDERLINE_MODE is not initial.
          FONTX-UNDERLINE_MODE = 'X'.
        endif.
      endif.
      move-corresponding IP_FONT  to COMPLETE_STYLE-FONT.
      move-corresponding FONTX    to COMPLETE_STYLEX-FONT.
* Correction for undeline mode
    endif.

    if IP_FILL is supplied.
      data: FILLX like IP_XFILL.
      if IP_XFILL is supplied.
        FILLX = IP_XFILL.
      else.
        clear FILLX with 'X'.
        if IP_FILL-FILLTYPE is initial.
          clear FILLX-FILLTYPE.
        endif.
        CLEAR_INITIAL_COLORXFIELDS IP_FILL-FGCOLOR FILLX-FGCOLOR.
        CLEAR_INITIAL_COLORXFIELDS IP_FILL-BGCOLOR FILLX-BGCOLOR.

      endif.
      move-corresponding IP_FILL  to COMPLETE_STYLE-FILL.
      move-corresponding FILLX    to COMPLETE_STYLEX-FILL.
    endif.


    if IP_BORDERS is supplied.
      data: BORDERSX like IP_XBORDERS.
      if IP_XBORDERS is supplied.
        BORDERSX = IP_XBORDERS.
      else.
        clear BORDERSX with 'X'.
        if IP_BORDERS-ALLBORDERS-BORDER_STYLE is initial.
          clear BORDERSX-ALLBORDERS-BORDER_STYLE.
        endif.
        if IP_BORDERS-DIAGONAL-BORDER_STYLE is initial.
          clear BORDERSX-DIAGONAL-BORDER_STYLE.
        endif.
        if IP_BORDERS-DOWN-BORDER_STYLE is initial.
          clear BORDERSX-DOWN-BORDER_STYLE.
        endif.
        if IP_BORDERS-LEFT-BORDER_STYLE is initial.
          clear BORDERSX-LEFT-BORDER_STYLE.
        endif.
        if IP_BORDERS-RIGHT-BORDER_STYLE is initial.
          clear BORDERSX-RIGHT-BORDER_STYLE.
        endif.
        if IP_BORDERS-TOP-BORDER_STYLE is initial.
          clear BORDERSX-TOP-BORDER_STYLE.
        endif.
        CLEAR_INITIAL_COLORXFIELDS IP_BORDERS-ALLBORDERS-BORDER_COLOR BORDERSX-ALLBORDERS-BORDER_COLOR.
        CLEAR_INITIAL_COLORXFIELDS IP_BORDERS-DIAGONAL-BORDER_COLOR   BORDERSX-DIAGONAL-BORDER_COLOR.
        CLEAR_INITIAL_COLORXFIELDS IP_BORDERS-DOWN-BORDER_COLOR       BORDERSX-DOWN-BORDER_COLOR.
        CLEAR_INITIAL_COLORXFIELDS IP_BORDERS-LEFT-BORDER_COLOR       BORDERSX-LEFT-BORDER_COLOR.
        CLEAR_INITIAL_COLORXFIELDS IP_BORDERS-RIGHT-BORDER_COLOR      BORDERSX-RIGHT-BORDER_COLOR.
        CLEAR_INITIAL_COLORXFIELDS IP_BORDERS-TOP-BORDER_COLOR        BORDERSX-TOP-BORDER_COLOR.

      endif.
      move-corresponding IP_BORDERS  to COMPLETE_STYLE-BORDERS.
      move-corresponding BORDERSX    to COMPLETE_STYLEX-BORDERS.
    endif.

    if IP_ALIGNMENT is supplied.
      data: ALIGNMENTX like IP_XALIGNMENT.
      if IP_XALIGNMENT is supplied.
        ALIGNMENTX = IP_XALIGNMENT.
      else.
        clear ALIGNMENTX with 'X'.
        if IP_ALIGNMENT-HORIZONTAL is initial.
          clear ALIGNMENTX-HORIZONTAL.
        endif.
        if IP_ALIGNMENT-VERTICAL is initial.
          clear ALIGNMENTX-VERTICAL.
        endif.
      endif.
      move-corresponding IP_ALIGNMENT  to COMPLETE_STYLE-ALIGNMENT.
      move-corresponding ALIGNMENTX    to COMPLETE_STYLEX-ALIGNMENT.
    endif.

    if IP_PROTECTION is supplied.
      move-corresponding IP_PROTECTION  to COMPLETE_STYLE-PROTECTION.
      if IP_XPROTECTION is supplied.
        move-corresponding IP_XPROTECTION to COMPLETE_STYLEX-PROTECTION.
      else.
        if IP_PROTECTION-HIDDEN is not initial.
          COMPLETE_STYLEX-PROTECTION-HIDDEN = 'X'.
        endif.
        if IP_PROTECTION-LOCKED is not initial.
          COMPLETE_STYLEX-PROTECTION-LOCKED = 'X'.
        endif.
      endif.
    endif.


    MOVE_SUPPLIED_BORDERS    : BORDERS_ALLBORDERS BORDERS-ALLBORDERS,
                               BORDERS_DIAGONAL   BORDERS-DIAGONAL  ,
                               BORDERS_DOWN       BORDERS-DOWN      ,
                               BORDERS_LEFT       BORDERS-LEFT      ,
                               BORDERS_RIGHT      BORDERS-RIGHT     ,
                               BORDERS_TOP        BORDERS-TOP       .

    define MOVE_SUPPLIED_SINGLESTYLES.
      if IP_&1 is supplied.
        COMPLETE_STYLE-&2 = IP_&1.
        COMPLETE_STYLEX-&2 = 'X'.
      endif.
    end-of-definition.

    MOVE_SUPPLIED_SINGLESTYLES: NUMBER_FORMAT_FORMAT_CODE  NUMBER_FORMAT-FORMAT_CODE,
                                FONT_BOLD                     FONT-BOLD,
                                FONT_COLOR                    FONT-COLOR,
                                FONT_COLOR_RGB                FONT-COLOR-RGB,
                                FONT_COLOR_INDEXED            FONT-COLOR-INDEXED,
                                FONT_COLOR_THEME              FONT-COLOR-THEME,
                                FONT_COLOR_TINT               FONT-COLOR-TINT,

                                FONT_FAMILY                   FONT-FAMILY,
                                FONT_ITALIC                   FONT-ITALIC,
                                FONT_NAME                     FONT-NAME,
                                FONT_SCHEME                   FONT-SCHEME,
                                FONT_SIZE                     FONT-SIZE,
                                FONT_STRIKETHROUGH            FONT-STRIKETHROUGH,
                                FONT_UNDERLINE                FONT-UNDERLINE,
                                FONT_UNDERLINE_MODE           FONT-UNDERLINE_MODE,
                                FILL_FILLTYPE                 FILL-FILLTYPE,
                                FILL_ROTATION                 FILL-ROTATION,
                                FILL_FGCOLOR                  FILL-FGCOLOR,
                                FILL_FGCOLOR_RGB              FILL-FGCOLOR-RGB,
                                FILL_FGCOLOR_INDEXED          FILL-FGCOLOR-INDEXED,
                                FILL_FGCOLOR_THEME            FILL-FGCOLOR-THEME,
                                FILL_FGCOLOR_TINT             FILL-FGCOLOR-TINT,

                                FILL_BGCOLOR                  FILL-BGCOLOR,
                                FILL_BGCOLOR_RGB              FILL-BGCOLOR-RGB,
                                FILL_BGCOLOR_INDEXED          FILL-BGCOLOR-INDEXED,
                                FILL_BGCOLOR_THEME            FILL-BGCOLOR-THEME,
                                FILL_BGCOLOR_TINT             FILL-BGCOLOR-TINT,

                                FILL_GRADTYPE_TYPE            FILL-GRADTYPE-TYPE,
                                FILL_GRADTYPE_DEGREE          FILL-GRADTYPE-DEGREE,
                                FILL_GRADTYPE_BOTTOM          FILL-GRADTYPE-BOTTOM,
                                FILL_GRADTYPE_LEFT            FILL-GRADTYPE-LEFT,
                                FILL_GRADTYPE_TOP             FILL-GRADTYPE-TOP,
                                FILL_GRADTYPE_RIGHT           FILL-GRADTYPE-RIGHT,
                                FILL_GRADTYPE_POSITION1       FILL-GRADTYPE-POSITION1,
                                FILL_GRADTYPE_POSITION2       FILL-GRADTYPE-POSITION2,
                                FILL_GRADTYPE_POSITION3       FILL-GRADTYPE-POSITION3,



                                BORDERS_DIAGONAL_MODE         BORDERS-DIAGONAL_MODE,
                                ALIGNMENT_HORIZONTAL          ALIGNMENT-HORIZONTAL,
                                ALIGNMENT_VERTICAL            ALIGNMENT-VERTICAL,
                                ALIGNMENT_TEXTROTATION        ALIGNMENT-TEXTROTATION,
                                ALIGNMENT_WRAPTEXT            ALIGNMENT-WRAPTEXT,
                                ALIGNMENT_SHRINKTOFIT         ALIGNMENT-SHRINKTOFIT,
                                ALIGNMENT_INDENT              ALIGNMENT-INDENT,
                                PROTECTION_HIDDEN             PROTECTION-HIDDEN,
                                PROTECTION_LOCKED             PROTECTION-LOCKED,

                                BORDERS_ALLBORDERS_STYLE      BORDERS-ALLBORDERS-BORDER_STYLE,
                                BORDERS_ALLBORDERS_COLOR      BORDERS-ALLBORDERS-BORDER_COLOR,
                                BORDERS_ALLBO_COLOR_RGB       BORDERS-ALLBORDERS-BORDER_COLOR-RGB,
                                BORDERS_ALLBO_COLOR_INDEXED   BORDERS-ALLBORDERS-BORDER_COLOR-INDEXED,
                                BORDERS_ALLBO_COLOR_THEME     BORDERS-ALLBORDERS-BORDER_COLOR-THEME,
                                BORDERS_ALLBO_COLOR_TINT      BORDERS-ALLBORDERS-BORDER_COLOR-TINT,

                                BORDERS_DIAGONAL_STYLE        BORDERS-DIAGONAL-BORDER_STYLE,
                                BORDERS_DIAGONAL_COLOR        BORDERS-DIAGONAL-BORDER_COLOR,
                                BORDERS_DIAGONAL_COLOR_RGB    BORDERS-DIAGONAL-BORDER_COLOR-RGB,
                                BORDERS_DIAGONAL_COLOR_INDE   BORDERS-DIAGONAL-BORDER_COLOR-INDEXED,
                                BORDERS_DIAGONAL_COLOR_THEM   BORDERS-DIAGONAL-BORDER_COLOR-THEME,
                                BORDERS_DIAGONAL_COLOR_TINT   BORDERS-DIAGONAL-BORDER_COLOR-TINT,

                                BORDERS_DOWN_STYLE            BORDERS-DOWN-BORDER_STYLE,
                                BORDERS_DOWN_COLOR            BORDERS-DOWN-BORDER_COLOR,
                                BORDERS_DOWN_COLOR_RGB        BORDERS-DOWN-BORDER_COLOR-RGB,
                                BORDERS_DOWN_COLOR_INDEXED    BORDERS-DOWN-BORDER_COLOR-INDEXED,
                                BORDERS_DOWN_COLOR_THEME      BORDERS-DOWN-BORDER_COLOR-THEME,
                                BORDERS_DOWN_COLOR_TINT       BORDERS-DOWN-BORDER_COLOR-TINT,

                                BORDERS_LEFT_STYLE            BORDERS-LEFT-BORDER_STYLE,
                                BORDERS_LEFT_COLOR            BORDERS-LEFT-BORDER_COLOR,
                                BORDERS_LEFT_COLOR_RGB        BORDERS-LEFT-BORDER_COLOR-RGB,
                                BORDERS_LEFT_COLOR_INDEXED    BORDERS-LEFT-BORDER_COLOR-INDEXED,
                                BORDERS_LEFT_COLOR_THEME      BORDERS-LEFT-BORDER_COLOR-THEME,
                                BORDERS_LEFT_COLOR_TINT       BORDERS-LEFT-BORDER_COLOR-TINT,

                                BORDERS_RIGHT_STYLE           BORDERS-RIGHT-BORDER_STYLE,
                                BORDERS_RIGHT_COLOR           BORDERS-RIGHT-BORDER_COLOR,
                                BORDERS_RIGHT_COLOR_RGB       BORDERS-RIGHT-BORDER_COLOR-RGB,
                                BORDERS_RIGHT_COLOR_INDEXED   BORDERS-RIGHT-BORDER_COLOR-INDEXED,
                                BORDERS_RIGHT_COLOR_THEME     BORDERS-RIGHT-BORDER_COLOR-THEME,
                                BORDERS_RIGHT_COLOR_TINT      BORDERS-RIGHT-BORDER_COLOR-TINT,

                                BORDERS_TOP_STYLE             BORDERS-TOP-BORDER_STYLE,
                                BORDERS_TOP_COLOR             BORDERS-TOP-BORDER_COLOR,
                                BORDERS_TOP_COLOR_RGB         BORDERS-TOP-BORDER_COLOR-RGB,
                                BORDERS_TOP_COLOR_INDEXED     BORDERS-TOP-BORDER_COLOR-INDEXED,
                                BORDERS_TOP_COLOR_THEME       BORDERS-TOP-BORDER_COLOR-THEME,
                                BORDERS_TOP_COLOR_TINT        BORDERS-TOP-BORDER_COLOR-TINT.


* Now we have a completly filled styles.
* This can be used to get the guid
* Return guid if requested.  Might be used if copy&paste of styles is requested
    EP_GUID = ME->EXCEL->GET_STATIC_CELLSTYLE_GUID( IP_CSTYLE_COMPLETE  = COMPLETE_STYLE
                                                    IP_CSTYLEX_COMPLETE = COMPLETE_STYLEX  ).
    ME->SET_CELL_STYLE( IP_COLUMN = IP_COLUMN
                        IP_ROW    = IP_ROW
                        IP_STYLE  = EP_GUID ).

  endmethod.


  method CONSTRUCTOR.
    data: LV_TITLE type ZEXCEL_SHEET_TITLE.

    ME->EXCEL = IP_EXCEL.

*  CALL FUNCTION 'GUID_CREATE'                                    " del issue #379 - function is outdated in newer releases
*    IMPORTING
*      ev_guid_16 = me->guid.
    ME->GUID = ZCL_EXCEL_OBSOLETE_FUNC_WRAP=>GUID_CREATE( ).        " ins issue #379 - replacement for outdated function call

    if IP_TITLE is not initial.
      LV_TITLE = IP_TITLE.
    else.
*    lv_title = me->guid.             " del issue #154 - Names of worksheets
      LV_TITLE = ME->GENERATE_TITLE( ). " ins issue #154 - Names of worksheets
    endif.

    ME->SET_TITLE( IP_TITLE = LV_TITLE ).

    create object SHEET_SETUP.
    create object STYLES_COND.
    create object DATA_VALIDATIONS.
    create object TABLES.
    create object COLUMNS.
    create object ROWS.
    create object RANGES. " issue #163
    create object MO_PAGEBREAKS.
    create object DRAWINGS
      exporting
        IP_TYPE = ZCL_EXCEL_DRAWING=>TYPE_IMAGE.
    create object CHARTS
      exporting
        IP_TYPE = ZCL_EXCEL_DRAWING=>TYPE_CHART.
    ME->ZIF_EXCEL_SHEET_PROTECTION~INITIALIZE( ).
    ME->ZIF_EXCEL_SHEET_PROPERTIES~INITIALIZE( ).
    create object HYPERLINKS.

* initialize active cell coordinates
    ACTIVE_CELL-CELL_ROW = 1.
    ACTIVE_CELL-CELL_COLUMN = 1.

* inizialize dimension range
    LOWER_CELL-CELL_ROW     = 1.
    LOWER_CELL-CELL_COLUMN  = 1.
    UPPER_CELL-CELL_ROW     = 1.
    UPPER_CELL-CELL_COLUMN  = 1.

  endmethod.


  method DELETE_MERGE.

    data: LV_COLUMN type I.
*--------------------------------------------------------------------*
* If cell information is passed delete merge including this cell,
* otherwise delete all merges
*--------------------------------------------------------------------*
    if   IP_CELL_COLUMN is initial
      or IP_CELL_ROW    is initial.
      clear ME->MT_MERGED_CELLS.
    else.
      LV_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_CELL_COLUMN ).

      loop at ME->MT_MERGED_CELLS transporting no fields
      where
          ( ROW_FROM <= IP_CELL_ROW and ROW_TO >= IP_CELL_ROW )
      and
          ( COL_FROM <= LV_COLUMN and COL_TO >= LV_COLUMN ).

        delete ME->MT_MERGED_CELLS.
        exit.
      endloop.
    endif.

  endmethod.


  method DELETE_ROW_OUTLINE.

    delete ME->MT_ROW_OUTLINES where ROW_FROM = IV_ROW_FROM
                                 and ROW_TO   = IV_ROW_TO.
    if SY-SUBRC <> 0.  " didn't find outline that was to be deleted
      zcx_excel=>raise_text( 'Row outline to be deleted does not exist' ).
    endif.

  endmethod.


  method FREEZE_PANES.

    if IP_NUM_COLUMNS is not supplied and IP_NUM_ROWS is not supplied.
      zcx_excel=>raise_text( 'Pleas provide number of rows and/or columns to freeze' ).
    endif.

    if IP_NUM_COLUMNS is supplied and IP_NUM_COLUMNS <= 0.
      zcx_excel=>raise_text( 'Number of columns to freeze should be positive' ).
    endif.

    if IP_NUM_ROWS is supplied and IP_NUM_ROWS <= 0.
      zcx_excel=>raise_text( 'Number of rows to freeze should be positive' ).
    endif.

    FREEZE_PANE_CELL_COLUMN = IP_NUM_COLUMNS + 1.
    FREEZE_PANE_CELL_ROW = IP_NUM_ROWS + 1.
  endmethod.


  method GENERATE_TITLE.
    data: LO_WORKSHEETS_ITERATOR type ref to CL_OBJECT_COLLECTION_ITERATOR,
          LO_WORKSHEET           type ref to ZCL_EXCEL_WORKSHEET.

    data: T_TITLES    type hashed table of ZEXCEL_SHEET_TITLE with unique key TABLE_LINE,
          TITLE       type ZEXCEL_SHEET_TITLE,
          SHEETNUMBER type I.

* Get list of currently used titles
    LO_WORKSHEETS_ITERATOR = ME->EXCEL->GET_WORKSHEETS_ITERATOR( ).
    while LO_WORKSHEETS_ITERATOR->HAS_NEXT( ) = ABAP_TRUE.
      LO_WORKSHEET ?= LO_WORKSHEETS_ITERATOR->GET_NEXT( ).
      TITLE = LO_WORKSHEET->GET_TITLE( ).
      insert TITLE into table T_TITLES.
      add 1 to SHEETNUMBER.
    endwhile.

* Now build sheetnumber.  Increase counter until we hit a number that is not used so far
    add 1 to SHEETNUMBER.  " Start counting with next number
    do.
      TITLE = SHEETNUMBER.
      shift TITLE left deleting leading SPACE.
      concatenate 'Sheet'(001) TITLE into EP_TITLE.
      insert EP_TITLE into table T_TITLES.
      if SY-SUBRC = 0.  " Title not used so far --> take it
        exit.
      endif.

      add 1 to SHEETNUMBER.
    enddo.
  endmethod.


  method GET_ACTIVE_CELL.

    data: LV_ACTIVE_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA,
          LV_ACTIVE_ROW    type STRING.

    LV_ACTIVE_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( ACTIVE_CELL-CELL_COLUMN ).
    LV_ACTIVE_ROW    = ACTIVE_CELL-CELL_ROW.
    shift LV_ACTIVE_ROW right deleting trailing SPACE.
    shift LV_ACTIVE_ROW left deleting leading SPACE.
    concatenate LV_ACTIVE_COLUMN LV_ACTIVE_ROW into EP_ACTIVE_CELL.

  endmethod.


  method GET_CELL.

    data: LV_COLUMN        type ZEXCEL_CELL_COLUMN,
          LS_SHEET_CONTENT type ZEXCEL_S_CELL_DATA.

    LV_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN ).

    read table SHEET_CONTENT into LS_SHEET_CONTENT with table key CELL_ROW     = IP_ROW
                                                                  CELL_COLUMN  = LV_COLUMN.

    EP_RC = SY-SUBRC.
    EP_VALUE    = LS_SHEET_CONTENT-CELL_VALUE.
    EP_GUID     = LS_SHEET_CONTENT-CELL_STYLE.       " issue 139 - added this to be used for columnwidth calculation
    EP_FORMULA  = LS_SHEET_CONTENT-CELL_FORMULA.

    " Addition to solve issue #120, contribution by Stefan Schmöcker
    data: STYLE_ITERATOR type ref to CL_OBJECT_COLLECTION_ITERATOR,
          STYLE          type ref to ZCL_EXCEL_STYLE.
    if EP_STYLE is requested.
      STYLE_ITERATOR = ME->EXCEL->GET_STYLES_ITERATOR( ).
      while STYLE_ITERATOR->HAS_NEXT( ) = 'X'.
        STYLE ?= STYLE_ITERATOR->GET_NEXT( ).
        if STYLE->GET_GUID( ) = LS_SHEET_CONTENT-CELL_STYLE.
          EP_STYLE = STYLE.
          exit.
        endif.
      endwhile.
    endif.
  endmethod.


  method GET_COLUMN.

    data: LV_COLUMN type ZEXCEL_CELL_COLUMN.

    LV_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN ).

    EO_COLUMN = ME->COLUMNS->GET( IP_INDEX = LV_COLUMN ).

    if EO_COLUMN is not bound.
      EO_COLUMN = ME->ADD_NEW_COLUMN( IP_COLUMN ).
    endif.

  endmethod.


  method GET_COLUMNS.
    EO_COLUMNS = ME->COLUMNS.
  endmethod.


  method GET_COLUMNS_ITERATOR.

    EO_ITERATOR = ME->COLUMNS->GET_ITERATOR( ).

  endmethod.


  method GET_DATA_VALIDATIONS_ITERATOR.

    EO_ITERATOR = ME->DATA_VALIDATIONS->GET_ITERATOR( ).
  endmethod.


  method GET_DATA_VALIDATIONS_SIZE.
    EP_SIZE = ME->DATA_VALIDATIONS->SIZE( ).
  endmethod.


  method GET_DEFAULT_COLUMN.
    if ME->COLUMN_DEFAULT is not bound.
      create object ME->COLUMN_DEFAULT
        exporting
          IP_INDEX     = 'A'         " ????
          IP_WORKSHEET = ME
          IP_EXCEL     = ME->EXCEL.
    endif.

    EO_COLUMN = ME->COLUMN_DEFAULT.
  endmethod.


  method GET_DEFAULT_EXCEL_DATE_FORMAT.
    constants: C_LANG_E type LANG value 'E'.

    if DEFAULT_EXCEL_DATE_FORMAT is not initial.
      EP_DEFAULT_EXCEL_DATE_FORMAT = DEFAULT_EXCEL_DATE_FORMAT.
      return.
    endif.

    "try to get defaults
    try.
        CL_ABAP_DATFM=>GET_DATE_FORMAT_DES( exporting IM_LANGU = C_LANG_E
                                            importing EX_DATEFORMAT = DEFAULT_EXCEL_DATE_FORMAT ).
      catch CX_ABAP_DATFM_FORMAT_UNKNOWN.

    endtry.

    " and fallback to fixed format
    if DEFAULT_EXCEL_DATE_FORMAT is initial.
      DEFAULT_EXCEL_DATE_FORMAT = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_DDMMYYYYDOT.
    endif.

    EP_DEFAULT_EXCEL_DATE_FORMAT = DEFAULT_EXCEL_DATE_FORMAT.
  endmethod.


  method GET_DEFAULT_EXCEL_TIME_FORMAT.
    data: L_TIMEFM type XUTIMEFM.

    if DEFAULT_EXCEL_TIME_FORMAT is not initial.
      EP_DEFAULT_EXCEL_TIME_FORMAT = DEFAULT_EXCEL_TIME_FORMAT.
      return.
    endif.

* Let's get default
    L_TIMEFM = CL_ABAP_TIMEFM=>GET_ENVIRONMENT_TIMEFM( ).
    case L_TIMEFM.
      when 0.
*0  24 Hour Format (Example: 12:05:10)
        DEFAULT_EXCEL_TIME_FORMAT = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_TIME6.
      when 1.
*1  12 Hour Format (Example: 12:05:10 PM)
        DEFAULT_EXCEL_TIME_FORMAT = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_TIME2.
      when 2.
*2  12 Hour Format (Example: 12:05:10 pm) for now all the same. no chnage upper lower
        DEFAULT_EXCEL_TIME_FORMAT = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_TIME2.
      when 3.
*3  Hours from 0 to 11 (Example: 00:05:10 PM)  for now all the same. no chnage upper lower
        DEFAULT_EXCEL_TIME_FORMAT = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_TIME2.
      when 4.
*4  Hours from 0 to 11 (Example: 00:05:10 pm)  for now all the same. no chnage upper lower
        DEFAULT_EXCEL_TIME_FORMAT = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_TIME2.
      when others.
        " and fallback to fixed format
        DEFAULT_EXCEL_TIME_FORMAT = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_TIME6.
    endcase.

    EP_DEFAULT_EXCEL_TIME_FORMAT = DEFAULT_EXCEL_TIME_FORMAT.
  endmethod.


  method GET_DEFAULT_ROW.
    if ME->ROW_DEFAULT is not bound.
      create object ME->ROW_DEFAULT.
    endif.

    EO_ROW = ME->ROW_DEFAULT.
  endmethod.


  method GET_DIMENSION_RANGE.

    ME->UPDATE_DIMENSION_RANGE( ).
    if UPPER_CELL eq LOWER_CELL. "only one cell
      " Worksheet not filled
*    IF upper_cell-cell_coords = '0'.
      if UPPER_CELL-CELL_COORDS is initial.
        EP_DIMENSION_RANGE = 'A1'.
      else.
        EP_DIMENSION_RANGE = UPPER_CELL-CELL_COORDS.
      endif.
    else.
      concatenate UPPER_CELL-CELL_COORDS ':' LOWER_CELL-CELL_COORDS into EP_DIMENSION_RANGE.
    endif.

  endmethod.


  method GET_DRAWINGS.

    data: LO_DRAWING  type ref to ZCL_EXCEL_DRAWING,
          LO_ITERATOR type ref to CL_OBJECT_COLLECTION_ITERATOR.

    case IP_TYPE.
      when ZCL_EXCEL_DRAWING=>TYPE_IMAGE.
        R_DRAWINGS = DRAWINGS.
      when ZCL_EXCEL_DRAWING=>TYPE_CHART.
        R_DRAWINGS = CHARTS.
      when SPACE.
        create object R_DRAWINGS
          exporting
            IP_TYPE = ''.

        LO_ITERATOR = DRAWINGS->GET_ITERATOR( ).
        while LO_ITERATOR->HAS_NEXT( ) = ABAP_TRUE.
          LO_DRAWING ?= LO_ITERATOR->GET_NEXT( ).
          R_DRAWINGS->INCLUDE( LO_DRAWING ).
        endwhile.
        LO_ITERATOR = CHARTS->GET_ITERATOR( ).
        while LO_ITERATOR->HAS_NEXT( ) = ABAP_TRUE.
          LO_DRAWING ?= LO_ITERATOR->GET_NEXT( ).
          R_DRAWINGS->INCLUDE( LO_DRAWING ).
        endwhile.
      when others.
    endcase.
  endmethod.


  method GET_DRAWINGS_ITERATOR.
    case IP_TYPE.
      when ZCL_EXCEL_DRAWING=>TYPE_IMAGE.
        EO_ITERATOR = DRAWINGS->GET_ITERATOR( ).
      when ZCL_EXCEL_DRAWING=>TYPE_CHART.
        EO_ITERATOR = CHARTS->GET_ITERATOR( ).
    endcase.
  endmethod.


  method GET_FREEZE_CELL.
    EP_ROW = ME->FREEZE_PANE_CELL_ROW.
    EP_COLUMN = ME->FREEZE_PANE_CELL_COLUMN.
  endmethod.


  method GET_GUID.

    EP_GUID = ME->GUID.

  endmethod.


  method GET_HIGHEST_COLUMN.
    ME->UPDATE_DIMENSION_RANGE( ).
    R_HIGHEST_COLUMN = ME->LOWER_CELL-CELL_COLUMN.
  endmethod.


  method GET_HIGHEST_ROW.
    ME->UPDATE_DIMENSION_RANGE( ).
    R_HIGHEST_ROW = ME->LOWER_CELL-CELL_ROW.
  endmethod.


  method GET_HYPERLINKS_ITERATOR.
    EO_ITERATOR = HYPERLINKS->GET_ITERATOR( ).
  endmethod.


  method GET_HYPERLINKS_SIZE.
    EP_SIZE = HYPERLINKS->SIZE( ).
  endmethod.


  method GET_MERGE.

    field-symbols: <LS_MERGED_CELL> like line of ME->MT_MERGED_CELLS.

    data: LV_COL_FROM    type STRING,
          LV_COL_TO      type STRING,
          LV_ROW_FROM    type STRING,
          LV_ROW_TO      type STRING,
          LV_MERGE_RANGE type STRING.

    loop at ME->MT_MERGED_CELLS assigning <LS_MERGED_CELL>.

      LV_COL_FROM = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( <LS_MERGED_CELL>-COL_FROM ).
      LV_COL_TO   = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( <LS_MERGED_CELL>-COL_TO   ).
      LV_ROW_FROM = <LS_MERGED_CELL>-ROW_FROM.
      LV_ROW_TO   = <LS_MERGED_CELL>-ROW_TO  .
      concatenate LV_COL_FROM LV_ROW_FROM ':' LV_COL_TO LV_ROW_TO
         into LV_MERGE_RANGE.
      condense LV_MERGE_RANGE no-gaps.
      append LV_MERGE_RANGE to MERGE_RANGE.

    endloop.

  endmethod.


  method GET_PAGEBREAKS.
    RO_PAGEBREAKS = MO_PAGEBREAKS.
  endmethod.


  method GET_RANGES_ITERATOR.

    EO_ITERATOR = ME->RANGES->GET_ITERATOR( ).

  endmethod.


  method GET_ROW.
    EO_ROW = ME->ROWS->GET( IP_INDEX = IP_ROW ).

    if EO_ROW is not bound.
      EO_ROW = ME->ADD_NEW_ROW( IP_ROW ).
    endif.
  endmethod.


  method GET_ROWS.
    EO_ROWS = ME->ROWS.
  endmethod.


  method GET_ROWS_ITERATOR.

    EO_ITERATOR = ME->ROWS->GET_ITERATOR( ).

  endmethod.


  method GET_ROW_OUTLINES.

    RT_ROW_OUTLINES = ME->MT_ROW_OUTLINES.

  endmethod.


  method GET_STYLE_COND.

    data: LO_STYLE_ITERATOR type ref to CL_OBJECT_COLLECTION_ITERATOR,
          LO_STYLE_COND     type ref to ZCL_EXCEL_STYLE_COND.

    LO_STYLE_ITERATOR = ME->GET_STYLE_COND_ITERATOR( ).
    while LO_STYLE_ITERATOR->HAS_NEXT( ) = ABAP_TRUE.
      LO_STYLE_COND ?= LO_STYLE_ITERATOR->GET_NEXT( ).
      if LO_STYLE_COND->GET_GUID( ) = IP_GUID.
        EO_STYLE_COND = LO_STYLE_COND.
        exit.
      endif.
    endwhile.

  endmethod.


  method GET_STYLE_COND_ITERATOR.

    EO_ITERATOR = STYLES_COND->GET_ITERATOR( ).
  endmethod.


  method GET_TABCOLOR.
    EV_TABCOLOR = ME->TABCOLOR.
  endmethod.


  method GET_TABLE.
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

    field-symbols: <LS_LINE> type DATA.
    field-symbols: <LV_VALUE> type DATA.

    data LV_ACTUAL_ROW type INT4.
    data LV_ACTUAL_COL type INT4.
    data LV_ERRORMESSAGE type STRING.
    data LV_MAX_COL type ZEXCEL_CELL_COLUMN.
    data LV_MAX_ROW type INT4.
    data LV_DELTA_COL type INT4.
    data LV_VALUE  type ZEXCEL_CELL_VALUE.
    data LV_RC  type SYSUBRC.


    LV_MAX_COL =  ME->GET_HIGHEST_COLUMN( ).
    LV_MAX_ROW =  ME->GET_HIGHEST_ROW( ).

*--------------------------------------------------------------------*
* The row counter begins with 1 and should be corrected with the skips
*--------------------------------------------------------------------*
    LV_ACTUAL_ROW =  IV_SKIPPED_ROWS + 1.
    LV_ACTUAL_COL =  IV_SKIPPED_COLS + 1.


    try.
*--------------------------------------------------------------------*
* Check if we the basic features are possible with given "any table"
*--------------------------------------------------------------------*
        append initial line to ET_TABLE assigning <LS_LINE>.
        if SY-SUBRC <> 0 or <LS_LINE> is not assigned.

          LV_ERRORMESSAGE = 'Error at inserting new Line to internal Table'(002).
          zcx_excel=>raise_text( lv_errormessage ).

        else.
          LV_DELTA_COL = LV_MAX_COL - IV_SKIPPED_COLS.
          assign component LV_DELTA_COL of structure <LS_LINE> to <LV_VALUE>.
          if SY-SUBRC <> 0 or <LV_VALUE> is not assigned.
            LV_ERRORMESSAGE = 'Internal table has less columns than excel'(003).
            zcx_excel=>raise_text( lv_errormessage ).
          else.
*--------------------------------------------------------------------*
*now we are ready for handle the table data
*--------------------------------------------------------------------*
            refresh ET_TABLE.
*--------------------------------------------------------------------*
* Handle each Row until end on right side
*--------------------------------------------------------------------*
            while LV_ACTUAL_ROW <= LV_MAX_ROW .

*--------------------------------------------------------------------*
* Handle each Column until end on bottom
* First step is to step back on first column
*--------------------------------------------------------------------*
              LV_ACTUAL_COL =  IV_SKIPPED_COLS + 1.

              unassign <LS_LINE>.
              append initial line to ET_TABLE assigning <LS_LINE>.
              if SY-SUBRC <> 0 or <LS_LINE> is not assigned.
                LV_ERRORMESSAGE = 'Error at inserting new Line to internal Table'(002).
                zcx_excel=>raise_text( lv_errormessage ).
              endif.
              while LV_ACTUAL_COL <= LV_MAX_COL.

                LV_DELTA_COL = LV_ACTUAL_COL - IV_SKIPPED_COLS.
                assign component LV_DELTA_COL of structure <LS_LINE> to <LV_VALUE>.
                if SY-SUBRC <> 0.
                  LV_ERRORMESSAGE = |{ 'Error at assigning field (Col:'(004) } { LV_ACTUAL_COL } { ' Row:'(005) } { LV_ACTUAL_ROW }|.
                  zcx_excel=>raise_text( lv_errormessage ).
                endif.

                ME->GET_CELL(
                  exporting
                    IP_COLUMN  = LV_ACTUAL_COL    " Cell Column
                    IP_ROW     = LV_ACTUAL_ROW    " Cell Row
                  importing
                    EP_VALUE   = LV_VALUE    " Cell Value
                    EP_RC      = LV_RC    " Return Value of ABAP Statements
                ).
                if LV_RC <> 0
                  and LV_RC <> 4.                                                   "No found error means, zero/no value in cell
                  LV_ERRORMESSAGE = |{ 'Error at reading field value (Col:'(007) } { LV_ACTUAL_COL } { ' Row:'(005) } { LV_ACTUAL_ROW }|.
                  zcx_excel=>raise_text( lv_errormessage ).
                endif.

                <LV_VALUE> = LV_VALUE.
*  CATCH zcx_excel.    "
                add 1 to LV_ACTUAL_COL.
              endwhile.
              add 1 to LV_ACTUAL_ROW.
            endwhile.
          endif.


        endif.

      catch CX_SY_ASSIGN_CAST_ILLEGAL_CAST.
        LV_ERRORMESSAGE = |{ 'Error at assigning field (Col:'(004) } { LV_ACTUAL_COL } { ' Row:'(005) } { LV_ACTUAL_ROW }|.
        zcx_excel=>raise_text( lv_errormessage ).
      catch CX_SY_ASSIGN_CAST_UNKNOWN_TYPE.
        LV_ERRORMESSAGE = |{ 'Error at assigning field (Col:'(004) } { LV_ACTUAL_COL } { ' Row:'(005) } { LV_ACTUAL_ROW }|.
        zcx_excel=>raise_text( lv_errormessage ).
      catch CX_SY_ASSIGN_OUT_OF_RANGE.
        LV_ERRORMESSAGE = 'Internal table has less columns than excel'(003).
        zcx_excel=>raise_text( lv_errormessage ).
      catch CX_SY_CONVERSION_ERROR.
        LV_ERRORMESSAGE = |{ 'Error at converting field value (Col:'(006) } { LV_ACTUAL_COL } { ' Row:'(005) } { LV_ACTUAL_ROW }|.
        zcx_excel=>raise_text( lv_errormessage ).

    endtry.
  endmethod.


  method GET_TABLES_ITERATOR.
    EO_ITERATOR = TABLES->IF_OBJECT_COLLECTION~GET_ITERATOR( ).
  endmethod.


  method GET_TABLES_SIZE.
    EP_SIZE = TABLES->IF_OBJECT_COLLECTION~SIZE( ).
  endmethod.


  method GET_TITLE.
    data LV_VALUE type STRING.
    if IP_ESCAPED eq ABAP_TRUE.
      LV_VALUE = ME->TITLE.
      EP_TITLE = ZCL_EXCEL_COMMON=>ESCAPE_STRING( LV_VALUE ).
    else.
      EP_TITLE = ME->TITLE.
    endif.
  endmethod.


  method GET_VALUE_TYPE.
    data: LO_ADDIT    type ref to CL_ABAP_ELEMDESCR,
          LS_DFIES    type DFIES,
          L_FUNCTION  type FUNCNAME,
          L_VALUE(50) type C.

    EP_VALUE = IP_VALUE.
    EP_VALUE_TYPE = CL_ABAP_TYPEDESCR=>TYPEKIND_STRING. " Thats our default if something goes wrong.

    try.
        LO_ADDIT            ?= CL_ABAP_TYPEDESCR=>DESCRIBE_BY_DATA( IP_VALUE ).
      catch CX_SY_MOVE_CAST_ERROR.
        clear LO_ADDIT.
    endtry.
    if LO_ADDIT is bound.
      LO_ADDIT->GET_DDIC_FIELD( receiving  P_FLDDESCR   = LS_DFIES
                                exceptions NOT_FOUND    = 1
                                           NO_DDIC_TYPE = 2
                                           others       = 3 ) .
      if SY-SUBRC = 0.
        EP_VALUE_TYPE = LS_DFIES-INTTYPE.

        if LS_DFIES-CONVEXIT is not initial.
* We need to convert with output conversion function
          concatenate 'CONVERSION_EXIT_' LS_DFIES-CONVEXIT '_OUTPUT' into L_FUNCTION.
          select single FUNCNAME into L_FUNCTION
                from TFDIR
                where FUNCNAME = L_FUNCTION.
          if SY-SUBRC = 0.
            call function L_FUNCTION
              exporting
                INPUT  = IP_VALUE
              importing
*               LONG_TEXT  =
                OUTPUT = L_VALUE
*               SHORT_TEXT =
              exceptions
                others = 1.
            if SY-SUBRC <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
            else.
              try.
                  EP_VALUE = L_VALUE.
                catch CX_ROOT.
                  EP_VALUE = IP_VALUE.
              endtry.
            endif.
          endif.
        endif.
      else.
        EP_VALUE_TYPE = LO_ADDIT->GET_DATA_TYPE_KIND( IP_VALUE ).
      endif.
    endif.

  endmethod.


  method IS_CELL_MERGED.

    data: LV_COLUMN type I.

    field-symbols: <LS_MERGED_CELL> like line of ME->MT_MERGED_CELLS.

    LV_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN ).

    RP_IS_MERGED = ABAP_FALSE.                                        " Assume not in merged area

    loop at ME->MT_MERGED_CELLS assigning <LS_MERGED_CELL>.

      if    <LS_MERGED_CELL>-COL_FROM <= LV_COLUMN
        and <LS_MERGED_CELL>-COL_TO   >= LV_COLUMN
        and <LS_MERGED_CELL>-ROW_FROM <= IP_ROW
        and <LS_MERGED_CELL>-ROW_TO   >= IP_ROW.
        RP_IS_MERGED = ABAP_TRUE.                                     " until we are proven different
        return.
      endif.

    endloop.

  endmethod.


  method PRINT_TITLE_SET_RANGE.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                            2012-12-02
*--------------------------------------------------------------------*


    data: LO_RANGE_ITERATOR         type ref to CL_OBJECT_COLLECTION_ITERATOR,
          LO_RANGE                  type ref to ZCL_EXCEL_RANGE,
          LV_REPEAT_RANGE_SHEETNAME type STRING,
          LV_REPEAT_RANGE_COL       type STRING,
          LV_ROW_CHAR_FROM          type CHAR10,
          LV_ROW_CHAR_TO            type CHAR10,
          LV_REPEAT_RANGE_ROW       type STRING,
          LV_REPEAT_RANGE           type STRING.


*--------------------------------------------------------------------*
* Get range that represents printarea
* if non-existant, create it
*--------------------------------------------------------------------*
    LO_RANGE_ITERATOR = ME->GET_RANGES_ITERATOR( ).
    while LO_RANGE_ITERATOR->HAS_NEXT( ) = ABAP_TRUE.

      LO_RANGE ?= LO_RANGE_ITERATOR->GET_NEXT( ).
      if LO_RANGE->NAME = ZIF_EXCEL_SHEET_PRINTSETTINGS=>GCV_PRINT_TITLE_NAME.
        exit.  " Found it
      endif.
      clear LO_RANGE.

    endwhile.


    if ME->PRINT_TITLE_COL_FROM is initial and
       ME->PRINT_TITLE_ROW_FROM is initial.
*--------------------------------------------------------------------*
* No print titles are present,
*--------------------------------------------------------------------*
      if LO_RANGE is bound.
        ME->RANGES->REMOVE( LO_RANGE ).
      endif.
    else.
*--------------------------------------------------------------------*
* Print titles are present,
*--------------------------------------------------------------------*
      if LO_RANGE is not bound.
        LO_RANGE =  ME->ADD_NEW_RANGE( ).
        LO_RANGE->NAME = ZIF_EXCEL_SHEET_PRINTSETTINGS=>GCV_PRINT_TITLE_NAME.
      endif.

      LV_REPEAT_RANGE_SHEETNAME = ME->GET_TITLE( ).
      LV_REPEAT_RANGE_SHEETNAME = ZCL_EXCEL_COMMON=>ESCAPE_STRING( LV_REPEAT_RANGE_SHEETNAME ).

*--------------------------------------------------------------------*
* Repeat-columns
*--------------------------------------------------------------------*
      if ME->PRINT_TITLE_COL_FROM is not initial.
        concatenate LV_REPEAT_RANGE_SHEETNAME
                    '!$' ME->PRINT_TITLE_COL_FROM
                    ':$' ME->PRINT_TITLE_COL_TO
            into LV_REPEAT_RANGE_COL.
      endif.

*--------------------------------------------------------------------*
* Repeat-rows
*--------------------------------------------------------------------*
      if ME->PRINT_TITLE_ROW_FROM is not initial.
        LV_ROW_CHAR_FROM = ME->PRINT_TITLE_ROW_FROM.
        LV_ROW_CHAR_TO   = ME->PRINT_TITLE_ROW_TO.
        concatenate '!$' LV_ROW_CHAR_FROM
                    ':$' LV_ROW_CHAR_TO
            into LV_REPEAT_RANGE_ROW.
        condense LV_REPEAT_RANGE_ROW no-gaps.
        concatenate LV_REPEAT_RANGE_SHEETNAME
                    LV_REPEAT_RANGE_ROW
            into LV_REPEAT_RANGE_ROW.
      endif.

*--------------------------------------------------------------------*
* Concatenate repeat-rows and columns
*--------------------------------------------------------------------*
      if LV_REPEAT_RANGE_COL is initial.
        LV_REPEAT_RANGE = LV_REPEAT_RANGE_ROW.
      elseif LV_REPEAT_RANGE_ROW is initial.
        LV_REPEAT_RANGE = LV_REPEAT_RANGE_COL.
      else.
        concatenate LV_REPEAT_RANGE_COL LV_REPEAT_RANGE_ROW
            into LV_REPEAT_RANGE separated by ','.
      endif.


      LO_RANGE->SET_RANGE_VALUE( LV_REPEAT_RANGE ).
    endif.



  endmethod.


METHOD set_area.

  DATA: lv_row              TYPE zexcel_cell_row,
        lv_row_end          TYPE zexcel_cell_row,
        lv_column_start     TYPE zexcel_cell_column_alpha,
        lv_column_end       TYPE zexcel_cell_column_alpha,
        lv_column_start_int TYPE zexcel_cell_column_alpha,
        lv_column_end_int   TYPE zexcel_cell_column_alpha.

  MOVE: ip_row_to TO lv_row_end,
        ip_row    TO lv_row.

  IF lv_row_end IS INITIAL OR ip_row_to IS NOT SUPPLIED.
    lv_row_end = lv_row.
  ENDIF.

  MOVE: ip_column_start TO lv_column_start,
        ip_column_end   TO lv_column_end.

  IF lv_column_end IS INITIAL OR ip_column_end IS NOT SUPPLIED.
    lv_column_end = lv_column_start.
  ENDIF.

  lv_column_start_int = zcl_excel_common=>convert_column2int( lv_column_start ).
  lv_column_end_int   = zcl_excel_common=>convert_column2int( lv_column_end ).

  IF lv_column_start_int > lv_column_end_int OR lv_row > lv_row_end.

    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Wrong Merging Parameters'.

  ENDIF.

  IF ip_data_type IS SUPPLIED OR
     ip_abap_type IS SUPPLIED.

    me->set_cell( ip_column    = lv_column_start
                  ip_row       = lv_row
                  ip_value     = ip_value
                  ip_formula   = ip_formula
                  ip_style     = ip_style
                  ip_hyperlink = ip_hyperlink
                  ip_data_type = ip_data_type
                  ip_abap_type = ip_abap_type ).

  ELSE.

    me->set_cell( ip_column    = lv_column_start
                  ip_row       = lv_row
                  ip_value     = ip_value
                  ip_formula   = ip_formula
                  ip_style     = ip_style
                  ip_hyperlink = ip_hyperlink ).

  ENDIF.

  IF ip_style IS SUPPLIED.

    me->set_area_style( ip_column_start = lv_column_start
                        ip_column_end   = lv_column_end
                        ip_row          = lv_row
                        ip_row_to       = lv_row_end
                        ip_style        = ip_style ).
  ENDIF.

  IF ip_merge IS SUPPLIED AND ip_merge = abap_true.

    me->set_merge( ip_column_start = lv_column_start
                   ip_column_end   = lv_column_end
                   ip_row          = lv_row
                   ip_row_to       = lv_row_end ).

  ENDIF.

ENDMETHOD.


METHOD set_area_formula.
  DATA: ld_row TYPE zexcel_cell_row,
        ld_row_end TYPE zexcel_cell_row,
        ld_column TYPE zexcel_cell_column_alpha,
        ld_column_end TYPE zexcel_cell_column_alpha,
        ld_column_int TYPE zexcel_cell_column_alpha,
        ld_column_end_int TYPE zexcel_cell_column_alpha.

  MOVE: ip_row_to TO ld_row_end,
        ip_row    TO ld_row.
  IF ld_row_end IS INITIAL or ip_row_to is not supplied.
    ld_row_end = ld_row.
  ENDIF.

  MOVE: ip_column_start TO ld_column,
        ip_column_end   TO ld_column_end.

  if ld_column_end is initial or ip_column_end is not supplied.
    ld_column_end = ld_column.
  endif.

  ld_column_int      = zcl_excel_common=>convert_column2int( ld_column ).
  ld_column_end_int  = zcl_excel_common=>convert_column2int( ld_column_end ).

  if ld_column_int > ld_column_end_int or ld_row > ld_row_end.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Wrong Merging Parameters'.
  endif.

  me->set_cell_formula( ip_column = ld_column ip_row = ld_row
                        ip_formula = ip_formula ).

  IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
    me->set_merge( ip_column_start = ld_column ip_row = ld_row
                   ip_column_end   = ld_column_end   ip_row_to = ld_row_end ).
  ENDIF.
ENDMETHOD.


METHOD SET_AREA_STYLE.
  DATA: ld_row_start TYPE zexcel_cell_row,
      ld_row_end TYPE zexcel_cell_row,
      ld_column_start_int TYPE zexcel_cell_column,
      ld_column_end_int TYPE zexcel_cell_column,
      ld_current_column TYPE zexcel_cell_column_alpha,
      ld_current_row TYPE zexcel_cell_row.

  MOVE: ip_row_to TO ld_row_end,
        ip_row    TO ld_row_start.
  IF ld_row_end IS INITIAL or ip_row_to is not supplied.
    ld_row_end = ld_row_start.
  ENDIF.
  ld_column_start_int = zcl_excel_common=>convert_column2int( ip_column_start ).
  ld_column_end_int   = zcl_excel_common=>convert_column2int( ip_column_end ).
  IF ld_column_end_int IS INITIAL or ip_column_end is not supplied.
    ld_column_end_int = ld_column_start_int.
  ENDIF.

  WHILE ld_column_start_int <= ld_column_end_int.
    ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_start_int ).
    ld_current_row = ld_row_start.
    WHILE ld_current_row <= ld_row_end.
      me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                          ip_style = ip_style ).
      ADD 1 TO ld_current_row.
    ENDWHILE.
    ADD 1 TO ld_column_start_int.
  ENDWHILE.
  IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
    me->set_merge( ip_column_start = ip_column_start ip_row = ld_row_start
                   ip_column_end   = ld_current_column    ip_row_to = ld_row_end ).
  ENDIF.
ENDMETHOD.


  method SET_CELL.

    data: LV_COLUMN        type ZEXCEL_CELL_COLUMN,
          LS_SHEET_CONTENT type ZEXCEL_S_CELL_DATA,
          LV_ROW_ALPHA     type STRING,
          LV_COL_ALPHA     type ZEXCEL_CELL_COLUMN_ALPHA,
          LV_VALUE         type ZEXCEL_CELL_VALUE,
          LV_DATA_TYPE     type ZEXCEL_CELL_DATA_TYPE,
          LV_VALUE_TYPE    type ABAP_TYPEKIND,
          LV_STYLE_GUID    type ZEXCEL_CELL_STYLE,
          LO_ADDIT         type ref to CL_ABAP_ELEMDESCR,
          LO_VALUE         type ref to DATA,
          LO_VALUE_NEW     type ref to DATA.

    field-symbols: <FS_SHEET_CONTENT> type ZEXCEL_S_CELL_DATA,
                   <FS_NUMERIC>       type NUMERIC,
                   <FS_DATE>          type D,
                   <FS_TIME>          type T,
                   <FS_VALUE>         type SIMPLE.

    if IP_VALUE  is not supplied and IP_FORMULA is not supplied.
      zcx_excel=>raise_text( 'Please provide the value or formula' ).
    endif.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
*  lv_style_guid = ip_style.
    LV_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN ).
    read table SHEET_CONTENT assigning <FS_SHEET_CONTENT> with table key CELL_ROW    = IP_ROW      " Changed to access via table key , Stefan Schmöcker, 2013-08-03
                                                                         CELL_COLUMN = LV_COLUMN.
    if SY-SUBRC = 0.
      if IP_STYLE is initial.
        " If no style is provided as method-parameter and cell is found use cell's current style
        LV_STYLE_GUID = <FS_SHEET_CONTENT>-CELL_STYLE.
      else.
        " Style provided as method-parameter --> use this
        LV_STYLE_GUID = IP_STYLE.
      endif.
    else.
      " No cell found --> use supplied style even if empty
      LV_STYLE_GUID = IP_STYLE.
    endif.
* End of change issue #152 - don't touch exisiting style if only value is passed

    if IP_VALUE is supplied.
      "if data type is passed just write the value. Otherwise map abap type to excel and perform conversion
      "IP_DATA_TYPE is passed by excel reader so source types are preserved
*First we get reference into local var.
      create data LO_VALUE like IP_VALUE.
      assign LO_VALUE->* to <FS_VALUE>.
      <FS_VALUE> = IP_VALUE.
      if IP_DATA_TYPE is supplied.
        if IP_ABAP_TYPE is not supplied.
          GET_VALUE_TYPE( exporting IP_VALUE      = IP_VALUE
                          importing EP_VALUE      = <FS_VALUE> ) .
        endif.
        LV_VALUE = <FS_VALUE>.
        LV_DATA_TYPE = IP_DATA_TYPE.
      else.
        if IP_ABAP_TYPE is supplied.
          LV_VALUE_TYPE = IP_ABAP_TYPE.
        else.
          GET_VALUE_TYPE( exporting IP_VALUE      = IP_VALUE
                          importing EP_VALUE      = <FS_VALUE>
                                    EP_VALUE_TYPE = LV_VALUE_TYPE ).
        endif.
        case LV_VALUE_TYPE.
          when CL_ABAP_TYPEDESCR=>TYPEKIND_INT or CL_ABAP_TYPEDESCR=>TYPEKIND_INT1 or CL_ABAP_TYPEDESCR=>TYPEKIND_INT2.
            LO_ADDIT = CL_ABAP_ELEMDESCR=>GET_I( ).
            create data LO_VALUE_NEW type handle LO_ADDIT.
            assign LO_VALUE_NEW->* to <FS_NUMERIC>.
            if SY-SUBRC = 0.
              <FS_NUMERIC> = <FS_VALUE>.
              LV_VALUE = ZCL_EXCEL_COMMON=>NUMBER_TO_EXCEL_STRING( IP_VALUE = <FS_NUMERIC> ).
            endif.

          when CL_ABAP_TYPEDESCR=>TYPEKIND_FLOAT or CL_ABAP_TYPEDESCR=>TYPEKIND_PACKED.
            LO_ADDIT = CL_ABAP_ELEMDESCR=>GET_F( ).
            create data LO_VALUE_NEW type handle LO_ADDIT.
            assign LO_VALUE_NEW->* to <FS_NUMERIC>.
            if SY-SUBRC = 0.
              <FS_NUMERIC> = <FS_VALUE>.
              LV_VALUE = ZCL_EXCEL_COMMON=>NUMBER_TO_EXCEL_STRING( IP_VALUE = <FS_NUMERIC> ).
            endif.

          when CL_ABAP_TYPEDESCR=>TYPEKIND_CHAR or CL_ABAP_TYPEDESCR=>TYPEKIND_STRING or CL_ABAP_TYPEDESCR=>TYPEKIND_NUM or
               CL_ABAP_TYPEDESCR=>TYPEKIND_HEX.
            LV_VALUE = <FS_VALUE>.
            LV_DATA_TYPE = 's'.

          when CL_ABAP_TYPEDESCR=>TYPEKIND_DATE.
            LO_ADDIT = CL_ABAP_ELEMDESCR=>GET_D( ).
            create data LO_VALUE_NEW type handle LO_ADDIT.
            assign LO_VALUE_NEW->* to <FS_DATE>.
            if SY-SUBRC = 0.
              <FS_DATE> = <FS_VALUE>.
              LV_VALUE = ZCL_EXCEL_COMMON=>DATE_TO_EXCEL_STRING( IP_VALUE = <FS_DATE> ) .
            endif.
* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Moved to end of routine - apply date-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default date format in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = get_default_excel_date_format( ).
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

          when CL_ABAP_TYPEDESCR=>TYPEKIND_TIME.
            LO_ADDIT = CL_ABAP_ELEMDESCR=>GET_T( ).
            create data LO_VALUE_NEW type handle LO_ADDIT.
            assign LO_VALUE_NEW->* to <FS_TIME>.
            if SY-SUBRC = 0.
              <FS_TIME> = <FS_VALUE>.
              LV_VALUE = ZCL_EXCEL_COMMON=>TIME_TO_EXCEL_STRING( IP_VALUE = <FS_TIME> ).
            endif.
* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Moved to end of routine - apply time-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default time format for user in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_time6.
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

          when others.
            zcx_excel=>raise_text( 'Invalid data type of input value' ).
        endcase.
      endif.

    endif.

    if IP_HYPERLINK is bound.
      IP_HYPERLINK->SET_CELL_REFERENCE( IP_COLUMN = IP_COLUMN
                                        IP_ROW = IP_ROW ).
      ME->HYPERLINKS->ADD( IP_HYPERLINK ).
    endif.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Read table moved up, so that current style may be evaluated
*  lv_column = zcl_excel_common=>convert_column2int( ip_column ).

*  READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH KEY cell_row    = ip_row
*                                                                 cell_column = lv_column.
*
*  IF sy-subrc EQ 0.
    if <FS_SHEET_CONTENT> is assigned.
* End of change issue #152 - don't touch exisiting style if only value is passed
      <FS_SHEET_CONTENT>-CELL_VALUE   = LV_VALUE.
      <FS_SHEET_CONTENT>-CELL_FORMULA = IP_FORMULA.
      <FS_SHEET_CONTENT>-CELL_STYLE   = LV_STYLE_GUID.
      <FS_SHEET_CONTENT>-DATA_TYPE    = LV_DATA_TYPE.
    else.
      LS_SHEET_CONTENT-CELL_ROW     = IP_ROW.
      LS_SHEET_CONTENT-CELL_COLUMN  = LV_COLUMN.
      LS_SHEET_CONTENT-CELL_VALUE   = LV_VALUE.
      LS_SHEET_CONTENT-CELL_FORMULA = IP_FORMULA.
      LS_SHEET_CONTENT-CELL_STYLE   = LV_STYLE_GUID.
      LS_SHEET_CONTENT-DATA_TYPE    = LV_DATA_TYPE.
      LV_ROW_ALPHA = IP_ROW.
*    SHIFT lv_row_alpha RIGHT DELETING TRAILING space."del #152 - replaced with condense - should be faster
*    SHIFT lv_row_alpha LEFT DELETING LEADING space.  "del #152 - replaced with condense - should be faster
      condense LV_ROW_ALPHA no-gaps.                    "ins #152 - replaced 2 shifts      - should be faster
      LV_COL_ALPHA = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( IP_COLUMN ).       " issue #155 - less restrictive typing for ip_column
      concatenate LV_COL_ALPHA LV_ROW_ALPHA into LS_SHEET_CONTENT-CELL_COORDS.  " issue #155 - less restrictive typing for ip_column
      insert LS_SHEET_CONTENT into table SHEET_CONTENT assigning <FS_SHEET_CONTENT>. "ins #152 - Now <fs_sheet_content> always holds the data
*    APPEND ls_sheet_content TO sheet_content.
*    SORT sheet_content BY cell_row cell_column.
      " me->update_dimension_range( ).

    endif.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
* For Date- or Timefields change the formatcode if nothing is set yet
* Enhancement option:  Check if existing formatcode is a date/ or timeformat
*                      If not, use default
    data: LO_FORMAT_CODE_DATETIME type ZEXCEL_NUMBER_FORMAT.
    data: STYLEMAPPING    type ZEXCEL_S_STYLEMAPPING.
    case LV_VALUE_TYPE.
      when CL_ABAP_TYPEDESCR=>TYPEKIND_DATE.
        try.
            STYLEMAPPING = ME->EXCEL->GET_STYLE_TO_GUID( <FS_SHEET_CONTENT>-CELL_STYLE ).
          catch ZCX_EXCEL .
        endtry.
        if STYLEMAPPING-COMPLETE_STYLEX-NUMBER_FORMAT-FORMAT_CODE is initial or
           STYLEMAPPING-COMPLETE_STYLE-NUMBER_FORMAT-FORMAT_CODE is initial.
          LO_FORMAT_CODE_DATETIME = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_STD.
        else.
          LO_FORMAT_CODE_DATETIME = STYLEMAPPING-COMPLETE_STYLE-NUMBER_FORMAT-FORMAT_CODE.
        endif.
        ME->CHANGE_CELL_STYLE( IP_COLUMN                      = IP_COLUMN
                               IP_ROW                         = IP_ROW
                               IP_NUMBER_FORMAT_FORMAT_CODE   = LO_FORMAT_CODE_DATETIME ).

      when CL_ABAP_TYPEDESCR=>TYPEKIND_TIME.
        try.
            STYLEMAPPING = ME->EXCEL->GET_STYLE_TO_GUID( <FS_SHEET_CONTENT>-CELL_STYLE ).
          catch ZCX_EXCEL .
        endtry.
        if STYLEMAPPING-COMPLETE_STYLEX-NUMBER_FORMAT-FORMAT_CODE is initial or
           STYLEMAPPING-COMPLETE_STYLE-NUMBER_FORMAT-FORMAT_CODE is initial.
          LO_FORMAT_CODE_DATETIME = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>C_FORMAT_DATE_TIME6.
        else.
          LO_FORMAT_CODE_DATETIME = STYLEMAPPING-COMPLETE_STYLE-NUMBER_FORMAT-FORMAT_CODE.
        endif.
        ME->CHANGE_CELL_STYLE( IP_COLUMN                      = IP_COLUMN
                               IP_ROW                         = IP_ROW
                               IP_NUMBER_FORMAT_FORMAT_CODE   = LO_FORMAT_CODE_DATETIME ).

    endcase.
* End of change issue #152 - don't touch exisiting style if only value is passed

* Fix issue #162
    LV_VALUE = IP_VALUE.
    if LV_VALUE cs CL_ABAP_CHAR_UTILITIES=>CR_LF.
      ME->CHANGE_CELL_STYLE( IP_COLUMN               = IP_COLUMN
                             IP_ROW                  = IP_ROW
                             IP_ALIGNMENT_WRAPTEXT   = ABAP_TRUE ).
    endif.
* End of Fix issue #162

  endmethod.


  method SET_CELL_FORMULA.
    data:
      LV_COLUMN        type ZEXCEL_CELL_COLUMN,
      LS_SHEET_CONTENT like line of ME->SHEET_CONTENT.

    field-symbols:
                <SHEET_CONTENT>                 like line of ME->SHEET_CONTENT.

*--------------------------------------------------------------------*
* Get cell to set formula into
*--------------------------------------------------------------------*
    LV_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN ).
    read table ME->SHEET_CONTENT assigning <SHEET_CONTENT> with table key CELL_ROW    = IP_ROW
                                                                          CELL_COLUMN = LV_COLUMN.
    if SY-SUBRC <> 0.                                                                           " Create new entry in sheet_content if necessary
      check IP_FORMULA is initial.                                                              " no need to create new entry in sheet_content when no formula is passed
      LS_SHEET_CONTENT-CELL_ROW    = IP_ROW.
      LS_SHEET_CONTENT-CELL_COLUMN = LV_COLUMN.
      insert LS_SHEET_CONTENT into table ME->SHEET_CONTENT assigning <SHEET_CONTENT>.
    endif.

*--------------------------------------------------------------------*
* Fieldsymbol now holds the relevant cell
*--------------------------------------------------------------------*
    <SHEET_CONTENT>-CELL_FORMULA = IP_FORMULA.


  endmethod.


  method SET_CELL_STYLE.

    data: LV_COLUMN     type ZEXCEL_CELL_COLUMN,
          LV_STYLE_GUID type ZEXCEL_CELL_STYLE.

    field-symbols: <FS_SHEET_CONTENT> type ZEXCEL_S_CELL_DATA.

    LV_STYLE_GUID = IP_STYLE.

    LV_COLUMN = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN ).

    read table SHEET_CONTENT assigning <FS_SHEET_CONTENT> with key CELL_ROW    = IP_ROW
                                                                   CELL_COLUMN = LV_COLUMN.

    if SY-SUBRC eq 0.
      <FS_SHEET_CONTENT>-CELL_STYLE   = LV_STYLE_GUID.
    else.
      SET_CELL( IP_COLUMN = IP_COLUMN IP_ROW = IP_ROW IP_VALUE = '' IP_STYLE = IP_STYLE ).
    endif.

  endmethod.


  method SET_COLUMN_WIDTH.
    data: LO_COLUMN  type ref to ZCL_EXCEL_COLUMN.
    data: WIDTH             type FLOAT.

    LO_COLUMN = ME->GET_COLUMN( IP_COLUMN ).

* if a fix size is supplied use this
    if IP_WIDTH_FIX is supplied.
      try.
          WIDTH = IP_WIDTH_FIX.
          if WIDTH <= 0.
            zcx_excel=>raise_text( 'Please supply a positive number as column-width' ).
          endif.
          LO_COLUMN->SET_WIDTH( WIDTH ).
          exit.
        catch CX_SY_CONVERSION_NO_NUMBER.
* Strange stuff passed --> raise error
          zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
      endtry.
    endif.

* If we get down to here, we have to use whatever is found in autosize.
    LO_COLUMN->SET_AUTO_SIZE( IP_WIDTH_AUTOSIZE ).


  endmethod.


  method SET_DEFAULT_EXCEL_DATE_FORMAT.

    if IP_DEFAULT_EXCEL_DATE_FORMAT is initial.
      zcx_excel=>raise_text( 'Default date format cannot be blank' ).
    endif.

    DEFAULT_EXCEL_DATE_FORMAT = IP_DEFAULT_EXCEL_DATE_FORMAT.
  endmethod.


  method SET_MERGE.

    data: LS_MERGE        type MTY_MERGE,
          LV_ERRORMESSAGE type STRING.

...
"just after variables definition
if ip_value is supplied or ip_formula is supplied.
" if there is a value or formula set the value to the top-left cell
"maybe it is necessary to support other paramters for set_cell
    if ip_value is supplied.
      me->set_cell( ip_row = ip_row ip_column = ip_column_start
                    ip_value = ip_value ).
    endif.
    if ip_formula is supplied.
      me->set_cell( ip_row = ip_row ip_column = ip_column_start
                    ip_value = ip_formula ).
    endif.
  endif.
"call to set_merge_style to apply the style to all cells at the matrix
  IF ip_style IS SUPPLIED.
    me->set_merge_style( ip_row = ip_row ip_column_start = ip_column_start
                         ip_row_to = ip_row_to ip_column_end = ip_column_end
                         ip_style = ip_style ).
  ENDIF.
...
*--------------------------------------------------------------------*
* Build new range area to insert into range table
*--------------------------------------------------------------------*
    LS_MERGE-ROW_FROM = IP_ROW.
    if IP_ROW is supplied and IP_ROW is not initial and IP_ROW_TO is not supplied.
      LS_MERGE-ROW_TO   = LS_MERGE-ROW_FROM.
    else.
      LS_MERGE-ROW_TO   = IP_ROW_TO.
    endif.
    if LS_MERGE-ROW_FROM > LS_MERGE-ROW_TO.
      LV_ERRORMESSAGE = 'Merge: First row larger then last row'(405).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    LS_MERGE-COL_FROM = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN_START ).
    if IP_COLUMN_START is supplied and IP_COLUMN_START is not initial and IP_COLUMN_END is not supplied.
      LS_MERGE-COL_TO   = LS_MERGE-COL_FROM.
    else.
      LS_MERGE-COL_TO   = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_COLUMN_END ).
    endif.
    if LS_MERGE-COL_FROM > LS_MERGE-COL_TO.
      LV_ERRORMESSAGE = 'Merge: First column larger then last column'(406).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

*--------------------------------------------------------------------*
* Check merge not overlapping with existing merges
*--------------------------------------------------------------------*
    loop at ME->MT_MERGED_CELLS transporting no fields where not (    ROW_FROM > LS_MERGE-ROW_TO
                                                                   or ROW_TO   < LS_MERGE-ROW_FROM
                                                                   or COL_FROM > LS_MERGE-COL_TO
                                                                   or COL_TO   < LS_MERGE-COL_FROM ).
      LV_ERRORMESSAGE = 'Overlapping merges'(404).
      zcx_excel=>raise_text( lv_errormessage ).

    endloop.

*--------------------------------------------------------------------*
* Everything seems ok --> add to merge table
*--------------------------------------------------------------------*
    insert LS_MERGE into table ME->MT_MERGED_CELLS.

  endmethod.


METHOD set_merge_style.
  DATA: ld_row_start TYPE zexcel_cell_row,
        ld_row_end TYPE zexcel_cell_row,
        ld_column_start TYPE zexcel_cell_column,
        ld_column_end TYPE zexcel_cell_column,
        ld_current_column TYPE zexcel_cell_column_alpha,
        ld_current_row type zexcel_cell_row.

  MOVE: ip_row_to TO ld_row_end,
        ip_row    TO ld_row_start.
  IF ld_row_end IS INITIAL.
    ld_row_end = ld_row_start.
  ENDIF.
  ld_column_start = zcl_excel_common=>convert_column2int( ip_column_start ).
  ld_column_end   = zcl_excel_common=>convert_column2int( ip_column_end ).
  IF ld_column_end IS INITIAL.
    ld_column_end = ld_column_start.
  ENDIF.
  "set the style cell by cell
  WHILE ld_column_start <= ld_column_end.
    ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_start ).
    ld_current_row = ld_row_start.
    WHILE ld_current_row <= ld_row_end.
      me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                          ip_style = ip_style ).
      add 1 to ld_current_row.
    ENDWHILE.
    ADD 1 TO ld_column_start.
  ENDWHILE.
ENDMETHOD.


  method SET_PRINT_GRIDLINES.
    ME->PRINT_GRIDLINES = I_PRINT_GRIDLINES.
  endmethod.


  method SET_ROW_HEIGHT.
    data: LO_ROW  type ref to ZCL_EXCEL_ROW.
    data: HEIGHT  type FLOAT.

    LO_ROW = ME->GET_ROW( IP_ROW ).

* if a fix size is supplied use this
    try.
        HEIGHT = IP_HEIGHT_FIX.
        if HEIGHT <= 0.
          zcx_excel=>raise_text( 'Please supply a positive number as row-height' ).
        endif.
        LO_ROW->SET_ROW_HEIGHT( HEIGHT ).
        exit.
      catch CX_SY_CONVERSION_NO_NUMBER.
* Strange stuff passed --> raise error
        zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
    endtry.

  endmethod.


  method SET_ROW_OUTLINE.

    data: LS_ROW_OUTLINE like line of ME->MT_ROW_OUTLINES.
    field-symbols: <LS_ROW_OUTLINE> like line of ME->MT_ROW_OUTLINES.

    read table ME->MT_ROW_OUTLINES assigning <LS_ROW_OUTLINE> with table key ROW_FROM = IV_ROW_FROM
                                                                             ROW_TO   = IV_ROW_TO.
    if SY-SUBRC <> 0.
      if IV_ROW_FROM <= 0.
        zcx_excel=>raise_text( 'First row of outline must be a positive number' ).
      endif.
      if IV_ROW_TO < IV_ROW_FROM.
        zcx_excel=>raise_text( 'Last row of outline may not be less than first line of outline' ).
      endif.
      LS_ROW_OUTLINE-ROW_FROM = IV_ROW_FROM.
      LS_ROW_OUTLINE-ROW_TO   = IV_ROW_TO.
      insert LS_ROW_OUTLINE into table ME->MT_ROW_OUTLINES assigning <LS_ROW_OUTLINE>.
    endif.

    case IV_COLLAPSED.

      when ABAP_TRUE
        or ABAP_FALSE.
        <LS_ROW_OUTLINE>-COLLAPSED = IV_COLLAPSED.

      when others.
        zcx_excel=>raise_text( 'Unknown collapse state' ).

    endcase.
  endmethod.


  method SET_SHOW_GRIDLINES.
    ME->SHOW_GRIDLINES = I_SHOW_GRIDLINES.
  endmethod.


  method SET_SHOW_ROWCOLHEADERS.
    ME->SHOW_ROWCOLHEADERS = I_SHOW_ROWCOLHEADERS.
  endmethod.


  method SET_TABCOLOR.
    ME->TABCOLOR = IV_TABCOLOR.
  endmethod.


  method SET_TABLE.

    data: LO_TABDESCR     type ref to CL_ABAP_STRUCTDESCR,
          LR_DATA         type ref to DATA,
          LS_HEADER       type X030L,
          LT_DFIES        type DDFIELDS,
          LV_ROW_INT      type ZEXCEL_CELL_ROW,
          LV_COLUMN_INT   type ZEXCEL_CELL_COLUMN,
          LV_COLUMN_ALPHA type ZEXCEL_CELL_COLUMN_ALPHA,
          LV_CELL_VALUE   type ZEXCEL_CELL_VALUE.


    field-symbols: <FS_TABLE_LINE> type ANY,
                   <FS_FLDVAL>     type ANY,
                   <FS_DFIES>      type DFIES.

    LV_COLUMN_INT = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_TOP_LEFT_COLUMN ).
    LV_ROW_INT    = IP_TOP_LEFT_ROW.

    create data LR_DATA like line of IP_TABLE.

    LO_TABDESCR ?= CL_ABAP_STRUCTDESCR=>DESCRIBE_BY_DATA_REF( LR_DATA ).

    LS_HEADER = LO_TABDESCR->GET_DDIC_HEADER( ).

    LT_DFIES = LO_TABDESCR->GET_DDIC_FIELD_LIST( ).

* It is better to loop column by column
    loop at LT_DFIES assigning <FS_DFIES>.
      LV_COLUMN_ALPHA = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( LV_COLUMN_INT ).

      if IP_NO_HEADER = ABAP_FALSE.
        " First of all write column header
        LV_CELL_VALUE = <FS_DFIES>-SCRTEXT_M.
        ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                      IP_ROW    = LV_ROW_INT
                      IP_VALUE  = LV_CELL_VALUE
                      IP_STYLE  = IP_HDR_STYLE ).
        if IP_TRANSPOSE = ABAP_TRUE.
          add 1 to LV_COLUMN_INT.
        else.
          add 1 to LV_ROW_INT.
        endif.
      endif.

      loop at IP_TABLE assigning <FS_TABLE_LINE>.
        LV_COLUMN_ALPHA = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( LV_COLUMN_INT ).
        assign component <FS_DFIES>-FIELDNAME of structure <FS_TABLE_LINE> to <FS_FLDVAL>.
        move <FS_FLDVAL> to LV_CELL_VALUE.
        ME->SET_CELL( IP_COLUMN = LV_COLUMN_ALPHA
                      IP_ROW    = LV_ROW_INT
                      IP_VALUE  = <FS_FLDVAL>   "lv_cell_value
                      IP_STYLE  = IP_BODY_STYLE ).
        if IP_TRANSPOSE = ABAP_TRUE.
          add 1 to LV_COLUMN_INT.
        else.
          add 1 to LV_ROW_INT.
        endif.
      endloop.
      if IP_TRANSPOSE = ABAP_TRUE.
        LV_COLUMN_INT = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IP_TOP_LEFT_COLUMN ).
        add 1 to LV_ROW_INT.
      else.
        LV_ROW_INT = IP_TOP_LEFT_ROW.
        add 1 to LV_COLUMN_INT.
      endif.
    endloop.

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
    data: LO_WORKSHEETS_ITERATOR type ref to CL_OBJECT_COLLECTION_ITERATOR,
          LO_WORKSHEET           type ref to ZCL_EXCEL_WORKSHEET,
          ERRORMESSAGE           type STRING,
          LV_RANGESHEETNAME_OLD  type STRING,
          LV_RANGESHEETNAME_NEW  type STRING,
          LO_RANGES_ITERATOR     type ref to CL_OBJECT_COLLECTION_ITERATOR,
          LO_RANGE               type ref to ZCL_EXCEL_RANGE,
          LV_RANGE_VALUE         type ZEXCEL_RANGE_VALUE,
          LV_ERRORMESSAGE        type STRING.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Check whether title consists only of allowed characters
* Illegal characters are: /	\	[	]	*	?	:  --> http://msdn.microsoft.com/en-us/library/ff837411.aspx
* Illegal characters not in documentation:   ' as first character
*--------------------------------------------------------------------*
    if IP_TITLE ca '/\[]*?:'.
      LV_ERRORMESSAGE = 'Found illegal character in sheetname. List of forbidden characters: /\[]*?:'(402).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    if IP_TITLE is not initial and IP_TITLE(1) = `'`.
      LV_ERRORMESSAGE = 'Sheetname may not start with &'(403).   " & used instead of ' to allow fallbacklanguage
      replace '&' in LV_ERRORMESSAGE with `'`.
      zcx_excel=>raise_text( lv_errormessage ).
    endif.


*--------------------------------------------------------------------*
* Check whether title is unique in workbook
*--------------------------------------------------------------------*
    LO_WORKSHEETS_ITERATOR = ME->EXCEL->GET_WORKSHEETS_ITERATOR( ).
    while LO_WORKSHEETS_ITERATOR->HAS_NEXT( ) = 'X'.

      LO_WORKSHEET ?= LO_WORKSHEETS_ITERATOR->GET_NEXT( ).
      check ME->GUID <> LO_WORKSHEET->GET_GUID( ).  " Don't check against itself
      if IP_TITLE = LO_WORKSHEET->GET_TITLE( ).  " Not unique --> raise exception
        ERRORMESSAGE = 'Duplicate sheetname &'.
        replace '&' in ERRORMESSAGE with IP_TITLE.
        zcx_excel=>raise_text( errormessage ).
      endif.

    endwhile.

*--------------------------------------------------------------------*
* Remember old sheetname and rename sheet to desired name
*--------------------------------------------------------------------*
    concatenate ME->TITLE '!' into LV_RANGESHEETNAME_OLD.
    ME->TITLE = IP_TITLE.

*--------------------------------------------------------------------*
* After changing this worksheet's title we have to adjust
* all ranges that are referring to this worksheet.
*--------------------------------------------------------------------*
* 2do §1  -  Check if the following quickfix is solid
*           I fear it isn't - but this implementation is better then
*           nothing at all since it handles a supposed majority of cases
*--------------------------------------------------------------------*
    concatenate ME->TITLE '!' into LV_RANGESHEETNAME_NEW.

    LO_RANGES_ITERATOR = ME->EXCEL->GET_RANGES_ITERATOR( ).
    while LO_RANGES_ITERATOR->HAS_NEXT( ) = 'X'.

      LO_RANGE ?= LO_RANGES_ITERATOR->GET_NEXT( ).
      LV_RANGE_VALUE = LO_RANGE->GET_VALUE( ).
      replace all occurrences of LV_RANGESHEETNAME_OLD in LV_RANGE_VALUE with LV_RANGESHEETNAME_NEW.
      if SY-SUBRC = 0.
        LO_RANGE->SET_RANGE_VALUE( LV_RANGE_VALUE ).
      endif.

    endwhile.


  endmethod.


  method UPDATE_DIMENSION_RANGE.

    data: LS_SHEET_CONTENT type ZEXCEL_S_CELL_DATA,
          LV_ROW_ALPHA     type STRING,
          LV_COLUMN_ALPHA  type ZEXCEL_CELL_COLUMN_ALPHA.

    check SHEET_CONTENT is not initial.

    UPPER_CELL-CELL_ROW = ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MAX_ROW.
    UPPER_CELL-CELL_COLUMN = ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MAX_COL.

    LOWER_CELL-CELL_ROW = ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MIN_ROW.
    LOWER_CELL-CELL_COLUMN = ZCL_EXCEL_COMMON=>C_EXCEL_SHEET_MIN_COL.

    loop at SHEET_CONTENT into LS_SHEET_CONTENT.
      if UPPER_CELL-CELL_ROW > LS_SHEET_CONTENT-CELL_ROW.
        UPPER_CELL-CELL_ROW = LS_SHEET_CONTENT-CELL_ROW.
      endif.
      if UPPER_CELL-CELL_COLUMN > LS_SHEET_CONTENT-CELL_COLUMN.
        UPPER_CELL-CELL_COLUMN = LS_SHEET_CONTENT-CELL_COLUMN.
      endif.
      if LOWER_CELL-CELL_ROW < LS_SHEET_CONTENT-CELL_ROW.
        LOWER_CELL-CELL_ROW = LS_SHEET_CONTENT-CELL_ROW.
      endif.
      if LOWER_CELL-CELL_COLUMN < LS_SHEET_CONTENT-CELL_COLUMN.
        LOWER_CELL-CELL_COLUMN = LS_SHEET_CONTENT-CELL_COLUMN.
      endif.
    endloop.

    LV_ROW_ALPHA = UPPER_CELL-CELL_ROW.
    LV_COLUMN_ALPHA = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( UPPER_CELL-CELL_COLUMN ).
    shift LV_ROW_ALPHA right deleting trailing SPACE.
    shift LV_ROW_ALPHA left deleting leading SPACE.
    concatenate LV_COLUMN_ALPHA LV_ROW_ALPHA into UPPER_CELL-CELL_COORDS.

    LV_ROW_ALPHA = LOWER_CELL-CELL_ROW.
    LV_COLUMN_ALPHA = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2ALPHA( LOWER_CELL-CELL_COLUMN ).
    shift LV_ROW_ALPHA right deleting trailing SPACE.
    shift LV_ROW_ALPHA left deleting leading SPACE.
    concatenate LV_COLUMN_ALPHA LV_ROW_ALPHA into LOWER_CELL-CELL_COORDS.

  endmethod.


  method ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_COLUMNS.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    clear:  ME->PRINT_TITLE_COL_FROM,
            ME->PRINT_TITLE_COL_TO  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    ME->PRINT_TITLE_SET_RANGE( ).


  endmethod.


  method ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_ROWS.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    clear:  ME->PRINT_TITLE_ROW_FROM,
            ME->PRINT_TITLE_ROW_TO  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    ME->PRINT_TITLE_SET_RANGE( ).


  endmethod.


  method ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_COLUMNS.
    EV_COLUMNS_FROM = ME->PRINT_TITLE_COL_FROM.
    EV_COLUMNS_TO   = ME->PRINT_TITLE_COL_TO.
  endmethod.


  method ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_ROWS.
    EV_ROWS_FROM = ME->PRINT_TITLE_ROW_FROM.
    EV_ROWS_TO   = ME->PRINT_TITLE_ROW_TO.
  endmethod.


  method ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_COLUMNS.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    data: LV_COL_FROM_INT type I,
          LV_COL_TO_INT   type I,
          LV_ERRORMESSAGE type STRING.


    LV_COL_FROM_INT = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IV_COLUMNS_FROM ).
    LV_COL_TO_INT   = ZCL_EXCEL_COMMON=>CONVERT_COLUMN2INT( IV_COLUMNS_TO ).

*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    if LV_COL_FROM_INT < 1.
      LV_ERRORMESSAGE = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    if  LV_COL_FROM_INT > LV_COL_TO_INT.
      LV_ERRORMESSAGE = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    ME->PRINT_TITLE_COL_FROM = IV_COLUMNS_FROM.
    ME->PRINT_TITLE_COL_TO   = IV_COLUMNS_TO.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    ME->PRINT_TITLE_SET_RANGE( ).

  endmethod.


  method ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_ROWS.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    data:     LV_ERRORMESSAGE                 type STRING.


*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    if IV_ROWS_FROM < 1.
      LV_ERRORMESSAGE = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    if  IV_ROWS_FROM > IV_ROWS_TO.
      LV_ERRORMESSAGE = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    ME->PRINT_TITLE_ROW_FROM = IV_ROWS_FROM.
    ME->PRINT_TITLE_ROW_TO   = IV_ROWS_TO.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    ME->PRINT_TITLE_SET_RANGE( ).


  endmethod.


  method ZIF_EXCEL_SHEET_PROPERTIES~GET_STYLE.
    if ZIF_EXCEL_SHEET_PROPERTIES~STYLE is not initial.
      EP_STYLE = ZIF_EXCEL_SHEET_PROPERTIES~STYLE.
    else.
      EP_STYLE = ME->EXCEL->GET_DEFAULT_STYLE( ).
    endif.
  endmethod.


  method ZIF_EXCEL_SHEET_PROPERTIES~INITIALIZE.

    ZIF_EXCEL_SHEET_PROPERTIES~SHOW_ZEROS   = ZIF_EXCEL_SHEET_PROPERTIES=>C_SHOWZERO.
    ZIF_EXCEL_SHEET_PROPERTIES~SUMMARYBELOW = ZIF_EXCEL_SHEET_PROPERTIES=>C_BELOW_ON.
    ZIF_EXCEL_SHEET_PROPERTIES~SUMMARYRIGHT = ZIF_EXCEL_SHEET_PROPERTIES=>C_RIGHT_ON.

* inizialize zoomscale values
    ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE = 100.
    ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE_NORMAL = 100.
    ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE_PAGELAYOUTVIEW = 100 .
    ZIF_EXCEL_SHEET_PROPERTIES~ZOOMSCALE_SHEETLAYOUTVIEW = 100 .
  endmethod.


  method ZIF_EXCEL_SHEET_PROPERTIES~SET_STYLE.
    ZIF_EXCEL_SHEET_PROPERTIES~STYLE = IP_STYLE.
  endmethod.


  method ZIF_EXCEL_SHEET_PROTECTION~INITIALIZE.

    ME->ZIF_EXCEL_SHEET_PROTECTION~PROTECTED = ZIF_EXCEL_SHEET_PROTECTION=>C_UNPROTECTED.
    clear ME->ZIF_EXCEL_SHEET_PROTECTION~PASSWORD.
    ME->ZIF_EXCEL_SHEET_PROTECTION~AUTO_FILTER            = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~DELETE_COLUMNS         = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~DELETE_ROWS            = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~FORMAT_CELLS           = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~FORMAT_COLUMNS         = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~FORMAT_ROWS            = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~INSERT_COLUMNS         = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~INSERT_HYPERLINKS      = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~INSERT_ROWS            = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~OBJECTS                = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
*  me->zif_excel_sheet_protection~password               = zif_excel_sheet_protection=>c_noactive. "issue #68
    ME->ZIF_EXCEL_SHEET_PROTECTION~PIVOT_TABLES           = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~PROTECTED              = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~SCENARIOS              = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~SELECT_LOCKED_CELLS    = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~SELECT_UNLOCKED_CELLS  = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~SHEET                  = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.
    ME->ZIF_EXCEL_SHEET_PROTECTION~SORT                   = ZIF_EXCEL_SHEET_PROTECTION=>C_NOACTIVE.

  endmethod.


  method ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME.
    ME->ZIF_EXCEL_SHEET_VBA_PROJECT~CODENAME = IP_CODENAME.
  endmethod.


  method ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME_PR.
    ME->ZIF_EXCEL_SHEET_VBA_PROJECT~CODENAME_PR = IP_CODENAME_PR.
  endmethod.
ENDCLASS.
