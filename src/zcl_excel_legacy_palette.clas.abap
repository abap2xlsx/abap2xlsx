class ZCL_EXCEL_LEGACY_PALETTE definition
  public
  create public .

public section.
*"* public components of class ZCL_EXCEL_LEGACY_PALETTE
*"* do not include other source files here!!!
  type-pools ABAP .

  methods CONSTRUCTOR .
  methods IS_MODIFIED
    returning
      value(EP_MODIFIED) type ABAP_BOOL .
  methods GET_COLOR
    importing
      !IP_INDEX type I
    returning
      value(EP_COLOR) type ZEXCEL_STYLE_COLOR_ARGB
    raising
      ZCX_EXCEL .
  methods GET_COLORS
    returning
      value(EP_COLORS) type ZEXCEL_T_STYLE_COLOR_ARGB .
  methods SET_COLOR
    importing
      !IP_INDEX type I
      !IP_COLOR type ZEXCEL_STYLE_COLOR_ARGB .
protected section.
*"* protected components of class ZCL_EXCEL_LEGACY_PALETTE
*"* do not include other source files here!!!
private section.

*"* private components of class ZCL_EXCEL_LEGACY_PALETTE
*"* do not include other source files here!!!
  data MODIFIED type ABAP_BOOL value ABAP_FALSE. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data COLORS type ZEXCEL_T_STYLE_COLOR_ARGB .
ENDCLASS.



CLASS ZCL_EXCEL_LEGACY_PALETTE IMPLEMENTATION.


method CONSTRUCTOR.
  " default Excel palette based on
  " http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.indexedcolors.aspx

  APPEND '00000000' TO colors.
  APPEND '00FFFFFF' TO colors.
  APPEND '00FF0000' TO colors.
  APPEND '0000FF00' TO colors.
  APPEND '000000FF' TO colors.
  APPEND '00FFFF00' TO colors.
  APPEND '00FF00FF' TO colors.
  APPEND '0000FFFF' TO colors.
  APPEND '00000000' TO colors.
  APPEND '00FFFFFF' TO colors.

  APPEND '00FF0000' TO colors.
  APPEND '0000FF00' TO colors.
  APPEND '000000FF' TO colors.
  APPEND '00FFFF00' TO colors.
  APPEND '00FF00FF' TO colors.
  APPEND '0000FFFF' TO colors.
  APPEND '00800000' TO colors.
  APPEND '00008000' TO colors.
  APPEND '00000080' TO colors.
  APPEND '00808000' TO colors.

  APPEND '00800080' TO colors.
  APPEND '00008080' TO colors.
  APPEND '00C0C0C0' TO colors.
  APPEND '00808080' TO colors.
  APPEND '009999FF' TO colors.
  APPEND '00993366' TO colors.
  APPEND '00FFFFCC' TO colors.
  APPEND '00CCFFFF' TO colors.
  APPEND '00660066' TO colors.
  APPEND '00FF8080' TO colors.

  APPEND '000066CC' TO colors.
  APPEND '00CCCCFF' TO colors.
  APPEND '00000080' TO colors.
  APPEND '00FF00FF' TO colors.
  APPEND '00FFFF00' TO colors.
  APPEND '0000FFFF' TO colors.
  APPEND '00800080' TO colors.
  APPEND '00800000' TO colors.
  APPEND '00008080' TO colors.
  APPEND '000000FF' TO colors.

  APPEND '0000CCFF' TO colors.
  APPEND '00CCFFFF' TO colors.
  APPEND '00CCFFCC' TO colors.
  APPEND '00FFFF99' TO colors.
  APPEND '0099CCFF' TO colors.
  APPEND '00FF99CC' TO colors.
  APPEND '00CC99FF' TO colors.
  APPEND '00FFCC99' TO colors.
  APPEND '003366FF' TO colors.
  APPEND '0033CCCC' TO colors.

  APPEND '0099CC00' TO colors.
  APPEND '00FFCC00' TO colors.
  APPEND '00FF9900' TO colors.
  APPEND '00FF6600' TO colors.
  APPEND '00666699' TO colors.
  APPEND '00969696' TO colors.
  APPEND '00003366' TO colors.
  APPEND '00339966' TO colors.
  APPEND '00003300' TO colors.
  APPEND '00333300' TO colors.

  APPEND '00993300' TO colors.
  APPEND '00993366' TO colors.
  APPEND '00333399' TO colors.
  APPEND '00333333' TO colors.

  endmethod.


method GET_COLOR.
  DATA: lv_index type i.

  lv_index = ip_index + 1.
  READ TABLE colors INTO ep_color INDEX lv_index.
  IF sy-subrc <> 0.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Invalid color index'.
  ENDIF.
  endmethod.


method GET_COLORS.
  ep_colors = colors.
  endmethod.


method IS_MODIFIED.
  ep_modified = modified.
  endmethod.


method SET_COLOR.
  DATA: lv_index TYPE i.

  FIELD-SYMBOLS: <lv_color> LIKE LINE OF colors.

  lv_index = ip_index + 1.
  READ TABLE colors ASSIGNING <lv_color> INDEX lv_index.
  IF sy-subrc <> 0.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Invalid color index'.
  ENDIF.

  IF <lv_color> <> ip_color.
    modified = abap_true.
    <lv_color> = ip_color.
  ENDIF.

  endmethod.
ENDCLASS.
