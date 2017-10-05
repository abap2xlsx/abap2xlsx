class ZCL_EXCEL_STYLE_BORDER definition
  public
  final
  create public .

public section.

*"* public components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
  data BORDER_STYLE type ZEXCEL_BORDER .
  data BORDER_COLOR type ZEXCEL_S_STYLE_COLOR .
  constants C_BORDER_NONE type ZEXCEL_BORDER value 'none'. "#EC NOTEXT
  constants C_BORDER_DASHDOT type ZEXCEL_BORDER value 'dashDot'. "#EC NOTEXT
  constants C_BORDER_DASHDOTDOT type ZEXCEL_BORDER value 'dashDotDot'. "#EC NOTEXT
  constants C_BORDER_DASHED type ZEXCEL_BORDER value 'dashed'. "#EC NOTEXT
  constants C_BORDER_DOTTED type ZEXCEL_BORDER value 'dotted'. "#EC NOTEXT
  constants C_BORDER_DOUBLE type ZEXCEL_BORDER value 'double'. "#EC NOTEXT
  constants C_BORDER_HAIR type ZEXCEL_BORDER value 'hair'. "#EC NOTEXT
  constants C_BORDER_MEDIUM type ZEXCEL_BORDER value 'medium'. "#EC NOTEXT
  constants C_BORDER_MEDIUMDASHDOT type ZEXCEL_BORDER value 'mediumDashDot'. "#EC NOTEXT
  constants C_BORDER_MEDIUMDASHDOTDOT type ZEXCEL_BORDER value 'mediumDashDotDot'. "#EC NOTEXT
  constants C_BORDER_MEDIUMDASHED type ZEXCEL_BORDER value 'mediumDashed'. "#EC NOTEXT
  constants C_BORDER_SLANTDASHDOT type ZEXCEL_BORDER value 'slantDashDot'. "#EC NOTEXT
  constants C_BORDER_THICK type ZEXCEL_BORDER value 'thick'. "#EC NOTEXT
  constants C_BORDER_THIN type ZEXCEL_BORDER value 'thin'. "#EC NOTEXT

  methods CONSTRUCTOR .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
private section.
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_BORDER IMPLEMENTATION.


method CONSTRUCTOR.
  border_style = zcl_excel_style_border=>c_border_none.
  border_color-theme     = zcl_excel_style_color=>c_theme_not_set.
  border_color-indexed   = zcl_excel_style_color=>c_indexed_not_set.
  endmethod.
ENDCLASS.
