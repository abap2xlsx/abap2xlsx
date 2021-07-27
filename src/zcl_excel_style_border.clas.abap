CLASS zcl_excel_style_border DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
    DATA border_style TYPE zexcel_border .
    DATA border_color TYPE zexcel_s_style_color .
    CONSTANTS c_border_none TYPE zexcel_border VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_border_dashdot TYPE zexcel_border VALUE 'dashDot'. "#EC NOTEXT
    CONSTANTS c_border_dashdotdot TYPE zexcel_border VALUE 'dashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_dashed TYPE zexcel_border VALUE 'dashed'. "#EC NOTEXT
    CONSTANTS c_border_dotted TYPE zexcel_border VALUE 'dotted'. "#EC NOTEXT
    CONSTANTS c_border_double TYPE zexcel_border VALUE 'double'. "#EC NOTEXT
    CONSTANTS c_border_hair TYPE zexcel_border VALUE 'hair'. "#EC NOTEXT
    CONSTANTS c_border_medium TYPE zexcel_border VALUE 'medium'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdot TYPE zexcel_border VALUE 'mediumDashDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdotdot TYPE zexcel_border VALUE 'mediumDashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashed TYPE zexcel_border VALUE 'mediumDashed'. "#EC NOTEXT
    CONSTANTS c_border_slantdashdot TYPE zexcel_border VALUE 'slantDashDot'. "#EC NOTEXT
    CONSTANTS c_border_thick TYPE zexcel_border VALUE 'thick'. "#EC NOTEXT
    CONSTANTS c_border_thin TYPE zexcel_border VALUE 'thin'. "#EC NOTEXT

    METHODS constructor .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_style_border IMPLEMENTATION.


  METHOD constructor.
    border_style = zcl_excel_style_border=>c_border_none.
    border_color-theme     = zcl_excel_style_color=>c_theme_not_set.
    border_color-indexed   = zcl_excel_style_color=>c_indexed_not_set.
  ENDMETHOD.
ENDCLASS.
