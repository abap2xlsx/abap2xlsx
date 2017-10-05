class ZCL_EXCEL_STYLE_COLOR definition
  public
  final
  create public .

public section.

*"* public components of class ZCL_EXCEL_STYLE_COLOR
*"* do not include other source files here!!!
  constants C_BLACK type ZEXCEL_STYLE_COLOR_ARGB value 'FF000000'. "#EC NOTEXT
  constants C_BLUE type ZEXCEL_STYLE_COLOR_ARGB value 'FF0000FF'. "#EC NOTEXT
  constants C_DARKBLUE type ZEXCEL_STYLE_COLOR_ARGB value 'FF000080'. "#EC NOTEXT
  constants C_DARKGREEN type ZEXCEL_STYLE_COLOR_ARGB value 'FF008000'. "#EC NOTEXT
  constants C_DARKRED type ZEXCEL_STYLE_COLOR_ARGB value 'FF800000'. "#EC NOTEXT
  constants C_DARKYELLOW type ZEXCEL_STYLE_COLOR_ARGB value 'FF808000'. "#EC NOTEXT
  constants C_GRAY type ZEXCEL_STYLE_COLOR_ARGB value 'FFCCCCCC'. "#EC NOTEXT
  constants C_GREEN type ZEXCEL_STYLE_COLOR_ARGB value 'FF00FF00'. "#EC NOTEXT
  constants C_RED type ZEXCEL_STYLE_COLOR_ARGB value 'FFFF0000'. "#EC NOTEXT
  constants C_WHITE type ZEXCEL_STYLE_COLOR_ARGB value 'FFFFFFFF'. "#EC NOTEXT
  constants C_YELLOW type ZEXCEL_STYLE_COLOR_ARGB value 'FFFFFF00'. "#EC NOTEXT
  constants C_THEME_DARK1 type ZEXCEL_STYLE_COLOR_THEME value 0. "#EC NOTEXT
  constants C_THEME_LIGHT1 type ZEXCEL_STYLE_COLOR_THEME value 1. "#EC NOTEXT
  constants C_THEME_DARK2 type ZEXCEL_STYLE_COLOR_THEME value 2. "#EC NOTEXT
  constants C_THEME_LIGHT2 type ZEXCEL_STYLE_COLOR_THEME value 3. "#EC NOTEXT
  constants C_THEME_ACCENT1 type ZEXCEL_STYLE_COLOR_THEME value 4. "#EC NOTEXT
  constants C_THEME_ACCENT2 type ZEXCEL_STYLE_COLOR_THEME value 5. "#EC NOTEXT
  constants C_THEME_ACCENT3 type ZEXCEL_STYLE_COLOR_THEME value 6. "#EC NOTEXT
  constants C_THEME_ACCENT4 type ZEXCEL_STYLE_COLOR_THEME value 7. "#EC NOTEXT
  constants C_THEME_ACCENT5 type ZEXCEL_STYLE_COLOR_THEME value 8. "#EC NOTEXT
  constants C_THEME_ACCENT6 type ZEXCEL_STYLE_COLOR_THEME value 9. "#EC NOTEXT
  constants C_THEME_HYPERLINK type ZEXCEL_STYLE_COLOR_THEME value 10. "#EC NOTEXT
  constants C_THEME_HYPERLINK_FOLLOWED type ZEXCEL_STYLE_COLOR_THEME value 11. "#EC NOTEXT
  constants C_THEME_NOT_SET type ZEXCEL_STYLE_COLOR_THEME value -1. "#EC NOTEXT
  constants C_INDEXED_NOT_SET type ZEXCEL_STYLE_COLOR_INDEXED value -1. "#EC NOTEXT
  constants C_INDEXED_SYS_FOREGROUND type ZEXCEL_STYLE_COLOR_INDEXED value 64. "#EC NOTEXT

  class-methods CREATE_NEW_ARGB
    importing
      !IP_RED type ZEXCEL_STYLE_COLOR_COMPONENT
      !IP_GREEN type ZEXCEL_STYLE_COLOR_COMPONENT
      !IP_BLU type ZEXCEL_STYLE_COLOR_COMPONENT
    returning
      value(EP_COLOR_ARGB) type ZEXCEL_STYLE_COLOR_ARGB .
  class-methods CREATE_NEW_ARBG_INT
    importing
      !IV_RED type NUMERIC
      !IV_GREEN type NUMERIC
      !IV_BLUE type NUMERIC
    returning
      value(RV_COLOR_ARGB) type ZEXCEL_STYLE_COLOR_ARGB .
*"* protected components of class ZCL_EXCEL_STYLE_COLOR
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_STYLE_COLOR
*"* do not include other source files here!!!
protected section.
private section.

*"* private components of class ZCL_EXCEL_STYLE_COLOR
*"* do not include other source files here!!!
  constants C_ALPHA type CHAR2 value 'FF'. "#EC NOTEXT
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_COLOR IMPLEMENTATION.


METHOD create_new_arbg_int.
  DATA: lv_red        TYPE int1,
        lv_green      TYPE int1,
        lv_blue       TYPE int1,
        lv_hex        TYPE x,
        lv_char_red   TYPE zexcel_style_color_component,
        lv_char_green TYPE zexcel_style_color_component,
        lv_char_blue  TYPE zexcel_style_color_component.

  lv_red    = iv_red MOD 256.
  lv_green  = iv_green MOD 256.
  lv_blue   = iv_blue  MOD 256.

  lv_hex        = lv_red.
  lv_char_red   = lv_hex.

  lv_hex        = lv_green.
  lv_char_green = lv_hex.

  lv_hex        = lv_blue.
  lv_char_blue  = lv_hex.


  concatenate zcl_excel_style_color=>c_alpha lv_char_red lv_char_green lv_char_blue into rv_color_argb.


ENDMETHOD.


METHOD create_new_argb.

  CONCATENATE zcl_excel_style_color=>c_alpha ip_red ip_green ip_blu INTO ep_color_argb.

ENDMETHOD.
ENDCLASS.
