CLASS zcl_excel_legacy_palette DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_LEGACY_PALETTE
*"* do not include other source files here!!!

    METHODS constructor .
    METHODS is_modified
      RETURNING
        VALUE(ep_modified) TYPE abap_bool .
    METHODS get_color
      IMPORTING
        !ip_index       TYPE i
      RETURNING
        VALUE(ep_color) TYPE zexcel_style_color_argb
      RAISING
        zcx_excel .
    METHODS get_colors
      RETURNING
        VALUE(ep_colors) TYPE zexcel_t_style_color_argb .
    METHODS set_color
      IMPORTING
        !ip_index TYPE i
        !ip_color TYPE zexcel_style_color_argb
      RAISING
        zcx_excel .
  PROTECTED SECTION.
*"* protected components of class ZCL_EXCEL_LEGACY_PALETTE
*"* do not include other source files here!!!
  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_LEGACY_PALETTE
*"* do not include other source files here!!!
    DATA modified TYPE abap_bool VALUE abap_false. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA colors TYPE zexcel_t_style_color_argb .
ENDCLASS.



CLASS zcl_excel_legacy_palette IMPLEMENTATION.


  METHOD constructor.
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

  ENDMETHOD.


  METHOD get_color.
    DATA: lv_index TYPE i.

    lv_index = ip_index + 1.
    READ TABLE colors INTO ep_color INDEX lv_index.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( 'Invalid color index' ).
    ENDIF.
  ENDMETHOD.


  METHOD get_colors.
    ep_colors = colors.
  ENDMETHOD.


  METHOD is_modified.
    ep_modified = modified.
  ENDMETHOD.


  METHOD set_color.
    DATA: lv_index TYPE i.

    FIELD-SYMBOLS: <lv_color> LIKE LINE OF colors.

    lv_index = ip_index + 1.
    READ TABLE colors ASSIGNING <lv_color> INDEX lv_index.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( 'Invalid color index' ).
    ENDIF.

    IF <lv_color> <> ip_color.
      modified = abap_true.
      <lv_color> = ip_color.
    ENDIF.

  ENDMETHOD.
ENDCLASS.
