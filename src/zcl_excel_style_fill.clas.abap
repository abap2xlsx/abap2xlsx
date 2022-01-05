CLASS zcl_excel_style_fill DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_STYLE_FILL
*"* do not include other source files here!!!
    CONSTANTS c_fill_none TYPE zexcel_fill_type VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_fill_solid TYPE zexcel_fill_type VALUE 'solid'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_linear TYPE zexcel_fill_type VALUE 'linear'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_path TYPE zexcel_fill_type VALUE 'path'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkdown TYPE zexcel_fill_type VALUE 'darkDown'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkgray TYPE zexcel_fill_type VALUE 'darkGray'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkgrid TYPE zexcel_fill_type VALUE 'darkGrid'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkhorizontal TYPE zexcel_fill_type VALUE 'darkHorizontal'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darktrellis TYPE zexcel_fill_type VALUE 'darkTrellis'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkup TYPE zexcel_fill_type VALUE 'darkUp'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkvertical TYPE zexcel_fill_type VALUE 'darkVertical'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_gray0625 TYPE zexcel_fill_type VALUE 'gray0625'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_gray125 TYPE zexcel_fill_type VALUE 'gray125'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightdown TYPE zexcel_fill_type VALUE 'lightDown'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightgray TYPE zexcel_fill_type VALUE 'lightGray'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightgrid TYPE zexcel_fill_type VALUE 'lightGrid'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lighthorizontal TYPE zexcel_fill_type VALUE 'lightHorizontal'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lighttrellis TYPE zexcel_fill_type VALUE 'lightTrellis'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightup TYPE zexcel_fill_type VALUE 'lightUp'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightvertical TYPE zexcel_fill_type VALUE 'lightVertical'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_mediumgray TYPE zexcel_fill_type VALUE 'mediumGray'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontal90 TYPE zexcel_fill_type VALUE 'horizontal90'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontal270 TYPE zexcel_fill_type VALUE 'horizontal270'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontalb TYPE zexcel_fill_type VALUE 'horizontalb'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_vertical TYPE zexcel_fill_type VALUE 'vertical'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_fromcenter TYPE zexcel_fill_type VALUE 'fromCenter'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal45 TYPE zexcel_fill_type VALUE 'diagonal45'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal45b TYPE zexcel_fill_type VALUE 'diagonal45b'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal135 TYPE zexcel_fill_type VALUE 'diagonal135'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal135b TYPE zexcel_fill_type VALUE 'diagonal135b'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerlt TYPE zexcel_fill_type VALUE 'cornerLT'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerlb TYPE zexcel_fill_type VALUE 'cornerLB'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerrt TYPE zexcel_fill_type VALUE 'cornerRT'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerrb TYPE zexcel_fill_type VALUE 'cornerRB'. "#EC NOTEXT
    DATA gradtype TYPE zexcel_s_gradient_type .
    DATA filltype TYPE zexcel_fill_type .
    DATA rotation TYPE zexcel_rotation .
    DATA fgcolor TYPE zexcel_s_style_color .
    DATA bgcolor TYPE zexcel_s_style_color .

    METHODS constructor .
    METHODS get_structure
      RETURNING
        VALUE(es_fill) TYPE zexcel_s_style_fill .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_FILL
*"* do not include other source files here!!!
  PRIVATE SECTION.

    METHODS build_gradient .
    METHODS check_filltype_is_gradient
      RETURNING
        VALUE(rv_is_gradient) TYPE abap_bool .
ENDCLASS.



CLASS zcl_excel_style_fill IMPLEMENTATION.


  METHOD build_gradient.
    CHECK check_filltype_is_gradient( ) EQ abap_true.
    CLEAR gradtype.
    CASE filltype.
      WHEN c_fill_gradient_horizontal90.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_horizontal270.
        gradtype-degree = '270'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_horizontalb.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_vertical.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_fromcenter.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '0.5'.
        gradtype-top = '0.5'.
        gradtype-left = '0.5'.
        gradtype-right = '0.5'.
      WHEN c_fill_gradient_diagonal45.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_diagonal45b.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_diagonal135.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_diagonal135b.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_cornerlt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_cornerlb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '1'.
        gradtype-top = '1'.
      WHEN c_fill_gradient_cornerrt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-left = '1'.
        gradtype-right = '1'.
      WHEN c_fill_gradient_cornerrb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '0.5'.
        gradtype-top = '0.5'.
        gradtype-left = '0.5'.
        gradtype-right = '0.5'.
    ENDCASE.

  ENDMETHOD.                    "build_gradient


  METHOD check_filltype_is_gradient.
    CASE filltype.
      WHEN c_fill_gradient_horizontal90 OR
           c_fill_gradient_horizontal270 OR
           c_fill_gradient_horizontalb OR
           c_fill_gradient_vertical OR
           c_fill_gradient_fromcenter OR
           c_fill_gradient_diagonal45 OR
           c_fill_gradient_diagonal45b OR
           c_fill_gradient_diagonal135 OR
           c_fill_gradient_diagonal135b OR
           c_fill_gradient_cornerlt OR
           c_fill_gradient_cornerlb OR
           c_fill_gradient_cornerrt OR
           c_fill_gradient_cornerrb.
        rv_is_gradient = abap_true.
    ENDCASE.
  ENDMETHOD.                    "check_filltype_is_gradient


  METHOD constructor.
    filltype = zcl_excel_style_fill=>c_fill_none.
    fgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    fgcolor-indexed   = zcl_excel_style_color=>c_indexed_not_set.
    bgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    bgcolor-indexed   = zcl_excel_style_color=>c_indexed_sys_foreground.
    rotation = 0.

  ENDMETHOD.                    "CONSTRUCTOR


  METHOD get_structure.
    es_fill-rotation  = me->rotation.
    es_fill-filltype  = me->filltype.
    es_fill-fgcolor   = me->fgcolor.
    es_fill-bgcolor   = me->bgcolor.
    me->build_gradient( ).
    es_fill-gradtype = me->gradtype.
  ENDMETHOD.                    "GET_STRUCTURE
ENDCLASS.
