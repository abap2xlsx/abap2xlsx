class ZCL_EXCEL_STYLE_FILL definition
  public
  final
  create public .

public section.
  type-pools ABAP .

*"* public components of class ZCL_EXCEL_STYLE_FILL
*"* do not include other source files here!!!
  constants C_FILL_NONE type ZEXCEL_FILL_TYPE value 'none'. "#EC NOTEXT
  constants C_FILL_SOLID type ZEXCEL_FILL_TYPE value 'solid'. "#EC NOTEXT
  constants C_FILL_GRADIENT_LINEAR type ZEXCEL_FILL_TYPE value 'linear'. "#EC NOTEXT
  constants C_FILL_GRADIENT_PATH type ZEXCEL_FILL_TYPE value 'path'. "#EC NOTEXT
  constants C_FILL_PATTERN_DARKDOWN type ZEXCEL_FILL_TYPE value 'darkDown'. "#EC NOTEXT
  constants C_FILL_PATTERN_DARKGRAY type ZEXCEL_FILL_TYPE value 'darkGray'. "#EC NOTEXT
  constants C_FILL_PATTERN_DARKGRID type ZEXCEL_FILL_TYPE value 'darkGrid'. "#EC NOTEXT
  constants C_FILL_PATTERN_DARKHORIZONTAL type ZEXCEL_FILL_TYPE value 'darkHorizontal'. "#EC NOTEXT
  constants C_FILL_PATTERN_DARKTRELLIS type ZEXCEL_FILL_TYPE value 'darkTrellis'. "#EC NOTEXT
  constants C_FILL_PATTERN_DARKUP type ZEXCEL_FILL_TYPE value 'darkUp'. "#EC NOTEXT
  constants C_FILL_PATTERN_DARKVERTICAL type ZEXCEL_FILL_TYPE value 'darkVertical'. "#EC NOTEXT
  constants C_FILL_PATTERN_GRAY0625 type ZEXCEL_FILL_TYPE value 'gray0625'. "#EC NOTEXT
  constants C_FILL_PATTERN_GRAY125 type ZEXCEL_FILL_TYPE value 'gray125'. "#EC NOTEXT
  constants C_FILL_PATTERN_LIGHTDOWN type ZEXCEL_FILL_TYPE value 'lightDown'. "#EC NOTEXT
  constants C_FILL_PATTERN_LIGHTGRAY type ZEXCEL_FILL_TYPE value 'lightGray'. "#EC NOTEXT
  constants C_FILL_PATTERN_LIGHTGRID type ZEXCEL_FILL_TYPE value 'lightGrid'. "#EC NOTEXT
  constants C_FILL_PATTERN_LIGHTHORIZONTAL type ZEXCEL_FILL_TYPE value 'lightHorizontal'. "#EC NOTEXT
  constants C_FILL_PATTERN_LIGHTTRELLIS type ZEXCEL_FILL_TYPE value 'lightTrellis'. "#EC NOTEXT
  constants C_FILL_PATTERN_LIGHTUP type ZEXCEL_FILL_TYPE value 'lightUp'. "#EC NOTEXT
  constants C_FILL_PATTERN_LIGHTVERTICAL type ZEXCEL_FILL_TYPE value 'lightVertical'. "#EC NOTEXT
  constants C_FILL_PATTERN_MEDIUMGRAY type ZEXCEL_FILL_TYPE value 'mediumGray'. "#EC NOTEXT
  constants C_FILL_GRADIENT_HORIZONTAL90 type ZEXCEL_FILL_TYPE value 'horizontal90'. "#EC NOTEXT
  constants C_FILL_GRADIENT_HORIZONTAL270 type ZEXCEL_FILL_TYPE value 'horizontal270'. "#EC NOTEXT
  constants C_FILL_GRADIENT_HORIZONTALB type ZEXCEL_FILL_TYPE value 'horizontalb'. "#EC NOTEXT
  constants C_FILL_GRADIENT_VERTICAL type ZEXCEL_FILL_TYPE value 'vertical'. "#EC NOTEXT
  constants C_FILL_GRADIENT_FROMCENTER type ZEXCEL_FILL_TYPE value 'fromCenter'. "#EC NOTEXT
  constants C_FILL_GRADIENT_DIAGONAL45 type ZEXCEL_FILL_TYPE value 'diagonal45'. "#EC NOTEXT
  constants C_FILL_GRADIENT_DIAGONAL45B type ZEXCEL_FILL_TYPE value 'diagonal45b'. "#EC NOTEXT
  constants C_FILL_GRADIENT_DIAGONAL135 type ZEXCEL_FILL_TYPE value 'diagonal135'. "#EC NOTEXT
  constants C_FILL_GRADIENT_DIAGONAL135B type ZEXCEL_FILL_TYPE value 'diagonal135b'. "#EC NOTEXT
  constants C_FILL_GRADIENT_CORNERLT type ZEXCEL_FILL_TYPE value 'cornerLT'. "#EC NOTEXT
  constants C_FILL_GRADIENT_CORNERLB type ZEXCEL_FILL_TYPE value 'cornerLB'. "#EC NOTEXT
  constants C_FILL_GRADIENT_CORNERRT type ZEXCEL_FILL_TYPE value 'cornerRT'. "#EC NOTEXT
  constants C_FILL_GRADIENT_CORNERRB type ZEXCEL_FILL_TYPE value 'cornerRB'. "#EC NOTEXT
  data GRADTYPE type ZEXCEL_S_GRADIENT_TYPE .
  data FILLTYPE type ZEXCEL_FILL_TYPE .
  data ROTATION type ZEXCEL_ROTATION .
  data FGCOLOR type ZEXCEL_S_STYLE_COLOR .
  data BGCOLOR type ZEXCEL_S_STYLE_COLOR .

  methods CONSTRUCTOR .
  methods GET_STRUCTURE
    returning
      value(ES_FILL) type ZEXCEL_S_STYLE_FILL .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_STYLE_FILL
*"* do not include other source files here!!!
private section.

  methods BUILD_GRADIENT .
  methods CHECK_FILLTYPE_IS_GRADIENT
    returning
      value(RV_IS_GRADIENT) type ABAP_BOOL .
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_FILL IMPLEMENTATION.


method build_gradient.
    check check_filltype_is_gradient( ) eq abap_true.
    clear gradtype.
    case filltype.
      when c_fill_gradient_horizontal90.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      when c_fill_gradient_horizontal270.
        gradtype-degree = '270'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      when c_fill_gradient_horizontalb.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      when c_fill_gradient_vertical.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      when c_fill_gradient_fromcenter.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '0.5'.
        gradtype-top = '0.5'.
        gradtype-left = '0.5'.
        gradtype-right = '0.5'.
      when c_fill_gradient_diagonal45.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      when c_fill_gradient_diagonal45b.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      when c_fill_gradient_diagonal135.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      when c_fill_gradient_diagonal135b.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      when c_fill_gradient_cornerlt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      when c_fill_gradient_cornerlb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '1'.
        gradtype-top = '1'.
      when c_fill_gradient_cornerrt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-left = '1'.
        gradtype-right = '1'.
      when c_fill_gradient_cornerrb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '0.5'.
        gradtype-top = '0.5'.
        gradtype-left = '0.5'.
        gradtype-right = '0.5'.
    endcase.

  endmethod.                    "build_gradient


method check_filltype_is_gradient.
    case filltype.
      when c_fill_gradient_horizontal90 or
           c_fill_gradient_horizontal270 or
           c_fill_gradient_horizontalb or
           c_fill_gradient_vertical or
           c_fill_gradient_fromcenter or
           c_fill_gradient_diagonal45 or
           c_fill_gradient_diagonal45b or
           c_fill_gradient_diagonal135 or
           c_fill_gradient_diagonal135b or
           c_fill_gradient_cornerlt or
           c_fill_gradient_cornerlb or
           c_fill_gradient_cornerrt or
           c_fill_gradient_cornerrb.
        rv_is_gradient = abap_true.
    endcase.
  endmethod.                    "check_filltype_is_gradient


method constructor.
    filltype = zcl_excel_style_fill=>c_fill_none.
    fgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    fgcolor-indexed   = zcl_excel_style_color=>c_indexed_not_set.
    bgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    bgcolor-indexed   = zcl_excel_style_color=>c_indexed_sys_foreground.
    rotation = 0.

  endmethod.                    "CONSTRUCTOR


method get_structure.
    es_fill-rotation  = me->rotation.
    es_fill-filltype  = me->filltype.
    es_fill-fgcolor   = me->fgcolor.
    es_fill-bgcolor   = me->bgcolor.
    me->build_gradient( ).
    es_fill-gradtype = me->gradtype.
  endmethod.                    "GET_STRUCTURE
ENDCLASS.
