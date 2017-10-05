class ZCL_EXCEL_STYLE_BORDERS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLE_BORDERS
*"* do not include other source files here!!!
public section.

  data ALLBORDERS type ref to ZCL_EXCEL_STYLE_BORDER .
  constants C_DIAGONAL_BOTH type ZEXCEL_DIAGONAL value 3. "#EC NOTEXT
  constants C_DIAGONAL_DOWN type ZEXCEL_DIAGONAL value 2. "#EC NOTEXT
  constants C_DIAGONAL_NONE type ZEXCEL_DIAGONAL value 0. "#EC NOTEXT
  constants C_DIAGONAL_UP type ZEXCEL_DIAGONAL value 1. "#EC NOTEXT
  data DIAGONAL type ref to ZCL_EXCEL_STYLE_BORDER .
  data DIAGONAL_MODE type ZEXCEL_DIAGONAL .
  data DOWN type ref to ZCL_EXCEL_STYLE_BORDER .
  data LEFT type ref to ZCL_EXCEL_STYLE_BORDER .
  data RIGHT type ref to ZCL_EXCEL_STYLE_BORDER .
  data TOP type ref to ZCL_EXCEL_STYLE_BORDER .

  methods GET_STRUCTURE
    returning
      value(ES_FILL) type ZEXCEL_S_STYLE_BORDER .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_STYLE_BORDERS
*"* do not include other source files here!!!
private section.
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_BORDERS IMPLEMENTATION.


method GET_STRUCTURE.
*initialize colors to 'not set'
  es_fill-left_color-indexed = zcl_excel_style_color=>c_indexed_not_set.
  es_fill-left_color-theme = zcl_excel_style_color=>c_theme_not_set.
  es_fill-right_color-indexed = zcl_excel_style_color=>c_indexed_not_set.
  es_fill-right_color-theme = zcl_excel_style_color=>c_theme_not_set.
  es_fill-top_color-indexed = zcl_excel_style_color=>c_indexed_not_set.
  es_fill-top_color-theme = zcl_excel_style_color=>c_theme_not_set.
  es_fill-bottom_color-indexed = zcl_excel_style_color=>c_indexed_not_set.
  es_fill-bottom_color-theme = zcl_excel_style_color=>c_theme_not_set.
  es_fill-diagonal_color-indexed = zcl_excel_style_color=>c_indexed_not_set.
  es_fill-diagonal_color-theme = zcl_excel_style_color=>c_theme_not_set.

* Check if all borders is set otherwise check single border
  IF me->allborders IS BOUND.
    es_fill-left_color    = me->allborders->border_color.
    es_fill-left_style    = me->allborders->border_style.
    es_fill-right_color   = me->allborders->border_color.
    es_fill-right_style   = me->allborders->border_style.
    es_fill-top_color     = me->allborders->border_color.
    es_fill-top_style     = me->allborders->border_style.
    es_fill-bottom_color  = me->allborders->border_color.
    es_fill-bottom_style  = me->allborders->border_style.
  ELSE.
    IF me->left IS BOUND.
      es_fill-left_color = me->left->border_color.
      es_fill-left_style = me->left->border_style.
    ENDIF.
    IF me->right IS BOUND.
      es_fill-right_color = me->right->border_color.
      es_fill-right_style = me->right->border_style.
    ENDIF.
    IF me->top IS BOUND.
      es_fill-top_color = me->top->border_color.
      es_fill-top_style = me->top->border_style.
    ENDIF.
    IF me->down IS BOUND.
      es_fill-bottom_color = me->down->border_color.
      es_fill-bottom_style = me->down->border_style.
    ENDIF.
  ENDIF.

* Check if diagonal is set
  IF me->diagonal IS BOUND.
    es_fill-diagonal_color = me->diagonal->border_color.
    es_fill-diagonal_style = me->diagonal->border_style.
    CASE me->diagonal_mode.
      WHEN 1.
        es_fill-diagonalup     = 1.
        es_fill-diagonaldown   = 0.
      WHEN 2.
        es_fill-diagonalup     = 0.
        es_fill-diagonaldown   = 1.
      WHEN 3.
        es_fill-diagonalup     = 1.
        es_fill-diagonaldown   = 1.
      WHEN OTHERS.
        es_fill-diagonalup     = 0.
        es_fill-diagonaldown   = 0.
    ENDCASE.
  ENDIF.

  endmethod.
ENDCLASS.
