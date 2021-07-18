CLASS zcl_excel_style_borders DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLE_BORDERS
*"* do not include other source files here!!!
  PUBLIC SECTION.

    DATA allborders TYPE REF TO zcl_excel_style_border .
    CONSTANTS c_diagonal_both TYPE zexcel_diagonal VALUE 3. "#EC NOTEXT
    CONSTANTS c_diagonal_down TYPE zexcel_diagonal VALUE 2. "#EC NOTEXT
    CONSTANTS c_diagonal_none TYPE zexcel_diagonal VALUE 0. "#EC NOTEXT
    CONSTANTS c_diagonal_up TYPE zexcel_diagonal VALUE 1.   "#EC NOTEXT
    DATA diagonal TYPE REF TO zcl_excel_style_border .
    DATA diagonal_mode TYPE zexcel_diagonal .
    DATA down TYPE REF TO zcl_excel_style_border .
    DATA left TYPE REF TO zcl_excel_style_border .
    DATA right TYPE REF TO zcl_excel_style_border .
    DATA top TYPE REF TO zcl_excel_style_border .

    METHODS get_structure
      RETURNING
        VALUE(es_fill) TYPE zexcel_s_style_border .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_BORDERS
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_style_borders IMPLEMENTATION.


  METHOD get_structure.
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

  ENDMETHOD.
ENDCLASS.
