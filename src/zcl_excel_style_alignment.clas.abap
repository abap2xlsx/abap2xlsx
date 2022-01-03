CLASS zcl_excel_style_alignment DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLE_ALIGNMENT
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS c_horizontal_general TYPE zexcel_alignment VALUE 'general'. "#EC NOTEXT
    CONSTANTS c_horizontal_left TYPE zexcel_alignment VALUE 'left'. "#EC NOTEXT
    CONSTANTS c_horizontal_right TYPE zexcel_alignment VALUE 'right'. "#EC NOTEXT
    CONSTANTS c_horizontal_center TYPE zexcel_alignment VALUE 'center'. "#EC NOTEXT
    CONSTANTS c_horizontal_center_continuous TYPE zexcel_alignment VALUE 'centerContinuous'. "#EC NOTEXT
    CONSTANTS c_horizontal_justify TYPE zexcel_alignment VALUE 'justify'. "#EC NOTEXT
    CONSTANTS c_vertical_bottom TYPE zexcel_alignment VALUE 'bottom'. "#EC NOTEXT
    CONSTANTS c_vertical_top TYPE zexcel_alignment VALUE 'top'. "#EC NOTEXT
    CONSTANTS c_vertical_center TYPE zexcel_alignment VALUE 'center'. "#EC NOTEXT
    CONSTANTS c_vertical_justify TYPE zexcel_alignment VALUE 'justify'. "#EC NOTEXT
    DATA horizontal TYPE zexcel_alignment .
    DATA vertical TYPE zexcel_alignment .
    DATA textrotation TYPE zexcel_text_rotation VALUE 0. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA wraptext TYPE flag .
    DATA shrinktofit TYPE flag .
    DATA indent TYPE zexcel_indent VALUE 0. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .

    METHODS constructor .
    METHODS get_structure
      RETURNING
        VALUE(es_alignment) TYPE zexcel_s_style_alignment .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_ALIGNMENT
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_style_alignment IMPLEMENTATION.


  METHOD constructor.
    horizontal  = me->c_horizontal_general.
    vertical    = me->c_vertical_bottom.
    wraptext    = abap_false.
    shrinktofit = abap_false.
  ENDMETHOD.


  METHOD get_structure.

    es_alignment-horizontal   = me->horizontal.
    es_alignment-vertical     = me->vertical.
    es_alignment-textrotation = me->textrotation.
    es_alignment-wraptext     = me->wraptext.
    es_alignment-shrinktofit  = me->shrinktofit.
    es_alignment-indent       = me->indent.

  ENDMETHOD.
ENDCLASS.
