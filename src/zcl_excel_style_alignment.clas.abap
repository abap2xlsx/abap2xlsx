class ZCL_EXCEL_STYLE_ALIGNMENT definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLE_ALIGNMENT
*"* do not include other source files here!!!
public section.
  type-pools ABAP .

  constants C_HORIZONTAL_GENERAL type ZEXCEL_ALIGNMENT value 'general'. "#EC NOTEXT
  constants C_HORIZONTAL_LEFT type ZEXCEL_ALIGNMENT value 'left'. "#EC NOTEXT
  constants C_HORIZONTAL_RIGHT type ZEXCEL_ALIGNMENT value 'right'. "#EC NOTEXT
  constants C_HORIZONTAL_CENTER type ZEXCEL_ALIGNMENT value 'center'. "#EC NOTEXT
  constants C_HORIZONTAL_CENTER_CONTINUOUS type ZEXCEL_ALIGNMENT value 'centerContinuous'. "#EC NOTEXT
  constants C_HORIZONTAL_JUSTIFY type ZEXCEL_ALIGNMENT value 'justify'. "#EC NOTEXT
  constants C_VERTICAL_BOTTOM type ZEXCEL_ALIGNMENT value 'bottom'. "#EC NOTEXT
  constants C_VERTICAL_TOP type ZEXCEL_ALIGNMENT value 'top'. "#EC NOTEXT
  constants C_VERTICAL_CENTER type ZEXCEL_ALIGNMENT value 'center'. "#EC NOTEXT
  constants C_VERTICAL_JUSTIFY type ZEXCEL_ALIGNMENT value 'justify'. "#EC NOTEXT
  data HORIZONTAL type ZEXCEL_ALIGNMENT .
  data VERTICAL type ZEXCEL_ALIGNMENT .
  data TEXTROTATION type ZEXCEL_TEXT_ROTATION value 0. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data WRAPTEXT type FLAG .
  data SHRINKTOFIT type FLAG .
  data INDENT type ZEXCEL_INDENT value 0. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .

  methods CONSTRUCTOR .
  methods GET_STRUCTURE
    returning
      value(ES_ALIGNMENT) type ZEXCEL_S_STYLE_ALIGNMENT .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_STYLE_ALIGNMENT
*"* do not include other source files here!!!
private section.
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_ALIGNMENT IMPLEMENTATION.


method CONSTRUCTOR.
  horizontal  = me->c_horizontal_general.
  vertical    = me->c_vertical_bottom.
  wrapText    = abap_false.
  shrinkToFit = abap_false.
  endmethod.


method GET_STRUCTURE.

  es_alignment-horizontal   = me->horizontal.
  es_alignment-vertical     = me->vertical.
  es_alignment-textrotation = me->textrotation.
  es_alignment-wraptext     = me->wraptext.
  es_alignment-shrinktofit  = me->shrinktofit.
  es_alignment-indent       = me->indent.

  endmethod.
ENDCLASS.
