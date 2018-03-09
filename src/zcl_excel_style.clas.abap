class ZCL_EXCEL_STYLE definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLE
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE
*"* do not include other source files here!!!
public section.

  data FONT type ref to ZCL_EXCEL_STYLE_FONT .
  data FILL type ref to ZCL_EXCEL_STYLE_FILL .
  data BORDERS type ref to ZCL_EXCEL_STYLE_BORDERS .
  data ALIGNMENT type ref to ZCL_EXCEL_STYLE_ALIGNMENT .
  data NUMBER_FORMAT type ref to ZCL_EXCEL_STYLE_NUMBER_FORMAT .
  data PROTECTION type ref to ZCL_EXCEL_STYLE_PROTECTION .

  methods CONSTRUCTOR
    importing
      !IP_GUID type ZEXCEL_CELL_STYLE optional
      !IO_CLONE_OF type ref to ZCL_EXCEL_STYLE optional .
  methods GET_GUID
    returning
      value(EP_GUID) type ZEXCEL_CELL_STYLE .
protected section.
*"* private components of class ZCL_EXCEL_STYLE
*"* do not include other source files here!!!
private section.

  data GUID type ZEXCEL_CELL_STYLE .
ENDCLASS.



CLASS ZCL_EXCEL_STYLE IMPLEMENTATION.


METHOD constructor.

  CREATE OBJECT font.
  CREATE OBJECT fill.
  CREATE OBJECT borders.
  CREATE OBJECT alignment.
  CREATE OBJECT number_format.
  CREATE OBJECT protection.

  IF ip_guid IS NOT INITIAL.
    me->guid = ip_guid.
  ELSE.
    me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).
  ENDIF.

  IF io_clone_of IS NOT INITIAL.

    font->bold                 = io_clone_of->font->bold.
    font->color                = io_clone_of->font->color.
    font->family               = io_clone_of->font->family.
    font->italic               = io_clone_of->font->italic.
    font->name                 = io_clone_of->font->name.
    font->scheme               = io_clone_of->font->scheme.
    font->size                 = io_clone_of->font->size.
    font->strikethrough        = io_clone_of->font->strikethrough.
    font->underline            = io_clone_of->font->underline.
    font->underline_mode       = io_clone_of->font->underline_mode.

    fill->gradtype             = io_clone_of->fill->gradtype.
    fill->filltype             = io_clone_of->fill->filltype.
    fill->rotation             = io_clone_of->fill->rotation.
    fill->fgcolor              = io_clone_of->fill->fgcolor.
    fill->bgcolor              = io_clone_of->fill->bgcolor.

    borders->allborders        = io_clone_of->borders->allborders.
    borders->diagonal          = io_clone_of->borders->diagonal.
    borders->diagonal_mode     = io_clone_of->borders->diagonal_mode.
    borders->down              = io_clone_of->borders->down.
    borders->left              = io_clone_of->borders->left.
    borders->right             = io_clone_of->borders->right.
    borders->top               = io_clone_of->borders->top.

    alignment->horizontal      = io_clone_of->alignment->horizontal.
    alignment->vertical        = io_clone_of->alignment->vertical.
    alignment->textrotation    = io_clone_of->alignment->textrotation.
    alignment->wraptext        = io_clone_of->alignment->wraptext.
    alignment->shrinktofit     = io_clone_of->alignment->shrinktofit.
    alignment->indent          = io_clone_of->alignment->indent.

    number_format->format_code = io_clone_of->number_format->format_code.

    protection->hidden         = io_clone_of->protection->hidden.
    protection->locked         = io_clone_of->protection->locked.

  ENDIF.

ENDMETHOD.


method GET_GUID.


  ep_guid = me->guid.
  endmethod.
ENDCLASS.
