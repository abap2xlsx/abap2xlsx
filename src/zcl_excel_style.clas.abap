class ZCL_EXCEL_STYLE definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLE
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
      !IP_GUID type ZEXCEL_CELL_STYLE optional .
  methods GET_GUID
    returning
      value(EP_GUID) type ZEXCEL_CELL_STYLE .
*"* protected components of class ZABAP_EXCEL_STYLE
*"* do not include other source files here!!!
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

ENDMETHOD.


method GET_GUID.


  ep_guid = me->guid.
  endmethod.
ENDCLASS.
