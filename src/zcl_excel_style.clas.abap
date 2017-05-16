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

* Start of insertion # issue 139 - Dateretention of cellstyles
  IF ip_guid IS NOT INITIAL.
    me->guid = ip_guid.
  ELSE.
* End of insertion # issue 139 - Dateretention of cellstyles
*    CALL FUNCTION 'GUID_CREATE'                                  " del issue #379 - function is outdated in newer releases
*      IMPORTING
*        ev_guid_16 = me->guid.
    me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).      " ins issue #379 - replacement for outdated function call
* Start of insertion # issue 139 - Dateretention of cellstyles
  ENDIF.
* End of insertion # issue 139 - Dateretention of cellstyles

ENDMETHOD.


method GET_GUID.


  ep_guid = me->guid.
  endmethod.
ENDCLASS.
