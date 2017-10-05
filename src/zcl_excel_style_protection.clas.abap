class ZCL_EXCEL_STYLE_PROTECTION definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_STYLE_PROTECTION
*"* do not include other source files here!!!
public section.

  constants C_PROTECTION_HIDDEN type ZEXCEL_CELL_PROTECTION value '1'. "#EC NOTEXT
  constants C_PROTECTION_LOCKED type ZEXCEL_CELL_PROTECTION value '1'. "#EC NOTEXT
  constants C_PROTECTION_UNHIDDEN type ZEXCEL_CELL_PROTECTION value '0'. "#EC NOTEXT
  constants C_PROTECTION_UNLOCKED type ZEXCEL_CELL_PROTECTION value '0'. "#EC NOTEXT
  data HIDDEN type ZEXCEL_CELL_PROTECTION .
  data LOCKED type ZEXCEL_CELL_PROTECTION .

  methods CONSTRUCTOR .
  methods GET_STRUCTURE
    returning
      value(EP_PROTECTION) type ZEXCEL_S_STYLE_PROTECTION .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_STYLE_PROTECTION
*"* do not include other source files here!!!
private section.
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_PROTECTION IMPLEMENTATION.


method CONSTRUCTOR.
  locked = me->c_protection_locked.
  hidden = me->c_protection_unhidden.
  endmethod.


method GET_STRUCTURE.
  ep_protection-locked = me->locked.
  ep_protection-hidden = me->hidden.
  endmethod.
ENDCLASS.
