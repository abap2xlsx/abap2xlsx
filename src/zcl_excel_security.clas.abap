CLASS zcl_excel_security DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_SECURITY
*"* do not include other source files here!!!
  PUBLIC SECTION.

    DATA lockrevision TYPE flag .
    DATA lockstructure TYPE flag .
    DATA lockwindows TYPE flag .
    DATA revisionspassword TYPE zexcel_revisionspassword .
    DATA workbookpassword TYPE zexcel_workbookpassword .

    METHODS is_security_enabled
      RETURNING
        VALUE(ep_security_enabled) TYPE flag .
*"* protected components of class ZABAP_EXCEL_SECURITY
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_SECURITY
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZABAP_EXCEL_SECURITY
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_security IMPLEMENTATION.


  METHOD is_security_enabled.
    IF lockrevision EQ abap_true OR lockstructure EQ abap_true OR lockwindows EQ abap_true.
      ep_security_enabled = abap_true.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
