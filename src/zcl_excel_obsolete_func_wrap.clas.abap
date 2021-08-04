CLASS zcl_excel_obsolete_func_wrap DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    CLASS-METHODS guid_create
      RETURNING
        VALUE(rv_guid_16) TYPE guid_16 .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_obsolete_func_wrap IMPLEMENTATION.


  METHOD guid_create.

    TRY.
        rv_guid_16 = cl_system_uuid=>if_system_uuid_static~create_uuid_x16( ).
      CATCH cx_uuid_error.
    ENDTRY.

*--------------------------------------------------------------------*
* If you are on a release that does not yet have the class cl_system_uuid
* please use the following coding instead which is using the function
* call that was used before but which has been flagged as obsolete
* in newer SAP releases
*--------------------------------------------------------------------*
*
*  CALL FUNCTION 'GUID_CREATE'
*    IMPORTING
*      ev_guid_16 = rv_guid_16.

  ENDMETHOD.
ENDCLASS.
