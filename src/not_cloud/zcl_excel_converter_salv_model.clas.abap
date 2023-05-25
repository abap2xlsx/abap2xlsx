CLASS zcl_excel_converter_salv_model DEFINITION
  PUBLIC
  INHERITING FROM cl_salv_model
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    CLASS-METHODS is_get_metadata_callable
      IMPORTING
        io_salv       TYPE REF TO cl_salv_table
      RETURNING
        VALUE(result) TYPE abap_bool.
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_converter_salv_model IMPLEMENTATION.

  METHOD is_get_metadata_callable.
    DATA: lo_object TYPE REF TO object,
          lo_model  TYPE REF TO cl_salv_model.

    " In 7.52 and older versions, we have a short dump with CL_SALV_TABLE->GET_METADATA if the ALV is not displayed
    "   (due to io_salv->r_controller->r_adapter not instantiated yet). That's later fixed by SAP (no short dump in 7.57).
    "   NB: r_controller is always instantiated.
    lo_object = io_salv.
    TRY.
        lo_model ?= lo_object.
      CATCH cx_sy_move_cast_error.
        " In 7.57, CL_SALV_TABLE is no more a subclass of CL_SALV_MODEL, but CL_SALV_TABLE->GET_METADATA can be called.
        result = abap_true.
        RETURN.
    ENDTRY.
    result = boolc( lo_model->r_controller->r_adapter IS BOUND ).
  ENDMETHOD.

ENDCLASS.
