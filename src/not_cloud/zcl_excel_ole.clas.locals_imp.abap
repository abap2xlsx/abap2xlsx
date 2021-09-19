*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations

*&---------------------------------------------------------------------*
*&       Class (Implementation)  C_OI_PROXY_ERROR
*&---------------------------------------------------------------------*
CLASS c_oi_proxy_error IMPLEMENTATION.
  METHOD constructor.
*                IMPORTING object_name TYPE c
*                          method_name TYPE c.
    error_nr = ret_call_not_flushed.
    me->i_oi_error~error_code = c_oi_errors=>ret_call_not_flushed.
    me->i_oi_error~is_flushed = ' '.
    me->i_oi_error~has_failed = 'X'.
    me->i_oi_error~has_succeeded = ' '.
    me->message_id = 'SOFFICEINTEGRATION'.
    me->message_nr = '899'.
    me->param1 = object_name.
    me->param2 = method_name.
  ENDMETHOD.                    "constructor

  METHOD i_oi_error~flush_error.
    IF error_nr EQ 0.
      me->i_oi_error~error_code = c_oi_errors=>ret_ok.
      me->i_oi_error~is_flushed = 'X'.
      me->i_oi_error~has_failed = ' '.
      me->i_oi_error~has_succeeded = 'X'.
      me->message_id = ''.
      me->message_nr = '000'.
      CALL METHOD c_oi_errors=>translate_proxy_error_code
        EXPORTING
          errorcode = error_nr
        IMPORTING
          retcode   = me->i_oi_error~error_code.
    ELSEIF error_nr EQ ret_call_not_flushed.
      "call still not flushed
      CALL METHOD c_oi_errors=>translate_proxy_error_code
        EXPORTING
          errorcode   = error_nr
          errorstring = me->param2  "method name
          objectname  = me->param1
        IMPORTING
          retcode     = me->i_oi_error~error_code.
    ELSE.
      me->i_oi_error~is_flushed = 'X'.
      me->i_oi_error~has_succeeded = ' '.
      me->i_oi_error~has_failed = 'X'.
      CALL METHOD c_oi_errors=>translate_proxy_error_code
        EXPORTING
          errorcode   = error_nr
          errorstring = error_string
        IMPORTING
          retcode     = me->i_oi_error~error_code.
      CALL METHOD c_oi_errors=>get_message
        IMPORTING
          message_id     = me->message_id
          message_number = me->message_nr
          param1         = me->param1
          param2         = me->param2
          param3         = me->param3
          param4         = me->param4.
    ENDIF.
  ENDMETHOD.                    "i_oi_error~flush_error

  METHOD i_oi_error~raise_message.
*                         IMPORTING type TYPE c.
*                         EXCEPTIONS message_raised flush_failed.
    IF me->i_oi_error~has_succeeded IS INITIAL.
      IF NOT me->i_oi_error~is_flushed IS INITIAL.
        MESSAGE ID message_id TYPE type
            NUMBER message_nr WITH param1 param2 param3 param4
            RAISING message_raised.
      ELSE.
        RAISE flush_failed.
      ENDIF.
    ENDIF.
  ENDMETHOD.                    "i_oi_error~raise_message

  METHOD i_oi_error~get_message.
*                    EXPORTING message_id TYPE c
*                              message_number TYPE c
*                              param1 TYPE c
*                              param2 TYPE c
*                              param3 TYPE c
*                              param4 TYPE c.
    param1 = me->param1. param2 = me->param2.
    param3 = me->param3. param4 = me->param4.

    message_id = me->message_id.
    message_number = me->message_nr.
  ENDMETHOD.                    "i_oi_error~get_message
ENDCLASS.               "C_OI_PROXY_ERROR

*&---------------------------------------------------------------------*
*&       Class (Implementation)  CL_GRID_ACCESSION
*&---------------------------------------------------------------------*
CLASS lcl_gui_alv_grid IMPLEMENTATION.

  METHOD get_alv_attributes.
    CREATE DATA et_table LIKE io_grid->mt_outtab.
    et_table = io_grid->mt_outtab.
  ENDMETHOD.                    "get_data

ENDCLASS.               "CL_GRID_ACCESSION
