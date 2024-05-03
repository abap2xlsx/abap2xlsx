CLASS zcl_excel_comment DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    CONSTANTS default_right_column TYPE i VALUE 4.          "#EC NOTEXT
    CONSTANTS default_bottom_row TYPE i VALUE 15.           "#EC NOTEXT

    METHODS constructor .
    METHODS get_name
      RETURNING
        VALUE(r_name) TYPE string .
    METHODS get_index
      RETURNING
        VALUE(rp_index) TYPE string .
    METHODS get_ref
      RETURNING
        VALUE(rp_ref) TYPE string .
    METHODS get_text
      RETURNING
        VALUE(rp_text) TYPE string .
    METHODS set_text
      IMPORTING
        !ip_text TYPE string
        !ip_ref  TYPE string OPTIONAL
        !ip_right_column TYPE i OPTIONAL
        !ip_bottom_row TYPE i OPTIONAL .
    METHODS get_position
      EXPORTING
        !ep_right_column TYPE i
        !ep_bottom_row TYPE i .

  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA index TYPE string .
    DATA ref TYPE string .
    DATA text TYPE string .
    DATA right_column TYPE i .
    DATA bottom_row TYPE i .
ENDCLASS.



CLASS zcl_excel_comment IMPLEMENTATION.


  METHOD constructor.

  ENDMETHOD.


  METHOD get_index.
    rp_index = me->index.
  ENDMETHOD.


  METHOD get_name.

  ENDMETHOD.


  METHOD get_ref.
    rp_ref = me->ref.
  ENDMETHOD.


  METHOD get_text.
    rp_text = me->text.
  ENDMETHOD.


  METHOD set_text.
    me->text = ip_text.

    IF ip_ref IS SUPPLIED.
      me->ref = ip_ref.
    ENDIF.

  IF ip_right_column IS SUPPLIED.
    me->right_column = ip_right_column.
  ENDIF.

  IF ip_bottom_row IS SUPPLIED.
    me->bottom_row = ip_bottom_row.
  ENDIF.

  ENDMETHOD.

METHOD GET_POSITION.

  IF right_column IS NOT INITIAL.
    ep_right_column = right_column.
  ELSE.
    ep_right_column = default_right_column. "Default right_column
  ENDIF.

  IF bottom_row IS NOT INITIAL.
    ep_bottom_row = bottom_row.
  ELSE.
    ep_bottom_row = default_bottom_row. "Default bottom_row
  ENDIF.

ENDMETHOD.

ENDCLASS.
