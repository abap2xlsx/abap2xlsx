CLASS zcl_excel_comment DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    CONSTANTS default_right_column TYPE i VALUE 4.  "#EC NOTEXT
    CONSTANTS default_bottom_row   TYPE i VALUE 15. "#EC NOTEXT

    METHODS constructor .
    METHODS get_bottom_offset
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS get_bottom_row
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS get_index
      RETURNING
        VALUE(rp_index) TYPE string .
    METHODS get_left_column
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS get_left_offset
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS get_name
      RETURNING
        VALUE(r_name) TYPE string .
    METHODS get_ref
      RETURNING
        VALUE(rp_ref) TYPE string .
    METHODS get_right_column
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS get_right_offset
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS get_text
      RETURNING
        VALUE(rp_text) TYPE string .
    METHODS get_top_offset
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS get_top_row
      RETURNING
        VALUE(rp_result) TYPE i .
    METHODS set_text
      IMPORTING
        !ip_text          TYPE string
        !ip_ref           TYPE string OPTIONAL
        !ip_left_column   TYPE i DEFAULT 2
        !ip_left_offset   TYPE i DEFAULT 15
        !ip_top_row       TYPE i DEFAULT 11
        !ip_top_offset    TYPE i DEFAULT 10
        !ip_right_column  TYPE i DEFAULT default_right_column
        !ip_right_offset  TYPE i DEFAULT 31
        !ip_bottom_row    TYPE i DEFAULT default_bottom_row
        !ip_bottom_offset TYPE i DEFAULT 9.

  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA bottom_offset TYPE i .
    DATA bottom_row TYPE i .
    DATA index TYPE string .
    DATA ref TYPE string .
    DATA left_column TYPE i .
    DATA left_offset TYPE i .
    DATA right_column TYPE i .
    DATA right_offset TYPE i .
    DATA text TYPE string .
    DATA top_offset TYPE i .
    DATA top_row TYPE i .
ENDCLASS.



CLASS zcl_excel_comment IMPLEMENTATION.


  METHOD constructor.

  ENDMETHOD.


  METHOD get_bottom_offset.
    rp_result = bottom_offset.
  ENDMETHOD.


  METHOD get_bottom_row.
    rp_result = bottom_row.
  ENDMETHOD.


  METHOD get_index.
    rp_index = me->index.
  ENDMETHOD.


  METHOD get_left_column.
    rp_result = left_column.
  ENDMETHOD.


  METHOD get_left_offset.
    rp_result = left_offset.
  ENDMETHOD.


  METHOD get_name.

  ENDMETHOD.


  METHOD get_ref.
    rp_ref = me->ref.
  ENDMETHOD.


  METHOD get_right_column.
    rp_result = right_column.
  ENDMETHOD.


  METHOD get_right_offset.
    rp_result = right_offset.
  ENDMETHOD.


  METHOD get_text.
    rp_text = me->text.
  ENDMETHOD.


  METHOD get_top_offset.
    rp_result = top_offset.
  ENDMETHOD.


  METHOD get_top_row.
    rp_result = top_row.
  ENDMETHOD.


  METHOD set_text.
    me->text = ip_text.

    IF ip_ref IS SUPPLIED.
      me->ref = ip_ref.
    ENDIF.

    me->left_column = ip_left_column.
    me->left_offset = ip_left_offset.

    me->top_row    = ip_top_row.
    me->top_offset = ip_top_offset.

    IF ip_right_column IS NOT INITIAL.
      me->right_column = ip_right_column.
    ELSE.
      me->right_column = default_right_column.
    ENDIF.
    me->right_offset = ip_right_offset.

    IF ip_bottom_row IS NOT INITIAL.
      me->bottom_row = ip_bottom_row.
    ELSE.
      me->bottom_row = default_bottom_row.
    ENDIF.
    me->bottom_offset = ip_bottom_offset.
  ENDMETHOD.

ENDCLASS.
