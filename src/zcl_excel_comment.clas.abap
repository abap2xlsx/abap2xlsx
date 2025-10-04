CLASS zcl_excel_comment DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF ty_box,
        left_column   TYPE i,
        left_offset   TYPE i,
        top_row       TYPE i,
        top_offset    TYPE i,
        right_column  TYPE i,
        right_offset  TYPE i,
        bottom_row    TYPE i,
        bottom_offset TYPE i,
      END OF ty_box .

    CONSTANTS:
      BEGIN OF gc_default_box,
        left_column   TYPE i VALUE 2,
        left_offset   TYPE i VALUE 15,
        top_row       TYPE i VALUE 11,
        top_offset    TYPE i VALUE 10,
        right_column  TYPE i VALUE 4,
        right_offset  TYPE i VALUE 31,
        bottom_row    TYPE i VALUE 15,
        bottom_offset TYPE i VALUE 9,
      END OF gc_default_box .

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
    METHODS set_box
      IMPORTING
        !is_box TYPE ty_box .
    METHODS set_text
      IMPORTING
        !ip_text          TYPE string
        !ip_ref           TYPE string OPTIONAL
        !ip_left_column   TYPE i DEFAULT gc_default_box-left_column
        !ip_left_offset   TYPE i DEFAULT gc_default_box-left_offset
        !ip_top_row       TYPE i DEFAULT gc_default_box-top_row
        !ip_top_offset    TYPE i DEFAULT gc_default_box-top_offset
        !ip_right_column  TYPE i DEFAULT gc_default_box-right_column
        !ip_right_offset  TYPE i DEFAULT gc_default_box-right_offset
        !ip_bottom_row    TYPE i DEFAULT gc_default_box-bottom_row
        !ip_bottom_offset TYPE i DEFAULT gc_default_box-bottom_offset.

  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA index TYPE string .
    DATA ref TYPE string .
    DATA text TYPE string .
    DATA gs_box TYPE ty_box .

ENDCLASS.



CLASS zcl_excel_comment IMPLEMENTATION.


  METHOD constructor.

  ENDMETHOD.


  METHOD get_bottom_offset.
    rp_result = gs_box-bottom_offset.
  ENDMETHOD.


  METHOD get_bottom_row.
    rp_result = gs_box-bottom_row.
  ENDMETHOD.


  METHOD get_index.
    rp_index = me->index.
  ENDMETHOD.


  METHOD get_left_column.
    rp_result = gs_box-left_column.
  ENDMETHOD.


  METHOD get_left_offset.
    rp_result = gs_box-left_offset.
  ENDMETHOD.


  METHOD get_name.

  ENDMETHOD.


  METHOD get_ref.
    rp_ref = me->ref.
  ENDMETHOD.


  METHOD get_right_column.
    rp_result = gs_box-right_column.
  ENDMETHOD.


  METHOD get_right_offset.
    rp_result = gs_box-right_offset.
  ENDMETHOD.


  METHOD get_text.
    rp_text = me->text.
  ENDMETHOD.


  METHOD get_top_offset.
    rp_result = gs_box-top_offset.
  ENDMETHOD.


  METHOD get_top_row.
    rp_result = gs_box-top_row.
  ENDMETHOD.


  METHOD set_text.
    me->text = ip_text.

    IF ip_ref IS SUPPLIED.
      me->ref = ip_ref.
    ENDIF.

* Parameters of the containing box
    DATA ls_box TYPE ty_box.
    ls_box-left_column   = ip_left_column.
    ls_box-left_offset   = ip_left_offset.
    ls_box-top_row       = ip_top_row.
    ls_box-top_offset    = ip_top_offset.
    IF ip_right_column IS NOT INITIAL.
      ls_box-right_column = ip_right_column.
    ELSE.
      ls_box-right_column = gc_default_box-right_column.
    ENDIF.
    ls_box-right_offset = ip_right_offset.
    IF ip_bottom_row IS NOT INITIAL.
      ls_box-bottom_row = ip_bottom_row.
    ELSE.
      ls_box-bottom_row = gc_default_box-bottom_row.
    ENDIF.
    ls_box-bottom_offset = ip_bottom_offset.
    set_box( ls_box ).

  ENDMETHOD.

  METHOD set_box.

    gs_box = is_box.

  ENDMETHOD.

ENDCLASS.
