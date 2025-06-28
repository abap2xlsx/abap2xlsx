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
        !ip_text          TYPE string OPTIONAL
        !ip_ref           TYPE string OPTIONAL
        !ip_topleft       TYPE string OPTIONAL
        !ip_width         TYPE i OPTIONAL
        !ip_height        TYPE i OPTIONAL
        !ip_left_column   TYPE i DEFAULT gc_default_box-left_column
        !ip_left_offset   TYPE i OPTIONAL
        !ip_top_row       TYPE i DEFAULT gc_default_box-top_row
        !ip_top_offset    TYPE i OPTIONAL
        !ip_right_column  TYPE i DEFAULT gc_default_box-right_column
        !ip_right_offset  TYPE i OPTIONAL
        !ip_bottom_row    TYPE i DEFAULT gc_default_box-bottom_row
        !ip_bottom_offset TYPE i OPTIONAL
      RAISING
        zcx_excel .

  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA text TYPE string .
    DATA ref TYPE string .
    DATA index TYPE string .
    DATA topleft_column TYPE zexcel_cell_column .
    DATA topleft_row TYPE zexcel_cell_row .
    DATA width TYPE i .
    DATA height TYPE i .
    DATA left_offset TYPE i .
    DATA top_offset TYPE i .
    DATA right_offset TYPE i .
    DATA bottom_offset TYPE i .

ENDCLASS.



CLASS zcl_excel_comment IMPLEMENTATION.


  METHOD constructor.
    " If visibility in excel is set to on:
    " The comment box will be placed at the specified coordinates starting at 0.
    " If visibility is set to hidden (this is the ABAP2XLSX standard):
    " The comment's position is adjusted right above to its referenced cell position.

    me->width  = gc_default_box-right_column - gc_default_box-left_column.  "2
    me->height = gc_default_box-bottom_row   - gc_default_box-top_row.      "4

    me->left_offset   = gc_default_box-left_offset.
    me->top_offset    = gc_default_box-top_offset.
    me->right_offset  = gc_default_box-right_offset.
    me->bottom_offset = gc_default_box-bottom_offset.

  ENDMETHOD.


  METHOD get_bottom_offset.
    rp_result = me->bottom_offset.
  ENDMETHOD.


  METHOD get_bottom_row.
    rp_result = me->topleft_row + me->height - 1.  "starting from 0
  ENDMETHOD.


  METHOD get_index.
    rp_index = me->index.
  ENDMETHOD.


  METHOD get_left_column.
    rp_result = me->topleft_column - 1.  "starting from 0
  ENDMETHOD.


  METHOD get_left_offset.
    rp_result = me->left_offset.
  ENDMETHOD.


  METHOD get_ref.
    rp_ref = me->ref.
  ENDMETHOD.


  METHOD get_right_column.
    rp_result = me->topleft_column + me->width - 1.  "starting from 0
  ENDMETHOD.


  METHOD get_right_offset.
    rp_result = me->right_offset.
  ENDMETHOD.


  METHOD get_text.
    rp_text = me->text.
  ENDMETHOD.


  METHOD get_top_offset.
    rp_result = me->top_offset.
  ENDMETHOD.


  METHOD get_top_row.
    rp_result = me->topleft_row - 1.  "starting from 0
  ENDMETHOD.


  METHOD set_box.
    "Keeping backwards compatibility (do not use this obsolet method)
    me->topleft_column = is_box-left_column + 1.
    me->topleft_row    = is_box-top_row + 1.
    me->width          = is_box-right_column - is_box-left_column.
    me->height         = is_box-bottom_row - is_box-top_row.
    me->left_offset    = is_box-left_offset.
    me->top_offset     = is_box-top_offset.
    me->right_offset   = is_box-right_offset.
    me->bottom_offset  = is_box-bottom_offset.
  ENDMETHOD.


  METHOD set_text.
    me->text = ip_text.

    IF ip_ref IS NOT INITIAL AND ip_ref <> me->ref.
      me->ref = ip_ref.
      IF ip_topleft IS INITIAL.
        zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING i_columnrow  = me->ref
                                                          IMPORTING e_column_int = me->topleft_column
                                                                    e_row        = me->topleft_row ).
        "By default the comment's box starts right above the referenced cell.
        ADD 1 TO me->topleft_column.
        IF me->topleft_row > 1.
          SUBTRACT 1 FROM me->topleft_row.
        ENDIF.
      ENDIF.
    ENDIF.

    IF ip_topleft IS NOT INITIAL.
      zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING i_columnrow  = ip_topleft
                                                        IMPORTING e_column_int = me->topleft_column
                                                                  e_row        = me->topleft_row ).
    ENDIF.

    IF ip_width > 0.
      me->width = ip_width.
    ENDIF.
    IF ip_height > 0.
      me->height = ip_height.
    ENDIF.

    IF ip_left_offset IS SUPPLIED AND ip_left_offset >= 0.
      me->left_offset = ip_left_offset.
    ENDIF.
    IF ip_top_offset IS SUPPLIED AND ip_top_offset >= 0.
      me->top_offset = ip_top_offset.
    ENDIF.
    IF ip_right_offset IS SUPPLIED AND ip_right_offset >= 0.
      me->right_offset = ip_right_offset.
    ENDIF.
    IF ip_bottom_offset IS SUPPLIED AND ip_bottom_offset >= 0.
      me->bottom_offset = ip_bottom_offset.
    ENDIF.

    "Keeping backwards compatibility (do not use these obsolet parameters)
    IF ip_left_column  IS SUPPLIED OR
       ip_top_row      IS SUPPLIED OR
       ip_right_column IS SUPPLIED OR
       ip_bottom_row   IS SUPPLIED.
      me->topleft_column = ip_left_column + 1.
      me->topleft_row    = ip_top_row + 1.
      me->width          = ip_right_column - ip_left_column.
      me->height         = ip_bottom_row - ip_top_row.
    ENDIF.

  ENDMETHOD.

ENDCLASS.
