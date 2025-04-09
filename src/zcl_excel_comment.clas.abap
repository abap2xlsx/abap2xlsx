CLASS zcl_excel_comment DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF ty_rtf_fragment.
        INCLUDE TYPE zexcel_s_style_font AS rtf.
    TYPES:
        text TYPE string,
      END OF ty_rtf_fragment .
    TYPES:
      ty_rtf_fragments TYPE STANDARD TABLE OF ty_rtf_fragment
                WITH NON-UNIQUE DEFAULT KEY .
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

    CLASS-METHODS get_default_style
      RETURNING
        VALUE(es_default) TYPE zexcel_s_style_font .
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
        !ip_text          TYPE string OPTIONAL
        !is_style         TYPE zexcel_s_style_font OPTIONAL
        !ip_ref           TYPE string OPTIONAL
        !ip_left_column   TYPE i DEFAULT gc_default_box-left_column
        !ip_left_offset   TYPE i DEFAULT gc_default_box-left_offset
        !ip_top_row       TYPE i DEFAULT gc_default_box-top_row
        !ip_top_offset    TYPE i DEFAULT gc_default_box-top_offset
        !ip_right_column  TYPE i DEFAULT gc_default_box-right_column
        !ip_right_offset  TYPE i DEFAULT gc_default_box-right_offset
        !ip_bottom_row    TYPE i DEFAULT gc_default_box-bottom_row
        !ip_bottom_offset TYPE i DEFAULT gc_default_box-bottom_offset .
    METHODS get_text_rtf
      RETURNING
        VALUE(et_rtf) TYPE ty_rtf_fragments .
    METHODS set_text_rtf
      IMPORTING
        !it_rtf TYPE ty_rtf_fragments OPTIONAL
        !ip_ref TYPE string OPTIONAL
        !is_box TYPE ty_box OPTIONAL .
  PROTECTED SECTION.
PRIVATE SECTION.

  DATA index TYPE string .
  DATA ref TYPE string .
  DATA gt_rtf TYPE ty_rtf_fragments .
  DATA gs_box TYPE ty_box .

  METHODS add_text
    IMPORTING
      !ip_text  TYPE string
      !is_style TYPE zexcel_s_style_font .
ENDCLASS.



CLASS ZCL_EXCEL_COMMENT IMPLEMENTATION.


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
    FIELD-SYMBOLS: <ls_rtf> LIKE LINE OF gt_rtf.
    LOOP AT gt_rtf ASSIGNING <ls_rtf>.
      CONCATENATE rp_text <ls_rtf>-text INTO rp_text.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_top_offset.
    rp_result = gs_box-top_offset.
  ENDMETHOD.


  METHOD get_top_row.
    rp_result = gs_box-top_row.
  ENDMETHOD.


  METHOD set_text.

    IF ip_ref IS SUPPLIED.
      ref = ip_ref.
    ENDIF.

* Add a simple text with parameter IP_TEXT and style IS_STYLE
    IF ip_text IS NOT INITIAL.
      add_text(
        ip_text  = ip_text
        is_style = is_style ).
    ENDIF.

* Parameters of the containing box
    DATA ls_box TYPE ty_box.
    ls_box-left_column   = ip_left_column.
    ls_box-left_offset   = ip_left_offset.
    ls_box-top_row       = ip_top_row.
    ls_box-top_offset    = ip_top_offset.
    ls_box-right_column  = ip_right_column.
    ls_box-right_offset  = ip_right_offset.
    ls_box-bottom_row    = ip_bottom_row.
    ls_box-bottom_offset = ip_bottom_offset.
    set_box( ls_box ).

  ENDMETHOD.


  METHOD set_box.

    gs_box = is_box.

  ENDMETHOD.


  METHOD add_text.

    DATA ls_rtf LIKE LINE OF gt_rtf.
    ls_rtf-text = ip_text.
    IF is_style IS INITIAL.
      ls_rtf-rtf = get_default_style( ).
    ELSE.
      ls_rtf-rtf = is_style.
    ENDIF.
    APPEND ls_rtf TO gt_rtf.

  ENDMETHOD.


  METHOD get_default_style.

    es_default-bold            = abap_true.
    es_default-size            = 9.
    es_default-color-indexed   = 81.
    es_default-color-theme     = zcl_excel_style_color=>c_theme_not_set.
    es_default-name            = `Tahoma`.
    es_default-family          = 2.

  ENDMETHOD.


  method GET_TEXT_RTF.
    et_rtf = gt_rtf.
  endmethod.


  METHOD set_text_rtf.

* Set a text, consisting of differently styled parts
    gt_rtf = it_rtf.

    IF ip_ref IS SUPPLIED.
      ref = ip_ref.
    ENDIF.

* Parameters of the containing box
    IF is_box IS SUPPLIED.
      set_box( is_box ).
    ENDIF.

  ENDMETHOD.
ENDCLASS.
