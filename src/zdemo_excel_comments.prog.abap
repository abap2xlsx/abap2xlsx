*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL_COMMENTS
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*
REPORT  zdemo_excel_comments.

CONSTANTS: gc_save_file_name TYPE string VALUE 'Comments.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

CLASS lcl_repctl DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS:
      start_of_selection.
  PRIVATE SECTION.

    CLASS-DATA:
      go_excel     TYPE REF TO zcl_excel.

    CLASS-METHODS:
      add_normal_comments RAISING zcx_excel,
      add_rich_text_comments RAISING zcx_excel.

ENDCLASS.

START-OF-SELECTION.
  lcl_repctl=>start_of_selection( ).

CLASS lcl_repctl IMPLEMENTATION.
  METHOD start_of_selection.

    DATA:
      lo_ex TYPE REF TO zcx_excel.


    TRY.

        " Creates active sheet
        CREATE OBJECT go_excel.

        add_normal_comments(  ).

        add_rich_text_comments(  ).

        go_excel->set_active_sheet_index( 1 ).

*** Create output
        lcl_output=>output( go_excel ).

      CATCH zcx_excel INTO lo_ex.
        MESSAGE lo_ex TYPE 'I'.
    ENDTRY.
  ENDMETHOD.

  METHOD add_normal_comments.

    DATA: lo_comment   TYPE REF TO zcl_excel_comment,
          lo_worksheet TYPE REF TO zcl_excel_worksheet,
          lv_comment   TYPE string.

    " Get active sheet
    lo_worksheet = go_excel->get_active_worksheet( ).

    " Comments
    lo_comment = go_excel->add_new_comment( ).
    lo_comment->set_text( ip_ref = 'B13' ip_text = 'This is how it begins to be debug time...' ).
    lo_worksheet->add_comment( lo_comment ).
    lo_comment = go_excel->add_new_comment( ).
    lo_comment->set_text( ip_ref = 'C18' ip_text = 'Another comment' ).
    lo_worksheet->add_comment( lo_comment ).
    lo_comment = go_excel->add_new_comment( ).
    CONCATENATE 'A comment split' cl_abap_char_utilities=>cr_lf 'on 2 lines?' INTO lv_comment.
    lo_comment->set_text( ip_ref = 'F6' ip_text = lv_comment ).

    " Second sheet
    lo_worksheet = go_excel->add_new_worksheet( ).

    lo_comment = go_excel->add_new_comment( ).
    lo_comment->set_text( ip_ref = 'A8' ip_text = 'What about a comment on second sheet?' ).
    lo_worksheet->add_comment( lo_comment ).

  ENDMETHOD.

  METHOD add_rich_text_comments.

    DATA:
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_style     TYPE REF TO zcl_excel_style,
      lo_comment   TYPE REF TO zcl_excel_comment,
      lv_value     TYPE string,
      lt_rtf       TYPE zexcel_t_rtf,
      ls_rtf       LIKE LINE OF lt_rtf,
      ls_font      LIKE ls_rtf-font.

    lo_worksheet = go_excel->get_worksheet_by_index( 1 ).

    lo_style = go_excel->add_new_style( ).
    ls_font = lo_style->font->get_structure( ).  "Default settings from ZCL_EXCEL_STYLE_FONT's Constructor

    lv_value = 'normal red underline normal red-underline bold italic bigger Times-New-Roman'.

    " red
    ls_rtf-font = ls_font.
    ls_rtf-font-color-rgb = zcl_excel_style_color=>c_red.
    ls_rtf-offset = 7.
    ls_rtf-length = 3.
    INSERT ls_rtf INTO TABLE lt_rtf.

    " underline
    ls_rtf-font = ls_font.
    ls_rtf-font-underline = abap_true.
    ls_rtf-font-underline_mode = zcl_excel_style_font=>c_underline_single.
    ls_rtf-offset = 11.
    ls_rtf-length = 9.
    INSERT ls_rtf INTO TABLE lt_rtf.

    " red and underline
    ls_rtf-font = ls_font.
    ls_rtf-font-color-rgb = zcl_excel_style_color=>c_red.
    ls_rtf-font-underline = abap_true.
    ls_rtf-font-underline_mode = zcl_excel_style_font=>c_underline_single.
    ls_rtf-offset = 28.
    ls_rtf-length = 13.
    INSERT ls_rtf INTO TABLE lt_rtf.

    " bold
    ls_rtf-font = ls_font.
    ls_rtf-font-bold = abap_true.
    ls_rtf-offset = 42.
    ls_rtf-length = 4.
    INSERT ls_rtf INTO TABLE lt_rtf.

    " italic
    ls_rtf-font = ls_font.
    ls_rtf-font-italic = abap_true.
    ls_rtf-offset = 47.
    ls_rtf-length = 6.
    INSERT ls_rtf INTO TABLE lt_rtf.

    " bigger
    ls_rtf-font = ls_font.
    ls_rtf-font-size = 28.
    ls_rtf-offset = 54.
    ls_rtf-length = 6.
    INSERT ls_rtf INTO TABLE lt_rtf.

    " Times-New-Roman
    ls_rtf-font = ls_font.
    ls_rtf-font-name            = zcl_excel_style_font=>c_name_roman.
    ls_rtf-font-scheme          = zcl_excel_style_font=>c_scheme_none.
    ls_rtf-font-family          = zcl_excel_style_font=>c_family_roman.
    " Create an underline double style
    ls_rtf-font-underline       = abap_true.
    ls_rtf-font-underline_mode  = zcl_excel_style_font=>c_underline_double.
    ls_rtf-offset = 61.
    ls_rtf-length = 15.
    INSERT ls_rtf INTO TABLE lt_rtf.

    lo_worksheet->set_cell(
      ip_column = 'B'
      ip_row    = 6
      ip_style  = lo_style->get_guid( )
      ip_value  = lv_value
      it_rtf    = lt_rtf ).

    lo_comment = go_excel->add_new_comment( ).
    lo_comment->set_text( ip_ref          = 'B6'
                          ip_text         = lv_value
                          it_rtf          = lt_rtf
                          ip_right_column = 12
                          ip_bottom_row   = 16 ).
    lo_worksheet->add_comment( lo_comment ).
  ENDMETHOD.

ENDCLASS.
