*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL_COMMENTS
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*
REPORT  zdemo_excel_comments.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_comment   TYPE REF TO zcl_excel_comment,
      lo_hyperlink TYPE REF TO zcl_excel_hyperlink,
      lv_comment   TYPE string.

CONSTANTS: gc_save_file_name TYPE string VALUE 'Comments.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.
  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).

  " Comments
  lo_comment = lo_excel->add_new_comment( ).
  lo_comment->set_text( ip_ref = 'B13' ip_text = 'This is how it begins to be debug time...' ).
  lo_worksheet->add_comment( lo_comment ).
  lo_comment = lo_excel->add_new_comment( ).
  lo_comment->set_text( ip_ref = 'C18' ip_text = 'Another comment' ).
  lo_worksheet->add_comment( lo_comment ).
  lo_comment = lo_excel->add_new_comment( ).
  CONCATENATE 'A comment split' cl_abap_char_utilities=>cr_lf 'on 2 lines?' INTO lv_comment.
  lo_comment->set_text( ip_ref = 'F6' ip_text = lv_comment ).

  " Second sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).

  lo_comment = lo_excel->add_new_comment( ).
  lo_comment->set_text( ip_ref = 'A8' ip_text = 'What about a comment on second sheet?' ).
  lo_worksheet->add_comment( lo_comment ).

  lo_excel->set_active_sheet_index_by_name( 'Sheet1' ).

*** Create output
  lcl_output=>output( lo_excel ).
