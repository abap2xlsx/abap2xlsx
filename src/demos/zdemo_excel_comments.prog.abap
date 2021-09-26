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
*  lo_worksheet->set_title( ip_title = 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = sy-datum ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = sy-uzeit ).
  lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = 'http://www.abap2xlsx.org' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Click here to visit abap2xlsx homepage' ip_hyperlink = lo_hyperlink ).

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
  lo_worksheet->set_default_excel_date_format( zcl_excel_style_number_format=>c_format_date_yyyymmdd ).
  lo_worksheet->set_title( ip_title = 'Sheet2' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Date Format set to YYYYMMDD' ).
  " Insert current date
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Current Date:' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = sy-datum ).

  lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet3!B2' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = 'This is link to the third sheet' ip_hyperlink = lo_hyperlink ).

  lo_comment = lo_excel->add_new_comment( ).
  lo_comment->set_text( ip_ref = 'A8' ip_text = 'What about a comment on second sheet?' ).
  " lo_comment->set_text( ip_column = 'A' ip_row = 8 ip_text = 'What about a comment on second sheet?' ).
  lo_worksheet->add_comment( lo_comment ).

  lo_excel->set_active_sheet_index_by_name( 'Sheet1' ).

*** Create output
  lcl_output=>output( lo_excel ).
