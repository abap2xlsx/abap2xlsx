*&---------------------------------------------------------------------*
*& Report  ZTEST_EXCEL_IMAGE_HEADER
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT ztest_excel_image_header.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_drawing   TYPE REF TO zcl_excel_drawing,
      ls_key       TYPE wwwdatatab,
      ls_header    TYPE zexcel_s_worksheet_head_foot,
      ls_footer    TYPE zexcel_s_worksheet_head_foot,
      lv_content   TYPE xstring.


DATA: ls_io TYPE skwf_io.

CONSTANTS: gc_save_file_name TYPE string VALUE 'Image_Header_Footer.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

**********************************************************************
*** Header Center
  " create global drawing, set position and media from web repository
  lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_image_header_footer ).

  ls_key-relid = 'MI'.
  ls_key-objid = 'SAPLOGO.GIF'.
  lo_drawing->set_media_www( ip_key = ls_key
                            ip_width = 166
                            ip_height = 75 ).
**********************************************************************
  ls_header-center_image = lo_drawing.


**********************************************************************
*** Header Left
  " create global drawing, set position and media from web repository
  lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_image_header_footer ).

  ls_key-relid = 'MI'.
  ls_key-objid = 'SAPLOGO.GIF'.
  lo_drawing->set_media_www( ip_key = ls_key
                             ip_width = 166
                             ip_height = 75 ).


  ls_header-left_image = ls_footer-left_image = lo_drawing.
  ls_header-left_value = 'Hallo'.
  lo_worksheet = lo_excel->get_active_worksheet( ).

  lo_worksheet->sheet_setup->set_header_footer( ip_odd_header = ls_header
                                                ip_odd_footer = ls_footer ).

**********************************************************************
*** Normal Image
  " create global drawing, set position and media from web repository
  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_position( ip_from_row = 3
                            ip_from_col = 'B' ).

  ls_key-relid = 'MI'.
  ls_key-objid = 'SAPLOGO.GIF'.
  lo_drawing->set_media_www( ip_key = ls_key
                             ip_width = 166
                             ip_height = 75 ).

  " assign drawing to the worksheet
  lo_worksheet->add_drawing( lo_drawing ).

**********************************************************************
**********************************************************************
* New sheet
  lo_worksheet = lo_excel->add_new_worksheet( 'Sheet2' ).

  " Add some content otherwise the error "nothing to be printed" is shown
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'Hello world' ).

**********************************************************************
*** Header Left
  " create global drawing, set position and media from web repository
  lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_image_header_footer ).

  ls_key-relid = 'MI'.
  ls_key-objid = 'SAPLOGO.GIF'.
  lo_drawing->set_media_www( ip_key = ls_key
                             ip_width = 166
                             ip_height = 75 ).


  CLEAR ls_header.
  ls_header-left_image = ls_footer-left_image = lo_drawing.

  lo_worksheet->sheet_setup->set_header_footer( ip_odd_header = ls_header ).


*** Create output
  lcl_output=>output( lo_excel ).
