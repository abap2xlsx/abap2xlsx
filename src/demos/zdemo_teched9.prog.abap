*&---------------------------------------------------------------------*
*& Report  ZDEMO_TECHED3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_teched9.

*******************************
*   Data Object declaration   *
*******************************

DATA: lo_excel        TYPE REF TO zcl_excel,
      lo_excel_writer TYPE REF TO zif_excel_writer,
      lo_worksheet    TYPE REF TO zcl_excel_worksheet.

DATA: lo_style_title       TYPE REF TO zcl_excel_style,
      lo_style_green       TYPE REF TO zcl_excel_style,
      lo_style_yellow      TYPE REF TO zcl_excel_style,
      lo_style_red         TYPE REF TO zcl_excel_style,
      lo_drawing           TYPE REF TO zcl_excel_drawing,
      lo_range             TYPE REF TO zcl_excel_range,
      lo_data_validation   TYPE REF TO zcl_excel_data_validation,
      lo_column            TYPE REF TO zcl_excel_column,
      lo_style_conditional TYPE REF TO zcl_excel_style_cond,
      lv_style_title_guid  TYPE zexcel_cell_style,
      lv_style_green_guid  TYPE zexcel_cell_style,
      lv_style_yellow_guid TYPE zexcel_cell_style,
      lv_style_red_guid    TYPE zexcel_cell_style,
      ls_cellis            TYPE zexcel_conditional_cellis,
      ls_key               TYPE wwwdatatab.

DATA: lo_send_request TYPE REF TO cl_bcs,
      lo_document     TYPE REF TO cl_document_bcs,
      lo_sender       TYPE REF TO cl_sapuser_bcs,
      lo_recipient    TYPE REF TO cl_sapuser_bcs.

DATA: lv_file        TYPE xstring,
      lv_bytecount   TYPE i,
      lv_bytecount_c TYPE sood-objlen,
      lt_file_tab    TYPE solix_tab.

*******************************
*    abap2xlsx create XLSX    *
*******************************

" Create excel instance
CREATE OBJECT lo_excel.

" Styles
lo_style_title                   = lo_excel->add_new_style( ).
lo_style_title->font->bold       = abap_true.
lo_style_title->font->color-rgb  = zcl_excel_style_color=>c_blue.
lv_style_title_guid              = lo_style_title->get_guid( ).

" Get active sheet
lo_worksheet        = lo_excel->get_active_worksheet( ).
lo_worksheet->set_title( ip_title = 'Demo TechEd' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 5 ip_value = 'TechEd demo' ip_style = lv_style_title_guid ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 7 ip_value = 'Is abap2xlsx simple' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 8 ip_value = 'Is abap2xlsx CooL' ).

lo_worksheet->set_cell( ip_column = 'B' ip_row = 10 ip_value = 'Total score' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 10 ip_formula = 'SUM(C7:C8)' ).

" add logo from SMWO
lo_drawing = lo_excel->add_new_drawing( ).
lo_drawing->set_position( ip_from_row = 2
                          ip_from_col = 'B' ).

ls_key-relid = 'MI'.
ls_key-objid = 'SIWB_KW_LOGO'.
lo_drawing->set_media_www( ip_key = ls_key
                           ip_width = 140
                           ip_height = 64 ).

" assign drawing to the worksheet
lo_worksheet->add_drawing( lo_drawing ).

" Add new sheet
lo_worksheet = lo_excel->add_new_worksheet( ).
lo_worksheet->set_title( ip_title = 'Values' ).

" Set values for range
lo_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 1 ).
lo_worksheet->set_cell( ip_row = 5 ip_column = 'A' ip_value = 2 ).
lo_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 3 ).
lo_worksheet->set_cell( ip_row = 7 ip_column = 'A' ip_value = 4 ).
lo_worksheet->set_cell( ip_row = 8 ip_column = 'A' ip_value = 5 ).

lo_range            = lo_excel->add_new_range( ).
lo_range->name      = 'Values'.
lo_range->set_value( ip_sheet_name    = 'Values'
                     ip_start_column  = 'A'
                     ip_start_row     = 4
                     ip_stop_column   = 'A'
                     ip_stop_row      = 8 ).

lo_excel->set_active_sheet_index( 1 ).

" add data validation
lo_worksheet        = lo_excel->get_active_worksheet( ).

lo_data_validation              = lo_worksheet->add_new_data_validation( ).
lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
lo_data_validation->formula1    = 'Values'.
lo_data_validation->cell_row    = 7.
lo_data_validation->cell_column = 'C'.
lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 'Select a value' ).


lo_data_validation              = lo_worksheet->add_new_data_validation( ).
lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
lo_data_validation->formula1    = 'Values'.
lo_data_validation->cell_row    = 8.
lo_data_validation->cell_column = 'C'.
lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 'Select a value' ).

" add autosize (column width)
lo_column = lo_worksheet->get_column( ip_column = 'B' ).
lo_column->set_auto_size( ip_auto_size = abap_true ).
lo_column = lo_worksheet->get_column( ip_column = 'C' ).
lo_column->set_auto_size( ip_auto_size = abap_true ).

" defne conditional styles
lo_style_green                        = lo_excel->add_new_style( ).
lo_style_green->fill->filltype        = zcl_excel_style_fill=>c_fill_solid.
lo_style_green->fill->bgcolor-rgb     = zcl_excel_style_color=>c_green.
lv_style_green_guid                   = lo_style_green->get_guid( ).

lo_style_yellow                        = lo_excel->add_new_style( ).
lo_style_yellow->fill->filltype        = zcl_excel_style_fill=>c_fill_solid.
lo_style_yellow->fill->bgcolor-rgb     = zcl_excel_style_color=>c_yellow.
lv_style_yellow_guid                   = lo_style_yellow->get_guid( ).

lo_style_red                        = lo_excel->add_new_style( ).
lo_style_red->fill->filltype        = zcl_excel_style_fill=>c_fill_solid.
lo_style_red->fill->bgcolor-rgb     = zcl_excel_style_color=>c_red.
lv_style_red_guid                   = lo_style_red->get_guid( ).

" add conditional formatting
lo_style_conditional = lo_worksheet->add_new_style_cond( ).
lo_style_conditional->rule        = zcl_excel_style_cond=>c_rule_cellis.
ls_cellis-formula                 = '5'.
ls_cellis-operator                = zcl_excel_style_cond=>c_operator_greaterthan.
ls_cellis-cell_style              = lv_style_green_guid.
lo_style_conditional->mode_cellis = ls_cellis.
lo_style_conditional->priority    = 1.
lo_style_conditional->set_range( ip_start_column  = 'C'
                                 ip_start_row     = 10
                                 ip_stop_column   = 'C'
                                 ip_stop_row      = 10 ).

lo_style_conditional = lo_worksheet->add_new_style_cond( ).
lo_style_conditional->rule        = zcl_excel_style_cond=>c_rule_cellis.
ls_cellis-formula                 = '5'.
ls_cellis-operator                = zcl_excel_style_cond=>c_operator_equal.
ls_cellis-cell_style              = lv_style_yellow_guid.
lo_style_conditional->mode_cellis = ls_cellis.
lo_style_conditional->priority    = 2.
lo_style_conditional->set_range( ip_start_column  = 'C'
                                 ip_start_row     = 10
                                 ip_stop_column   = 'C'
                                 ip_stop_row      = 10 ).

lo_style_conditional = lo_worksheet->add_new_style_cond( ).
lo_style_conditional->rule        = zcl_excel_style_cond=>c_rule_cellis.
ls_cellis-formula                 = '0'.
ls_cellis-operator                = zcl_excel_style_cond=>c_operator_greaterthan.
ls_cellis-cell_style              = lv_style_red_guid.
lo_style_conditional->mode_cellis = ls_cellis.
lo_style_conditional->priority    = 3.
lo_style_conditional->set_range( ip_start_column  = 'C'
                                 ip_start_row     = 10
                                 ip_stop_column   = 'C'
                                 ip_stop_row      = 10 ).


" Create xlsx stream
CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
lv_file = lo_excel_writer->write_file( lo_excel ).

*******************************
*            Output           *
*******************************

" Convert to binary
lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
lv_bytecount = xstrlen( lv_file ).
lv_bytecount_c = lv_bytecount.

" Send via email
lo_document = cl_document_bcs=>create_document( i_type    = 'RAW'
                                                i_subject = 'Demo TechEd' ).

lo_document->add_attachment( i_attachment_type    = 'EXT'
                             i_attachment_subject = 'abap2xlsx.xlsx'
                             i_attachment_size    = lv_bytecount_c
                             i_att_content_hex    = lt_file_tab ).

lo_sender       = cl_sapuser_bcs=>create( sy-uname ).
lo_recipient    = cl_sapuser_bcs=>create( sy-uname ).

lo_send_request = cl_bcs=>create_persistent( ).
lo_send_request->set_document( lo_document ).
lo_send_request->set_sender( lo_sender ).
lo_send_request->add_recipient( lo_recipient ).
lo_send_request->set_send_immediately( abap_true ).
lo_send_request->send( ).
