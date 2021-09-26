*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL1
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel1.


DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_hyperlink TYPE REF TO zcl_excel_hyperlink,
      lo_column    TYPE REF TO zcl_excel_column.

CONSTANTS: gc_save_file_name TYPE string VALUE '01_HelloWorld.xlsx'.
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
  lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = 'https://sapmentors.github.io/abap2xlsx' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Click here to visit abap2xlsx homepage' ip_hyperlink = lo_hyperlink ).

  lo_worksheet->set_cell( ip_column = 'B' ip_row =  6 ip_value = '你好，世界' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row =  6 ip_value = '(Chinese)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row =  7 ip_value = 'नमस्ते दुनिया' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row =  7 ip_value = '(Hindi)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row =  8 ip_value = 'Hola Mundo' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row =  8 ip_value = '(Spanish)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row =  9 ip_value = 'مرحبا بالعالم' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row =  9 ip_value = '(Arabic)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 10 ip_value = 'ওহে বিশ্ব ' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 10 ip_value = '(Bengali)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 11 ip_value = 'Bonjour le monde' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 11 ip_value = '(French)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 12 ip_value = 'Olá Mundo' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 12 ip_value = '(Portuguese)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 13 ip_value = 'Привет, мир' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 13 ip_value = '(Russian)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 14 ip_value = 'ہیلو دنیا' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 14 ip_value = '(Urdu)' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 15 ip_value = '👋🌎, 👋🌍, 👋🌏' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 15 ip_value = '(Emoji waving hand + 3 parts of the world)' ).

  lo_column = lo_worksheet->get_column( ip_column = 'B' ).
  lo_column->set_width( ip_width = 11 ).



*** Create output
  lcl_output=>output( lo_excel ).
