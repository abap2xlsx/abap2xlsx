CLASS cl_excel_test DEFINITION PUBLIC.
  PUBLIC SECTION.
    CLASS-METHODS run
      RETURNING VALUE(xdata) TYPE xstring
      RAISING cx_static_check.
ENDCLASS.

CLASS cl_excel_test IMPLEMENTATION.
  METHOD run.
    DATA lo_excel     TYPE REF TO zcl_excel.
    DATA lo_worksheet TYPE REF TO zcl_excel_worksheet.
    DATA lo_hyperlink TYPE REF TO zcl_excel_hyperlink.
    DATA lo_column    TYPE REF TO zcl_excel_column.
    DATA lv_date      TYPE d.
    DATA lv_time      TYPE t.
    DATA li_writer    TYPE REF TO zif_excel_writer.

    CREATE OBJECT lo_excel.
    lo_worksheet = lo_excel->get_active_worksheet( ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
    lv_date = '20211231'.
    lv_time = '055817'.
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = lv_date ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = lv_time ).
    lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = 'https://abap2xlsx.github.io/abap2xlsx' ).
    lo_worksheet->set_cell( ip_columnrow = 'B4' ip_value = 'Click here to visit abap2xlsx homepage' ip_hyperlink = lo_hyperlink ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row =  6 ip_value = '你好，世界' ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row =  6 ip_value = '(Chinese)' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row =  7 ip_value = 'नमस्ते दुनिया' ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row =  7 ip_value = '(Hindi)' ).
    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_width( ip_width = 11 ).

    CREATE OBJECT li_writer TYPE zcl_excel_writer_2007.
    xdata = li_writer->write_file( lo_excel ).
  ENDMETHOD.
ENDCLASS.