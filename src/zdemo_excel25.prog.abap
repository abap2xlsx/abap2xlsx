*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL25
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel25.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_exception            TYPE REF TO cx_root.

DATA: lv_file                 TYPE xstring.

CONSTANTS: lv_file_name TYPE string VALUE '25_HelloWorld.xlsx'.
DATA: lv_default_file_name TYPE string.
DATA: lv_error TYPE string.

CALL FUNCTION 'FILE_GET_NAME_USING_PATH'
  EXPORTING
    logical_path        = 'LOCAL_TEMPORARY_FILES'  " Logical path'
    file_name           = lv_file_name    " File name
  IMPORTING
    file_name_with_path = lv_default_file_name.    " File name with path
" Creates active sheet
CREATE OBJECT lo_excel.

" Get active sheet
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_title( ip_title = 'Sheet1' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).

CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
lv_file = lo_excel_writer->write_file( lo_excel ).

TRY.
    OPEN DATASET lv_default_file_name FOR OUTPUT IN BINARY MODE.
    TRANSFER lv_file  TO lv_default_file_name.
    CLOSE DATASET lv_default_file_name.
  CATCH cx_root INTO lo_exception.
    lv_error = lo_exception->get_text( ).
    MESSAGE lv_error TYPE 'I'.
ENDTRY.
