*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL43
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel43.

"
"Locally created Structure, which should be equal to the excels structure
"
TYPES: BEGIN OF lty_excel_s,
    dummy TYPE dummy.
TYPES: END OF lty_excel_s.

DATA lt_tab TYPE TABLE OF lty_excel_s.
DATA: lt_filetable TYPE filetable,
      ls_filetable TYPE file_table.
DATA lv_subrc TYPE i.
DATA: lo_excel      TYPE REF TO zcl_excel,
      lo_reader     TYPE REF TO zif_excel_reader,
      lo_worksheet  TYPE REF TO zcl_excel_worksheet,
      lo_salv       TYPE REF TO cl_salv_table.

"
"Ask User to choose a path
"
cl_gui_frontend_services=>file_open_dialog( EXPORTING window_title = 'Excel selection'
                                                      file_filter = '*.xlsx'
                                                      multiselection = abap_false
                                            CHANGING  file_table = lt_filetable " Tabelle, die selektierte Dateien enthält
                                                      rc = lv_subrc
                                            EXCEPTIONS file_open_dialog_failed = 1
                                                       cntl_error = 2
                                                       error_no_gui = 3
                                                       not_supported_by_gui = 4
                                                       OTHERS = 5 ).
IF sy-subrc <> 0.
  MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
  WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
ELSE.
  CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
  TRY.
    LOOP AT lt_filetable INTO ls_filetable.
      lo_excel =  lo_reader->load_file( ls_filetable-filename ).
      lo_worksheet = lo_excel->get_worksheet_by_index( iv_index = 1 ).
      lo_worksheet->get_table( IMPORTING et_table = lt_tab ).
    ENDLOOP.
  ENDTRY.
ENDIF.
"
"Do the presentation stuff
"

cl_salv_table=>factory( IMPORTING r_salv_table = lo_salv
                        CHANGING t_table = lt_tab ).
lo_salv->display( ).
