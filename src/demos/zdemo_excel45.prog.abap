*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL45
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel45.

CONSTANTS:
  gc_ws_title_validation TYPE zexcel_sheet_title VALUE 'Validation'.

DATA:
  lo_excel             TYPE REF TO zcl_excel,
  lo_worksheet         TYPE REF TO zcl_excel_worksheet,
  lo_range             TYPE REF TO zcl_excel_range,
  lv_validation_string TYPE string,
  lo_data_validation   TYPE REF TO zcl_excel_data_validation,
  lv_row               TYPE zexcel_cell_row.


CONSTANTS:
  gc_save_file_name TYPE string VALUE '45_ShowDropdown.xlsx'.

INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

*** Sheet Admin

* Creates active sheet
  CREATE OBJECT lo_excel.

* Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).

* Set sheet name "Validation"
  lo_worksheet->set_title( gc_ws_title_validation ).


* short validations can be entered as string (<254Char)
  lv_validation_string = '"New York, Rio, Tokyo"'.

* create validation object
  lo_data_validation = lo_worksheet->add_new_data_validation( ).

* create new validation from validation string
  lo_data_validation->type           = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1       = lv_validation_string.
  lo_data_validation->cell_row       = 2.
  lo_data_validation->cell_row_to    = 4.
  lo_data_validation->cell_column    = 'A'.
  lo_data_validation->cell_column_to = 'A'.
  lo_data_validation->allowblank     = 'X'.
  lo_data_validation->showdropdown   = 'X'.

* add some fields with validation
  lv_row = 2.
  WHILE lv_row <= 4.
    lo_worksheet->set_cell( ip_row = lv_row ip_column = 'A' ip_value = 'Select' ).
    lv_row = lv_row + 1.
  ENDWHILE.

*** Create output
  lcl_output=>output( lo_excel ).
