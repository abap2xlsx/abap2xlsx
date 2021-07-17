INTERFACE zif_excel_sheet_printsettings
  PUBLIC .


  CONSTANTS gcv_print_title_name TYPE string VALUE '_xlnm.Print_Titles'. "#EC NOTEXT

  METHODS set_print_repeat_columns
    IMPORTING
      !iv_columns_from TYPE zexcel_cell_column_alpha
      !iv_columns_to   TYPE zexcel_cell_column_alpha
    RAISING
      zcx_excel .
  METHODS set_print_repeat_rows
    IMPORTING
      !iv_rows_from TYPE zexcel_cell_row
      !iv_rows_to   TYPE zexcel_cell_row
    RAISING
      zcx_excel .
  METHODS get_print_repeat_columns
    EXPORTING
      !ev_columns_from TYPE zexcel_cell_column_alpha
      !ev_columns_to   TYPE zexcel_cell_column_alpha .
  METHODS get_print_repeat_rows
    EXPORTING
      !ev_rows_from TYPE zexcel_cell_row
      !ev_rows_to   TYPE zexcel_cell_row .
  METHODS clear_print_repeat_columns .
  METHODS clear_print_repeat_rows .
ENDINTERFACE.
