interface ZIF_EXCEL_SHEET_PRINTSETTINGS
  public .


  constants GCV_PRINT_TITLE_NAME type STRING value '_xlnm.Print_Titles'. "#EC NOTEXT

  methods SET_PRINT_REPEAT_COLUMNS
    importing
      !IV_COLUMNS_FROM type ZEXCEL_CELL_COLUMN_ALPHA
      !IV_COLUMNS_TO type ZEXCEL_CELL_COLUMN_ALPHA
    raising
      ZCX_EXCEL .
  methods SET_PRINT_REPEAT_ROWS
    importing
      !IV_ROWS_FROM type ZEXCEL_CELL_ROW
      !IV_ROWS_TO type ZEXCEL_CELL_ROW
    raising
      ZCX_EXCEL .
  methods GET_PRINT_REPEAT_COLUMNS
    exporting
      !EV_COLUMNS_FROM type ZEXCEL_CELL_COLUMN_ALPHA
      !EV_COLUMNS_TO type ZEXCEL_CELL_COLUMN_ALPHA .
  methods GET_PRINT_REPEAT_ROWS
    exporting
      !EV_ROWS_FROM type ZEXCEL_CELL_ROW
      !EV_ROWS_TO type ZEXCEL_CELL_ROW .
  methods CLEAR_PRINT_REPEAT_COLUMNS .
  methods CLEAR_PRINT_REPEAT_ROWS .
endinterface.
