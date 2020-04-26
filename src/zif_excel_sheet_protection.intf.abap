INTERFACE zif_excel_sheet_protection
  PUBLIC .


  DATA auto_filter TYPE zexcel_sheet_protection_bool .
  CONSTANTS c_active TYPE zexcel_sheet_protection_bool VALUE '1'. "#EC NOTEXT
  CONSTANTS c_noactive TYPE zexcel_sheet_protection_bool VALUE '0'. "#EC NOTEXT
  CONSTANTS c_protected TYPE zexcel_sheet_protection VALUE 'X'. "#EC NOTEXT
  CONSTANTS c_unprotected TYPE zexcel_sheet_protection VALUE ''. "#EC NOTEXT
  DATA delete_columns TYPE zexcel_sheet_protection_bool .
  DATA delete_rows TYPE zexcel_sheet_protection_bool .
  DATA format_cells TYPE zexcel_sheet_protection_bool .
  DATA format_columns TYPE zexcel_sheet_protection_bool .
  DATA format_rows TYPE zexcel_sheet_protection_bool .
  DATA insert_columns TYPE zexcel_sheet_protection_bool .
  DATA insert_hyperlinks TYPE zexcel_sheet_protection_bool .
  DATA insert_rows TYPE zexcel_sheet_protection_bool .
  DATA objects TYPE zexcel_sheet_protection_bool .
  DATA password TYPE zexcel_aes_password .
  DATA pivot_tables TYPE zexcel_sheet_protection_bool .
  DATA protected TYPE zexcel_sheet_protection .
  DATA scenarios TYPE zexcel_sheet_protection_bool .
  DATA select_locked_cells TYPE zexcel_sheet_protection_bool .
  DATA select_unlocked_cells TYPE zexcel_sheet_protection_bool .
  DATA sheet TYPE zexcel_sheet_protection_bool .
  DATA sort TYPE zexcel_sheet_protection_bool .

  METHODS initialize .
ENDINTERFACE.
