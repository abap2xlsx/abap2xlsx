INTERFACE zif_excel_sheet_protection
  PUBLIC .

  TYPES tv_sheet_protection TYPE c LENGTH 1.
  TYPES tv_sheet_protection_bool TYPE c LENGTH 1.

  DATA auto_filter TYPE tv_sheet_protection_bool .
  CONSTANTS c_active TYPE tv_sheet_protection_bool VALUE '1'. "#EC NOTEXT
  CONSTANTS c_noactive TYPE tv_sheet_protection_bool VALUE '0'. "#EC NOTEXT
  CONSTANTS c_protected TYPE tv_sheet_protection VALUE 'X'. "#EC NOTEXT
  CONSTANTS c_unprotected TYPE tv_sheet_protection VALUE ''. "#EC NOTEXT
  DATA delete_columns TYPE tv_sheet_protection_bool .
  DATA delete_rows TYPE tv_sheet_protection_bool .
  DATA format_cells TYPE tv_sheet_protection_bool .
  DATA format_columns TYPE tv_sheet_protection_bool .
  DATA format_rows TYPE tv_sheet_protection_bool .
  DATA insert_columns TYPE tv_sheet_protection_bool .
  DATA insert_hyperlinks TYPE tv_sheet_protection_bool .
  DATA insert_rows TYPE tv_sheet_protection_bool .
  DATA objects TYPE tv_sheet_protection_bool .
  DATA password TYPE zexcel_aes_password .
  DATA pivot_tables TYPE tv_sheet_protection_bool .
  DATA protected TYPE tv_sheet_protection .
  DATA scenarios TYPE tv_sheet_protection_bool .
  DATA select_locked_cells TYPE tv_sheet_protection_bool .
  DATA select_unlocked_cells TYPE tv_sheet_protection_bool .
  DATA sheet TYPE tv_sheet_protection_bool .
  DATA sort TYPE tv_sheet_protection_bool .

  METHODS initialize .
ENDINTERFACE.
