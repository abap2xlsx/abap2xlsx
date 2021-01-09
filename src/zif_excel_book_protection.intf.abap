INTERFACE zif_excel_book_protection
  PUBLIC .

  TYPES tv_book_protection TYPE c LENGTH 1.

  CONSTANTS c_locked TYPE zexcel_book_protection_bool VALUE '1'. "#EC NOTEXT
  CONSTANTS c_protected TYPE tv_book_protection VALUE 'X'.  "#EC NOTEXT
  CONSTANTS c_unlocked TYPE zexcel_book_protection_bool VALUE '0'. "#EC NOTEXT
  CONSTANTS c_unprotected TYPE tv_book_protection VALUE ''. "#EC NOTEXT
  DATA lockrevision TYPE zexcel_book_protection_bool .
  DATA lockstructure TYPE zexcel_book_protection_bool .
  DATA lockwindows TYPE zexcel_book_protection_bool .
  DATA protected TYPE tv_book_protection .
  DATA revisionspassword TYPE zexcel_aes_password .
  DATA workbookpassword TYPE zexcel_aes_password .

  METHODS initialize .
ENDINTERFACE.
