INTERFACE zif_excel_book_protection
  PUBLIC .

  TYPES tv_book_protection TYPE c LENGTH 1.

  CONSTANTS c_locked TYPE tv_book_protection VALUE '1'. "#EC NOTEXT
  CONSTANTS c_protected TYPE tv_book_protection VALUE 'X'.  "#EC NOTEXT
  CONSTANTS c_unlocked TYPE tv_book_protection VALUE '0'. "#EC NOTEXT
  CONSTANTS c_unprotected TYPE tv_book_protection VALUE ''. "#EC NOTEXT
  DATA lockrevision TYPE tv_book_protection .
  DATA lockstructure TYPE tv_book_protection .
  DATA lockwindows TYPE tv_book_protection .
  DATA protected TYPE tv_book_protection .
  DATA revisionspassword TYPE zexcel_aes_password .
  DATA workbookpassword TYPE zexcel_aes_password .

  METHODS initialize .
ENDINTERFACE.
