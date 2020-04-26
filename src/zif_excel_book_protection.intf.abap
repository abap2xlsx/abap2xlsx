INTERFACE zif_excel_book_protection
  PUBLIC .


  CONSTANTS c_locked TYPE zexcel_book_protection VALUE '1'. "#EC NOTEXT
  CONSTANTS c_protected TYPE zexcel_book_protection VALUE 'X'. "#EC NOTEXT
  CONSTANTS c_unlocked TYPE zexcel_book_protection VALUE '0'. "#EC NOTEXT
  CONSTANTS c_unprotected TYPE zexcel_book_protection VALUE ''. "#EC NOTEXT
  DATA lockrevision TYPE zexcel_book_protection .
  DATA lockstructure TYPE zexcel_book_protection .
  DATA lockwindows TYPE zexcel_book_protection .
  DATA protected TYPE zexcel_book_protection .
  DATA revisionspassword TYPE zexcel_aes_password .
  DATA workbookpassword TYPE zexcel_aes_password .

  METHODS initialize .
ENDINTERFACE.
