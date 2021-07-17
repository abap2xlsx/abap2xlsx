INTERFACE zif_excel_book_vba_project
  PUBLIC .


  DATA vbaproject TYPE xstring READ-ONLY .
  DATA codename TYPE string READ-ONLY .
  DATA codename_pr TYPE string READ-ONLY .

  METHODS set_vbaproject
    IMPORTING
      !ip_vbaproject TYPE xstring .
  METHODS set_codename
    IMPORTING
      !ip_codename TYPE string .
  METHODS set_codename_pr
    IMPORTING
      !ip_codename_pr TYPE string .
ENDINTERFACE.
