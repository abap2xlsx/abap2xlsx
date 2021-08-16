INTERFACE zif_excel_book_properties
  PUBLIC .

  TYPES tv_excel_appversion TYPE c LENGTH 7.

  DATA creator TYPE zexcel_creator .
  DATA lastmodifiedby TYPE zexcel_creator .
  DATA created TYPE timestampl .
  DATA modified TYPE timestampl .
  DATA title TYPE zexcel_title .
  DATA subject TYPE zexcel_subject .
  DATA description TYPE zexcel_description .
  DATA keywords TYPE zexcel_keywords .
  DATA category TYPE zexcel_category .
  DATA company TYPE zexcel_company .
  DATA application TYPE zexcel_application .
  DATA docsecurity TYPE zexcel_docsecurity .
  DATA scalecrop TYPE zexcel_scalecrop .
  DATA linksuptodate TYPE flag .
  DATA shareddoc TYPE flag .
  DATA hyperlinkschanged TYPE flag .
  DATA appversion TYPE tv_excel_appversion .

  METHODS initialize .
ENDINTERFACE.
