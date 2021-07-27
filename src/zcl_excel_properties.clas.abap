CLASS zcl_excel_properties DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    DATA creator TYPE zexcel_creator VALUE 'Ivan Femia'. "#EC NOTEXT .  .  . " .
    DATA lastmodifiedby TYPE zexcel_creator VALUE 'Ivan Femia'. "#EC NOTEXT .  .  . " .
    DATA created TYPE timestampl .
    DATA modified TYPE timestampl .
    DATA title TYPE zexcel_title VALUE 'abap2xlsx'. "#EC NOTEXT .  .  . " .
    DATA subject TYPE zexcel_subject .
    DATA description TYPE zexcel_description VALUE 'Created using abap2xlsx'. "#EC NOTEXT .  .  . " .
    DATA keywords TYPE zexcel_keywords .
    DATA category TYPE zexcel_category .
    DATA company TYPE zexcel_company VALUE 'abap2xlsx'. "#EC NOTEXT .  .  . " .
    DATA application TYPE zexcel_application VALUE 'Microsoft Excel'. "#EC NOTEXT .  .  . " .
    DATA docsecurity TYPE zexcel_docsecurity VALUE '0'. "#EC NOTEXT .  .  . " .
    DATA scalecrop TYPE zexcel_scalecrop VALUE ''. "#EC NOTEXT .  .  . " .
    DATA linksuptodate TYPE flag .
    DATA shareddoc TYPE flag .
    DATA hyperlinkschanged TYPE flag .
    DATA appversion TYPE zexcel_appversion VALUE '12.0000'. "#EC NOTEXT .  .  . " .

    METHODS constructor .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_properties IMPLEMENTATION.


  METHOD constructor.

    DATA: lv_timestamp TYPE timestampl.

    GET TIME STAMP FIELD lv_timestamp.
    created   = lv_timestamp.
    modified  = lv_timestamp.

  ENDMETHOD.
ENDCLASS.
