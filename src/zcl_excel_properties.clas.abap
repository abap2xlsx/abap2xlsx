class ZCL_EXCEL_PROPERTIES definition
  public
  final
  create public .

public section.

  data CREATOR type ZEXCEL_CREATOR value 'Ivan Femia'. "#EC NOTEXT .  .  . " .
  data LASTMODIFIEDBY type ZEXCEL_CREATOR value 'Ivan Femia'. "#EC NOTEXT .  .  . " .
  data CREATED type TIMESTAMPL .
  data MODIFIED type TIMESTAMPL .
  data TITLE type ZEXCEL_TITLE value 'abap2xlsx'. "#EC NOTEXT .  .  . " .
  data SUBJECT type ZEXCEL_SUBJECT .
  data DESCRIPTION type ZEXCEL_DESCRIPTION value 'Created using abap2xlsx'. "#EC NOTEXT .  .  . " .
  data KEYWORDS type ZEXCEL_KEYWORDS .
  data CATEGORY type ZEXCEL_CATEGORY .
  data COMPANY type ZEXCEL_COMPANY value 'abap2xlsx'. "#EC NOTEXT .  .  . " .
  data APPLICATION type ZEXCEL_APPLICATION value 'Microsoft Excel'. "#EC NOTEXT .  .  . " .
  data DOCSECURITY type ZEXCEL_DOCSECURITY value '0'. "#EC NOTEXT .  .  . " .
  data SCALECROP type ZEXCEL_SCALECROP value ''. "#EC NOTEXT .  .  . " .
  data LINKSUPTODATE type FLAG .
  data SHAREDDOC type FLAG .
  data HYPERLINKSCHANGED type FLAG .
  data APPVERSION type ZEXCEL_APPVERSION value '12.0000'. "#EC NOTEXT .  .  . " .

  methods CONSTRUCTOR .
protected section.
private section.
ENDCLASS.



CLASS ZCL_EXCEL_PROPERTIES IMPLEMENTATION.


method CONSTRUCTOR.

  DATA: lv_timestamp TYPE timestampl.

  GET TIME STAMP FIELD lv_timestamp.
  created   = lv_timestamp.
  modified  = lv_timestamp.

endmethod.
ENDCLASS.
