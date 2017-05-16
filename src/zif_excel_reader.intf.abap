interface ZIF_EXCEL_READER
  public .


  methods LOAD_FILE
    importing
      !I_FILENAME type CSEQUENCE
      !I_USE_ALTERNATE_ZIP type SEOCLSNAME default SPACE
      !I_FROM_APPLSERVER type SYBATCH default SY-BATCH
      !IV_ZCL_EXCEL_CLASSNAME type CLIKE optional
    returning
      value(R_EXCEL) type ref to ZCL_EXCEL
    raising
      ZCX_EXCEL .
  methods LOAD
    importing
      !I_EXCEL2007 type XSTRING
      !I_USE_ALTERNATE_ZIP type SEOCLSNAME default SPACE
      !IV_ZCL_EXCEL_CLASSNAME type CLIKE optional
    returning
      value(R_EXCEL) type ref to ZCL_EXCEL
    raising
      ZCX_EXCEL .
endinterface.
