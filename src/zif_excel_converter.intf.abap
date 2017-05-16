interface ZIF_EXCEL_CONVERTER
  public .


  methods CAN_CONVERT_OBJECT
    importing
      !IO_OBJECT type ref to OBJECT
    raising
      ZCX_EXCEL .
  methods CREATE_FIELDCATALOG
    importing
      !IS_OPTION type ZEXCEL_S_CONVERTER_OPTION
      !IO_OBJECT type ref to OBJECT
      !IT_TABLE type STANDARD TABLE
    exporting
      !ES_LAYOUT type ZEXCEL_S_CONVERTER_LAYO
      !ET_FIELDCATALOG type ZEXCEL_T_CONVERTER_FCAT
      !EO_TABLE type ref to DATA
      !ET_COLORS type ZEXCEL_T_CONVERTER_COL
      !ET_FILTER type ZEXCEL_T_CONVERTER_FIL
    raising
      ZCX_EXCEL .
endinterface.
