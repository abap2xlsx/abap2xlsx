interface ZIF_EXCEL_BUILDER_TEMPLATE
  public .


  methods ADD_DRAWING
    importing
      !IV_HANDLE type STRING
      !IV_VARIABLE type STRING
      !IO_DRAWING type ref to ZCL_EXCEL_DRAWING .
  methods ON_COMMAND
    importing
      !IV_HANDLE type STRING
      !IV_INDEX type SY-INDEX
      !IV_COMMAND type STRING
    changing
      !CS_CONTEXT type ZCL_EXCEL_BUILDER=>TS_CONTEXT
    returning
      value(RV_NEXT) type FLAG .
  methods ON_END_OF_ROW
    importing
      !IV_HANDLE type STRING
      !IV_INDEX type SY-INDEX
    changing
      !CV_ELEMENT type STRING
    returning
      value(RV_NEXT) type FLAG .
  methods ON_END_OF_SHEET
    importing
      !IV_HANDLE type STRING
      !IV_INDEX type SY-INDEX
    returning
      value(RV_NEXT) type FLAG .
  methods SUBST
    importing
      !IV_VARIABLE type ZEXCEL_CELL_VALUE
      !IV_HANDLE type STRING
      !IV_ROW type ZEXCEL_CELL_ROW optional
      !IV_COLUMN type ZEXCEL_CELL_COLUMN optional
    returning
      value(RV_VALUE) type ZEXCEL_CELL_VALUE .
endinterface.
