interface ZIF_EXCEL_WRITER
  public .


  methods WRITE_FILE
    importing
      !IO_EXCEL type ref to ZCL_EXCEL
    returning
      value(EP_FILE) type XSTRING .
endinterface.
