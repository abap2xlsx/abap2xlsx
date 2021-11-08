INTERFACE zif_excel_writer
  PUBLIC .


  METHODS write_file
    IMPORTING
      !io_excel      TYPE REF TO zcl_excel
    RETURNING
      VALUE(ep_file) TYPE xstring
    RAISING
      zcx_excel.
ENDINTERFACE.
