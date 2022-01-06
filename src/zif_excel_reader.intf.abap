INTERFACE zif_excel_reader
  PUBLIC .


  METHODS load_file
    IMPORTING
      !i_filename             TYPE csequence
      !i_use_alternate_zip    TYPE seoclsname DEFAULT space
      !i_from_applserver      TYPE abap_bool DEFAULT sy-batch
      !iv_zcl_excel_classname TYPE clike OPTIONAL
    RETURNING
      VALUE(r_excel)          TYPE REF TO zcl_excel
    RAISING
      zcx_excel .
  METHODS load
    IMPORTING
      !i_excel2007            TYPE xstring
      !i_use_alternate_zip    TYPE seoclsname DEFAULT space
      !iv_zcl_excel_classname TYPE clike OPTIONAL
    RETURNING
      VALUE(r_excel)          TYPE REF TO zcl_excel
    RAISING
      zcx_excel .
ENDINTERFACE.
