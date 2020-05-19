class ZCL_EXCEL_TEMPLATE_DATA definition
  public
  final
  create public .

public section.

  data MT_DATA type ZEXCEL_T_TEMPLATE_DATA .

  methods ADD
    importing
      !IV_SHEET type ZEXCEL_SHEET_TITLE
      !IV_DATA type DATA .
protected section.
private section.
ENDCLASS.



CLASS ZCL_EXCEL_TEMPLATE_DATA IMPLEMENTATION.


  method ADD.

    APPEND INITIAL LINE TO mt_data ASSIGNING FIELD-SYMBOL(<fs_data>).
    <fs_data>-sheet = IV_SHEET.
    CREATE data  <fs_data>-data LIKE IV_DATA.

    ASSIGN <fs_data>-data->* to FIELD-SYMBOL(<fs_any>).
    <fs_any> = iv_data.

  endmethod.
ENDCLASS.
