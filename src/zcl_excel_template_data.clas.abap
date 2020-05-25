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
    FIELD-SYMBOLS
                   : <fs_data> TYPE ZEXCEL_S_TEMPLATE_DATA
                   , <fs_any> TYPE any
                   .

    APPEND INITIAL LINE TO mt_data ASSIGNING <fs_data>.
    <fs_data>-sheet = IV_SHEET.
    CREATE data  <fs_data>-data LIKE IV_DATA.

    ASSIGN <fs_data>-data->* to <fs_any>.
    <fs_any> = iv_data.

  endmethod.
ENDCLASS.
