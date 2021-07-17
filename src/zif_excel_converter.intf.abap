INTERFACE zif_excel_converter
  PUBLIC .


  METHODS can_convert_object
    IMPORTING
      !io_object TYPE REF TO object
    RAISING
      zcx_excel .
  METHODS create_fieldcatalog
    IMPORTING
      !is_option       TYPE zexcel_s_converter_option
      !io_object       TYPE REF TO object
      !it_table        TYPE STANDARD TABLE
    EXPORTING
      !es_layout       TYPE zexcel_s_converter_layo
      !et_fieldcatalog TYPE zexcel_t_converter_fcat
      !eo_table        TYPE REF TO data
      !et_colors       TYPE zexcel_t_converter_col
      !et_filter       TYPE zexcel_t_converter_fil
    RAISING
      zcx_excel .

   METHODS get_supported_class
     RETURNING VALUE(rv_supported_class) TYPE seoclsname.
ENDINTERFACE.
