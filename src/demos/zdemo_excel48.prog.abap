*&---------------------------------------------------------------------*
*& Report zdemo_excel48
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel48.

DATA:
  lo_excel             TYPE REF TO zcl_excel,
  lo_worksheet         TYPE REF TO zcl_excel_worksheet,
  lo_range             TYPE REF TO zcl_excel_range,
  lv_validation_string TYPE string,
  lo_data_validation   TYPE REF TO zcl_excel_data_validation,
  lv_row               TYPE zexcel_cell_row.


CONSTANTS:
  gc_save_file_name TYPE string VALUE '46_ValidationWarning.xlsx'.

INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

*** Sheet Validation

* Creates active sheet
  CREATE OBJECT lo_excel.

* Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).

    DATA: lo_style_1        TYPE REF TO zcl_excel_style,
          lo_style_2        TYPE REF TO zcl_excel_style,
          lv_style_1_guid TYPE zexcel_cell_style,
          lv_style_2_guid TYPE zexcel_cell_style.

    lo_style_1 = lo_excel->add_new_style( ).
lo_style_1->font->color-rgb = 'FF000000'.
    lo_style_2 = lo_excel->add_new_style( ).
lo_style_2->font->color-rgb = 'FFFF0000'.

 DATA(lt_rtf) = VALUE zexcel_t_rtf(
" no need to specify font style for the first part if it is to be the same as the cell style
    ( offset = 0
      length = 2
      font   = lo_style_1->font->get_structure( ) )
" and now, ladies and gents, a DIFFERENT!!! font style
    ( offset = 2
      length = 2
      font   = lo_style_2->font->get_structure( ) )
" but after that you must explicitly specify back the font style of the cell style,
" otherwise in Excel this will be rendered in the font style of the default cell style
" of the worksheet
    ( offset = 4
      length = 2
      font   = lo_style_1->font->get_structure( ) )
   ).

" rich text formatting info must exactly cover the entire string value, or an exception is thrown
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2
*    ip_style = lo_style->get_guid( )
    ip_value = 'ABCDEF'
    it_rtf   = lt_rtf  ).
* add some fields with validation
  lv_row = 2.
  WHILE lv_row <= 4.
    lo_worksheet->set_cell( ip_row = lv_row ip_column = 'A' ip_value = 'Select' ).
    lv_row = lv_row + 1.
  ENDWHILE.

*** Create output
  lcl_output=>output( lo_excel ).
