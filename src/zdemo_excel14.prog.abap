*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL14
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel14.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_style_center         TYPE REF TO zcl_excel_style,
      lo_style_right          TYPE REF TO zcl_excel_style,
      lo_style_left           TYPE REF TO zcl_excel_style,
      lo_style_general        TYPE REF TO zcl_excel_style,
      lo_style_bottom         TYPE REF TO zcl_excel_style,
      lo_style_middle         TYPE REF TO zcl_excel_style,
      lo_style_top            TYPE REF TO zcl_excel_style,
      lo_style_justify        TYPE REF TO zcl_excel_style,
      lo_style_mixed          TYPE REF TO zcl_excel_style,
      lo_style_mixed_wrap     TYPE REF TO zcl_excel_style,
      lo_style_rotated        TYPE REF TO zcl_excel_style,
      lo_style_shrink         TYPE REF TO zcl_excel_style,
      lo_style_indent         TYPE REF TO zcl_excel_style,
      lv_style_center_guid    TYPE zexcel_cell_style,
      lv_style_right_guid     TYPE zexcel_cell_style,
      lv_style_left_guid      TYPE zexcel_cell_style,
      lv_style_general_guid   TYPE zexcel_cell_style,
      lv_style_bottom_guid    TYPE zexcel_cell_style,
      lv_style_middle_guid    TYPE zexcel_cell_style,
      lv_style_top_guid       TYPE zexcel_cell_style,
      lv_style_justify_guid   TYPE zexcel_cell_style,
      lv_style_mixed_guid     TYPE zexcel_cell_style,
      lv_style_mixed_wrap_guid TYPE zexcel_cell_style,
      lv_style_rotated_guid   TYPE zexcel_cell_style,
      lv_style_shrink_guid    TYPE zexcel_cell_style,
      lv_style_indent_guid    TYPE zexcel_cell_style.

DATA: lo_row        TYPE REF TO zcl_excel_row.

CONSTANTS: gc_save_file_name TYPE string VALUE '14_Alignment.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'sheet1' ).

  "Center
  lo_style_center = lo_excel->add_new_style( ).
  lo_style_center->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lv_style_center_guid = lo_style_center->get_guid( ).
  "Right
  lo_style_right = lo_excel->add_new_style( ).
  lo_style_right->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_right.
  lv_style_right_guid = lo_style_right->get_guid( ).
  "Left
  lo_style_left = lo_excel->add_new_style( ).
  lo_style_left->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_left.
  lv_style_left_guid = lo_style_left->get_guid( ).
  "General
  lo_style_general = lo_excel->add_new_style( ).
  lo_style_general->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_general.
  lv_style_general_guid = lo_style_general->get_guid( ).
  "Bottom
  lo_style_bottom = lo_excel->add_new_style( ).
  lo_style_bottom->alignment->vertical = zcl_excel_style_alignment=>c_vertical_bottom.
  lv_style_bottom_guid = lo_style_bottom->get_guid( ).
  "Middle
  lo_style_middle = lo_excel->add_new_style( ).
  lo_style_middle->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
  lv_style_middle_guid = lo_style_middle->get_guid( ).
  "Top
  lo_style_top = lo_excel->add_new_style( ).
  lo_style_top->alignment->vertical = zcl_excel_style_alignment=>c_vertical_top.
  lv_style_top_guid = lo_style_top->get_guid( ).
  "Justify
  lo_style_justify = lo_excel->add_new_style( ).
  lo_style_justify->alignment->vertical = zcl_excel_style_alignment=>c_vertical_justify.
  lv_style_justify_guid = lo_style_justify->get_guid( ).

  "Shrink
  lo_style_shrink = lo_excel->add_new_style( ).
  lo_style_shrink->alignment->shrinktofit = abap_true.
  lv_style_shrink_guid = lo_style_shrink->get_guid( ).

  "Indent
  lo_style_indent = lo_excel->add_new_style( ).
  lo_style_indent->alignment->indent = 5.
  lv_style_indent_guid = lo_style_indent->get_guid( ).

  "Middle / Centered / Wrap
  lo_style_mixed_wrap = lo_excel->add_new_style( ).
  lo_style_mixed_wrap->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_mixed_wrap->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
  lo_style_mixed_wrap->alignment->wraptext     = abap_true.
  lv_style_mixed_wrap_guid = lo_style_mixed_wrap->get_guid( ).

  "Middle / Centered / Wrap
  lo_style_mixed = lo_excel->add_new_style( ).
  lo_style_mixed->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_mixed->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
  lv_style_mixed_guid = lo_style_mixed->get_guid( ).

  "Center
  lo_style_rotated = lo_excel->add_new_style( ).
  lo_style_rotated->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_rotated->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
  lo_style_rotated->alignment->textrotation = 165.                        " -75Ã‚Â° == 90Ã‚Â° + 75Ã‚Â°
  lv_style_rotated_guid = lo_style_rotated->get_guid( ).


  " Set row size for first 7 rows to 40
  DO 7 TIMES.
    lo_row = lo_worksheet->get_row( sy-index ).
    lo_row->set_row_height( 40 ).
  ENDDO.

  "Horizontal alignment
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'Centered Text' ip_style = lv_style_center_guid ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value = 'Right Text'    ip_style = lv_style_right_guid ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = 'Left Text'     ip_style = lv_style_left_guid ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'B' ip_value = 'General Text'  ip_style = lv_style_general_guid ).

  " Shrink & indent
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'F' ip_value = 'Text shrinked' ip_style = lv_style_shrink_guid ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'F' ip_value = 'Text indented' ip_style = lv_style_indent_guid ).

  "Vertical alignment

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = 'Bottom Text'    ip_style = lv_style_bottom_guid ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'D' ip_value = 'Middle Text'    ip_style = lv_style_middle_guid ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'D' ip_value = 'Top Text'       ip_style = lv_style_top_guid ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'D' ip_value = 'Justify Text'   ip_style = lv_style_justify_guid ).

  " Wrapped
  lo_worksheet->set_cell( ip_row = 10 ip_column = 'B'
                          ip_value = 'This is a wrapped text centered in the middle'
                          ip_style = lv_style_mixed_wrap_guid ).

  " Rotated
  lo_worksheet->set_cell( ip_row = 10 ip_column = 'D'
                          ip_value = 'This is a centered text rotated by -75Ã‚Â°'
                          ip_style = lv_style_rotated_guid ).

  " forced line break
  DATA: lv_value TYPE string.
  CONCATENATE 'This is a wrapped text centered in the middle' cl_abap_char_utilities=>cr_lf
    'and a manuall line break.' INTO lv_value.
  lo_worksheet->set_cell( ip_row = 11 ip_column = 'B'
                          ip_value = lv_value
                          ip_style = lv_style_mixed_guid ).

*** Create output
  lcl_output=>output( lo_excel ).
