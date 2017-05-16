*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL21
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel21.

TYPES:
  BEGIN OF t_color_style,
    color TYPE zexcel_style_color_argb,
    style TYPE zexcel_cell_style,
  END OF t_color_style.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_style_filled         TYPE REF TO zcl_excel_style.

DATA: color_styles TYPE TABLE OF t_color_style.

FIELD-SYMBOLS: <color_style> LIKE LINE OF color_styles.

CONSTANTS: max  TYPE i VALUE 255,
           step TYPE i VALUE 51.

DATA: red          TYPE i,
      green        TYPE i,
      blue         TYPE i,
      red_hex(1)   TYPE x,
      green_hex(1) TYPE x,
      blue_hex(1)  TYPE x,
      red_str      TYPE string,
      green_str    TYPE string,
      blue_str     TYPE string.

DATA: color TYPE zexcel_style_color_argb,
      tint TYPE zexcel_style_color_tint.

DATA: row     TYPE i,
      row_tmp TYPE i,
      column  TYPE zexcel_cell_column VALUE 1,
      col_str TYPE zexcel_cell_column_alpha.

CONSTANTS: gc_save_file_name TYPE string VALUE '21_BackgroundColorPicker.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  WHILE red <= max.
    green = 0.
    WHILE green <= max.
      blue = 0.
      WHILE blue <= max.
        red_hex = red.
        red_str = red_hex.
        green_hex = green.
        green_str = green_hex.
        blue_hex = blue.
        blue_str = blue_hex.
        " Create filled
        CONCATENATE 'FF' red_str green_str blue_str INTO color.
        APPEND INITIAL LINE TO color_styles ASSIGNING <color_style>.
        lo_style_filled                 = lo_excel->add_new_style( ).
        lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
        lo_style_filled->fill->fgcolor-rgb  = color.
        <color_style>-color = color.
        <color_style>-style = lo_style_filled->get_guid( ).
        blue = blue + step.
      ENDWHILE.
      green = green + step.
    ENDWHILE.
    red = red + step.
  ENDWHILE.
  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Color Picker' ).
  LOOP AT color_styles ASSIGNING <color_style>.
    row_tmp = ( max / step + 1 ) * 3.
    IF row = row_tmp.
      row = 0.
      column = column + 1.
    ENDIF.
    row = row + 1.
    col_str = zcl_excel_common=>convert_column2alpha( column ).

    " Fill the cell and apply one style
    lo_worksheet->set_cell( ip_column = col_str
                            ip_row    = row
                            ip_value  = <color_style>-color
                            ip_style  = <color_style>-style ).
  ENDLOOP.

  row = row + 2.
  tint = '-0.5'.
  DO 10 TIMES.
    column = 1.
    DO 10 TIMES.
      lo_style_filled                 = lo_excel->add_new_style( ).
      lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
      lo_style_filled->fill->fgcolor-theme  = sy-index - 1.
      lo_style_filled->fill->fgcolor-tint  = tint.
      <color_style>-style = lo_style_filled->get_guid( ).
      col_str = zcl_excel_common=>convert_column2alpha( column ).
      lo_worksheet->set_cell_style( ip_column = col_str
                                    ip_row    = row
                                    ip_style  = <color_style>-style ).

      ADD 1 TO column.
    ENDDO.
    ADD '0.1' TO tint.
    ADD 1 TO row.
  ENDDO.



*** Create output
  lcl_output=>output( lo_excel ).
