*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL2
*& Test Styles for ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel34.

CONSTANTS: width            TYPE f VALUE '10.14'.
CONSTANTS: height           TYPE f VALUE '57.75'.

DATA:      current_row      TYPE i,
           col              TYPE i,
           col_alpha        TYPE zexcel_cell_column_alpha,
           row              TYPE i,
           row_board        TYPE i,
           colorflag        TYPE i,
           color            TYPE zexcel_style_color_argb,

           lo_column        TYPE REF TO zcl_excel_column,
           lo_row           TYPE REF TO zcl_excel_row,

           writing1         TYPE string,
           writing2         TYPE string.



DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet.

CONSTANTS: gc_save_file_name TYPE string VALUE '34_Static Styles_Chess.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.
  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Spassky_vs_Bronstein' ).

* Header
  current_row = 1.

  ADD 1 TO current_row.
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'B' ip_value = 'White' ).
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'C' ip_value = 'Spassky, Boris V   --  wins in turn 23' ).

  ADD 1 TO current_row.
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'B' ip_value = 'Black' ).
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'C' ip_value = 'Bronstein, David I' ).

  ADD 1 TO current_row.
* Set size of column + Writing above chessboard
  DO 8 TIMES.

    writing1 = zcl_excel_common=>convert_column2alpha( sy-index ).
    writing2 =  sy-index .
    row = current_row + sy-index.

    col = sy-index + 1.
    col_alpha = zcl_excel_common=>convert_column2alpha( col ).

* Set size of column
    lo_column = lo_worksheet->get_column( col_alpha ).
    lo_column->set_width( width ).

* Set size of row
    lo_row = lo_worksheet->get_row( row ).
    lo_row->set_row_height( height ).

* Set writing on chessboard
    lo_worksheet->set_cell( ip_row = row
                            ip_column = 'A'
                            ip_value = writing2 ).
    lo_worksheet->change_cell_style(  ip_column               = 'A'
                                      ip_row                  = row
                                      ip_alignment_vertical   = zcl_excel_style_alignment=>c_vertical_center  ).
    lo_worksheet->set_cell( ip_row = row
                            ip_column = 'J'
                            ip_value = writing2 ).
    lo_worksheet->change_cell_style(  ip_column               = 'J'
                                      ip_row                  = row
                                      ip_alignment_vertical   = zcl_excel_style_alignment=>c_vertical_center  ).

    row = current_row + 9.
    lo_worksheet->set_cell( ip_row = current_row
                            ip_column = col_alpha
                            ip_value = writing1 ).
    lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                      ip_row                  = current_row
                                      ip_alignment_horizontal = zcl_excel_style_alignment=>c_horizontal_center ).
    lo_worksheet->set_cell( ip_row = row
                            ip_column = col_alpha
                            ip_value = writing1 ).
    lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                      ip_row                  = row
                                      ip_alignment_horizontal = zcl_excel_style_alignment=>c_horizontal_center ).
  ENDDO.
  lo_column = lo_worksheet->get_column( 'A' ).
  lo_column->set_auto_size( abap_true ).
  lo_column = lo_worksheet->get_column( 'J' ).
  lo_column->set_auto_size( abap_true ).

* Set win-position
  CONSTANTS: c_pawn   TYPE string VALUE 'Pawn'.
  CONSTANTS: c_rook   TYPE string VALUE 'Rook'.
  CONSTANTS: c_knight TYPE string VALUE 'Knight'.
  CONSTANTS: c_bishop TYPE string VALUE 'Bishop'.
  CONSTANTS: c_queen  TYPE string VALUE 'Queen'.
  CONSTANTS: c_king   TYPE string VALUE 'King'.

  row = current_row + 1.
  lo_worksheet->set_cell( ip_row = row ip_column = 'B' ip_value = c_rook ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_rook ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'G' ip_value = c_knight ).
  row = current_row + 2.
  lo_worksheet->set_cell( ip_row = row ip_column = 'B' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'C' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'D' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_queen ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'H' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'I' ip_value = c_king ).
  row = current_row + 3.
  lo_worksheet->set_cell( ip_row = row ip_column = 'I' ip_value = c_pawn ).
  row = current_row + 4.
  lo_worksheet->set_cell( ip_row = row ip_column = 'D' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_knight ).
  row = current_row + 5.
  lo_worksheet->set_cell( ip_row = row ip_column = 'E' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_queen ).
  row = current_row + 6.
  lo_worksheet->set_cell( ip_row = row ip_column = 'C' ip_value = c_bishop ).
  row = current_row + 7.
  lo_worksheet->set_cell( ip_row = row ip_column = 'B' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'C' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'H' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'I' ip_value = c_pawn ).
  row = current_row + 8.
  lo_worksheet->set_cell( ip_row = row ip_column = 'G' ip_value = c_rook ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'H' ip_value = c_king ).

* Set Chessboard
  DO 8 TIMES.
    IF sy-index <= 3.  " Black
      color = zcl_excel_style_color=>c_black.
    ELSE.
      color = zcl_excel_style_color=>c_white.
    ENDIF.
    row_board = sy-index.
    row = current_row + sy-index.
    DO 8 TIMES.
      col = sy-index + 1.
      col_alpha = zcl_excel_common=>convert_column2alpha( col ).
      TRY.
* Borders around outer limits
          IF row_board = 1.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_top_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_top_color_rgb =  zcl_excel_style_color=>c_black ).
          ENDIF.
          IF row_board = 8.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_down_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_down_color_rgb =  zcl_excel_style_color=>c_black ).
          ENDIF.
          IF col = 2.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_left_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_left_color_rgb =  zcl_excel_style_color=>c_black ).
          ENDIF.
          IF col = 9.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_right_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_right_color_rgb =  zcl_excel_style_color=>c_black ).
          ENDIF.
* Style for writing
          lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                            ip_row                  = row
                                            ip_font_color_rgb       = color
                                            ip_font_bold            = 'X'
                                            ip_font_size            = 16
                                            ip_alignment_horizontal = zcl_excel_style_alignment=>c_horizontal_center
                                            ip_alignment_vertical   = zcl_excel_style_alignment=>c_vertical_center
                                            ip_fill_filltype        = zcl_excel_style_fill=>c_fill_solid ).
* Color of field
          colorflag = ( row + col ) MOD 2.
          IF colorflag = 0.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_fill_fgcolor_rgb     = 'FFB5866A'
                                              ip_fill_filltype        = zcl_excel_style_fill=>c_fill_gradient_diagonal135 ).
          ELSE.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_fill_fgcolor_rgb     = 'FFF5DEBF'
                                              ip_fill_filltype        = zcl_excel_style_fill=>c_fill_gradient_diagonal45 ).
          ENDIF.



        CATCH zcx_excel .
      ENDTRY.

    ENDDO.
  ENDDO.


*** Create output
  lcl_output=>output( lo_excel ).
