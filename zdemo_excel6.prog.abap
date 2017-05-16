*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL6
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel6.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lv_row                  TYPE syindex,
      lv_formula              TYPE string.


CONSTANTS: gc_save_file_name TYPE string VALUE '06_Formulas.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).

*--------------------------------------------------------------------*
*  Get some testdata
*--------------------------------------------------------------------*
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 100  ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 1000  ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 150 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = -10  ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 500  ).


*--------------------------------------------------------------------*
*  Demonstrate using formulas
*--------------------------------------------------------------------*
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'C' ip_formula = 'SUM(C4:C8)' ).


*--------------------------------------------------------------------*
* Demonstrate standard EXCEL-behaviour when copying a formula to another cell
* by calculating the resulting formula to put into another cell
*--------------------------------------------------------------------*
  DO 10 TIMES.

    lv_formula = zcl_excel_common=>shift_formula( iv_reference_formula = 'SUM(C4:C8)'
                                                  iv_shift_cols        = 0                " Offset in Columns - here we copy in same column --> 0
                                                  iv_shift_rows        = sy-index ).      " Offset in Row     - here we copy downward --> sy-index
    lv_row = 9 + sy-index.                                                                " Absolute row = sy-index rows below reference cell
    lo_worksheet->set_cell( ip_row = lv_row ip_column = 'C' ip_formula = lv_formula ).

  ENDDO.

*--------------------------------------------------------------------*
*** Create output
*--------------------------------------------------------------------*
  lcl_output=>output( lo_excel ).
