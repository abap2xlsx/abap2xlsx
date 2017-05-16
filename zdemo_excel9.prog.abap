*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL9
*&
*&---------------------------------------------------------------------*
*& abap2xlsx Demo: Data validations
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel9.

CONSTANTS: c_fruits     TYPE string VALUE 'Fruits',
           c_vegetables TYPE string VALUE 'Vegetables',
           c_meat       TYPE string VALUE 'Meat',
           c_fish       TYPE string VALUE 'Fish'.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_range                TYPE REF TO zcl_excel_range,
      lo_data_validation      TYPE REF TO zcl_excel_data_validation.

DATA: row TYPE zexcel_cell_row.


DATA: lv_title          TYPE zexcel_sheet_title.


CONSTANTS: gc_save_file_name TYPE string VALUE '09_DataValidation.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_sbook TYPE flag.


START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet        = lo_excel->get_active_worksheet( ).
  lv_title = 'Data Validation'.
  lo_worksheet->set_title( lv_title ).
  " Set values for dropdown
  lo_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = c_fish ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 'Anchovy' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'A' ip_value = 'Carp' ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 'Catfish' ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'A' ip_value = 'Cod' ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'A' ip_value = 'Eel' ).
  lo_worksheet->set_cell( ip_row = 9 ip_column = 'A' ip_value = 'Haddock' ).

  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = c_fish.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'A'
                       ip_start_row     = 4
                       ip_stop_column   = 'A'
                       ip_stop_row      = 9 ).

  lo_worksheet->set_cell( ip_row = 2 ip_column = 'B' ip_value = c_meat ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'Pork' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value = 'Beef' ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = 'Chicken' ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'B' ip_value = 'Turkey' ).

  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = c_meat.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'B'
                       ip_start_row     = 4
                       ip_stop_column   = 'B'
                       ip_stop_row      = 7 ).

  lo_worksheet->set_cell( ip_row = 2 ip_column = 'C' ip_value = c_fruits ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'Apple' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 'Banana' ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 'Blueberry' ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 'Ananas' ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 'Grapes' ).

  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = c_fruits.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'C'
                       ip_start_row     = 4
                       ip_stop_column   = 'C'
                       ip_stop_row      = 8 ).

  lo_worksheet->set_cell( ip_row = 2 ip_column = 'D' ip_value = c_vegetables ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = 'Cucumber' ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'D' ip_value = 'Sweet pepper ' ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'D' ip_value = 'Lettuce' ).

  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = c_vegetables.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'D'
                       ip_start_row     = 4
                       ip_stop_column   = 'D'
                       ip_stop_row      = 6 ).

  lo_worksheet        = lo_excel->add_new_worksheet( ).
  lv_title = 'Table with Data Validation'.
  lo_worksheet->set_title( lv_title ).

  " Maximum Text length
  lo_worksheet->set_cell(  ip_row = 1 ip_column = 'A' ip_value = 'Validate Maximum Text length of <= 10 in Cell A2:' ).
  lo_worksheet->set_cell(  ip_row = 2 ip_column = 'A' ip_value = 'abcdefghij' ).
  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_textlength.
  lo_data_validation->operator    = zcl_excel_data_validation=>c_operator_lessthanorequal.
  lo_data_validation->formula1    = 10.
  lo_data_validation->cell_row    = 2.
  lo_data_validation->cell_column = 'A'.

  " Integer Value between 1 and 10
  lo_worksheet->set_cell(  ip_row = 4 ip_column = 'A' ip_value = 'Validate Integer Value between 1 and 10 in Cell A5:' ).
  lo_worksheet->set_cell(  ip_row = 5 ip_column = 'A' ip_value = '5' ).
  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_whole.
  lo_data_validation->operator    = zcl_excel_data_validation=>c_operator_between.
  lo_data_validation->formula1    = 1.
  lo_data_validation->formula2    = 10.
  lo_data_validation->prompttitle = 'Range'.
  lo_data_validation->prompt      = 'Enter a value between 1 and 10'.
  lo_data_validation->errortitle  = 'Error'.
  lo_data_validation->error       = 'You have entered a wrong value. Please use only numbers between 1 and 10.'.
  lo_data_validation->cell_row    = 5.
  lo_data_validation->cell_column = 'A'.

  " Evaluation by Formula from issue #161
  lo_worksheet->set_cell(  ip_row = 7 ip_column = 'A' ip_value = 'Validate if B8 contains a "-":' ).
  lo_worksheet->set_cell(  ip_row = 8 ip_column = 'A' ip_value = 'Text' ).
  lo_worksheet->set_cell(  ip_row = 8 ip_column = 'B' ip_value = '-' ).
  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_custom.
  lo_data_validation->formula1    = '"IF(B8<>"""";INDIRECT(LEFT(B8;SEARCH(""-"";B8;1)));EMPTY)"'.
  lo_data_validation->cell_row    = 8.
  lo_data_validation->cell_column = 'A'.

  " There was an error when data validation was combined with cell merges this should test that:
  lo_worksheet->set_cell(  ip_row = 10 ip_column = 'A' ip_value = 'Demo for data validation with a dropdown list' ).
  lo_worksheet->set_merge( ip_row = 10 ip_column_start = 'A' ip_column_end = 'F' ).

  " Headlines
  lo_worksheet->set_cell( ip_row = 11 ip_column = 'A' ip_value = c_fruits ).
  lo_worksheet->set_cell( ip_row = 11 ip_column = 'B' ip_value = c_vegetables ).

  row = 12.
  WHILE row < 20. " Starting with 14500 the data validation is dropped 14000 are still ok
    " 1st validation
    lo_data_validation              = lo_worksheet->add_new_data_validation( ).
    lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
    lo_data_validation->formula1    = c_fruits.
    lo_data_validation->cell_row    = row.
    lo_data_validation->cell_column = 'A'.
    lo_worksheet->set_cell( ip_row = row ip_column = 'A' ip_value = 'Select a value' ).
                                                            " 2nd
    lo_data_validation              = lo_worksheet->add_new_data_validation( ).
    lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
    lo_data_validation->formula1    = c_vegetables.
    lo_data_validation->cell_row    = row.
    lo_data_validation->cell_column = 'B'.
    lo_worksheet->set_cell( ip_row = row ip_column = 'B' ip_value = 'Select a value' ).
                                                            " 3rd
    lo_data_validation              = lo_worksheet->add_new_data_validation( ).
    lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
    lo_data_validation->formula1    = c_meat.
    lo_data_validation->cell_row    = row.
    lo_data_validation->cell_column = 'C'.
    lo_worksheet->set_cell( ip_row = row ip_column = 'C' ip_value = 'Select a value' ).
                                                            " 4th
    lo_data_validation              = lo_worksheet->add_new_data_validation( ).
    lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
    lo_data_validation->formula1    = c_fish.
    lo_data_validation->cell_row    = row.
    lo_data_validation->cell_column = 'D'.
    lo_worksheet->set_cell( ip_row = row ip_column = 'D' ip_value = 'Select a value' ).
    " Increment row
    row = row + 1.
  ENDWHILE.

  IF p_sbook = abap_true.
    DATA: bookings type TABLE OF sbook.

    lo_worksheet        = lo_excel->add_new_worksheet( ).
    lv_title = 'SBOOK'.
    lo_worksheet->set_title( lv_title ).

    SELECT * from sbook INTO TABLE bookings UP TO 4000 ROWS.

    lo_worksheet->bind_table(
      EXPORTING
        ip_table          = bookings
*        it_field_catalog  =     " Table binding field catalog
*        is_table_settings =     " Excel table binding settings
*      IMPORTING
*        es_table_settings =     " Excel table binding settings
    ).
  ENDIF.


*** Create output
  lcl_output=>output( lo_excel ).
