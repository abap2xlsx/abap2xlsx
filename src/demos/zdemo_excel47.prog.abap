*&---------------------------------------------------------------------*
*& Report zdemo_excel47
*&---------------------------------------------------------------------*
*&
*& - BIND_TABLE and Calculated Column Formulas
*& - Ignore cell errors
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel47.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_hyperlink TYPE REF TO zcl_excel_hyperlink,
      lo_column    TYPE REF TO zcl_excel_column.

CONSTANTS: gc_save_file_name TYPE string VALUE '47_CalculatedColumns.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

CLASS lcl_app DEFINITION.
  PUBLIC SECTION.

    METHODS main
      RAISING
        zcx_excel.

ENDCLASS.

CLASS lcl_app IMPLEMENTATION.

  METHOD main.

    TYPES: BEGIN OF helper_type,
             carrid             TYPE sflight-carrid,
             connid             TYPE sflight-connid,
             fldate             TYPE sflight-fldate,
             price              TYPE sflight-price,
             formula            TYPE string,
             calculated_formula TYPE string,
             shared_formula     TYPE sflight-price,
             shared_formula_2   TYPE sflight-price,
           END OF helper_type.
    DATA: f1                TYPE string,
          f2                TYPE string,
          line              TYPE helper_type,
          itab              TYPE STANDARD TABLE OF helper_type,
          field_catalog     TYPE zexcel_t_fieldcatalog,
          ls_catalog        TYPE zexcel_s_fieldcatalog,
          ls_ignored_errors TYPE zcl_excel_worksheet=>mty_s_ignored_errors,
          lt_ignored_errors TYPE zcl_excel_worksheet=>mty_th_ignored_errors.

    " Load data
    f1 = 'TblFlights[[#This Row],[Airfare]]+100'.
    f2 = 'TblFlights[[#This Row],[Airfare]]+222'.
    line-carrid = `AA`. line-connid = '0017'. line-fldate = '20180116'. line-price = '422.94'. line-formula = f1. line-calculated_formula = f2.
    APPEND line TO itab.
    line-carrid = `AZ`. line-connid = '0555'. line-fldate = '20180116'. line-price = '185.00'. line-formula = f1. line-calculated_formula = f1.
    APPEND line TO itab.
    line-carrid = `LH`. line-connid = '0400'. line-fldate = '20180119'. line-price = '666.00'. line-formula = f1. line-calculated_formula = f2.
    APPEND line TO itab.
    line-carrid = `UA`. line-connid = '0941'. line-fldate = '20180117'. line-price = '879.82'. line-formula = f1. line-calculated_formula = f2.
    APPEND line TO itab.

    " Creates active sheet
    CREATE OBJECT lo_excel.

    " Get active sheet
    lo_worksheet = lo_excel->get_active_worksheet( ).

    field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = itab ).

    field_catalog = VALUE #(
        LET aux_field_catalog = field_catalog IN
        FOR <line> IN aux_field_catalog
        ( SWITCH #( <line>-fieldname
          " No formula
          WHEN 'PRICE'              THEN VALUE #( BASE <line> totals_function = zcl_excel_table=>totals_function_average )
          " Each cell may have a distinct formula, none formula is applied to future new rows
          WHEN 'FORMULA'            THEN VALUE #( BASE <line> scrtext_m       = 'Formula and aggregate function'
                                                              formula         = abap_true
                                                              totals_function = zcl_excel_table=>totals_function_sum )
          " each cell may have a distinct formula, a formula is applied to future new rows
          WHEN 'CALCULATED_FORMULA' THEN VALUE #( BASE <line> scrtext_m       = 'Calculated formula except one and aggregate function'
                                                              formula         = abap_true
                                                              column_formula  = f2
                                                              totals_function = zcl_excel_table=>totals_function_min )
          " The column formula applies to all rows and to future new rows
          WHEN 'SHARED_FORMULA'     THEN VALUE #( BASE <line> scrtext_m       = 'Calculated formula'
                                                              column_formula  = 'D2+100'  )
          " The column formula applies to all rows and to future new rows
          WHEN 'SHARED_FORMULA_2'   THEN VALUE #( BASE <line> scrtext_m       = 'Calculated formula and aggregate function'
                                                              column_formula  = 'D2+100'
                                                              totals_function = zcl_excel_table=>totals_function_max )
          ELSE <line> ) ) ).


    lo_worksheet->bind_table(
      EXPORTING
        ip_table            = itab
        it_field_catalog    = field_catalog
        is_table_settings   = VALUE zexcel_s_table_settings(
                                table_style         = zcl_excel_table=>builtinstyle_medium2
                                table_name          = 'TblFlights'
                                top_left_column     = 'A'
                                top_left_row        = 1
                                show_row_stripes    = abap_true ) ).

    " Give one cell of the calculated column a different value and ignore the error "inconsistent calculated column formula"
    lo_worksheet->set_cell( ip_column = 7 ip_row = 2 ip_value = 'Text' ). " cell 'G2'
    CLEAR ls_ignored_errors.
    ls_ignored_errors-cell_coords = 'G2'.
    ls_ignored_errors-calculated_column = abap_true.
    INSERT ls_ignored_errors INTO TABLE lt_ignored_errors.
    lo_worksheet->set_ignored_errors( lt_ignored_errors ).

    " Numbers stored as texts
    CLEAR ls_ignored_errors.
    ls_ignored_errors-cell_coords = 'B2:B5'.
    ls_ignored_errors-number_stored_as_text = abap_true.
    INSERT ls_ignored_errors INTO TABLE lt_ignored_errors.
    lo_worksheet->set_ignored_errors( lt_ignored_errors ).


*** Create output
    lcl_output=>output( lo_excel ).

  ENDMETHOD.

ENDCLASS.

START-OF-SELECTION.
  TRY.
      NEW lcl_app( )->main( ).
    CATCH zcx_excel INTO DATA(lx).
      MESSAGE lx TYPE 'I' DISPLAY LIKE 'E'.
  ENDTRY.
