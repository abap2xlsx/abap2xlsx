*&---------------------------------------------------------------------*
*& Report zdemo_excel47
*&---------------------------------------------------------------------*
*&
*& - BIND_TABLE and Calculated Columns
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel47.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_hyperlink TYPE REF TO zcl_excel_hyperlink,
      lo_column    TYPE REF TO zcl_excel_column.

CONSTANTS: gc_save_file_name TYPE string VALUE '47_ColumnFormulas.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

CLASS lcl_app DEFINITION.
  PUBLIC SECTION.

    METHODS main
      RAISING
        zcx_excel.

ENDCLASS.

CLASS lcl_app IMPLEMENTATION.

  METHOD main.

    TYPES: BEGIN OF ty_tblsheet1_line,
             carrid           TYPE sflight-carrid,
             connid           TYPE sflight-connid,
             fldate           TYPE sflight-fldate,
             price            TYPE sflight-price,
             formula          TYPE string,
             formula_2        TYPE string,
             column_formula   TYPE string,
             column_formula_2 TYPE sflight-price,
             column_formula_3 TYPE sflight-price,
             column_formula_4 TYPE sflight-price,
             column_formula_5 TYPE string,
             column_formula_6 TYPE string,
             column_formula_7 TYPE string,
           END OF ty_tblsheet1_line,
           BEGIN OF ty_tblsheet2_line,
             carrid   TYPE scarr-carrid,
             carrname TYPE scarr-carrname,
           END OF ty_tblsheet2_line.
    DATA: lv_f1             TYPE string,
          ls_tblsheet1      TYPE ty_tblsheet1_line,
          lt_tblsheet1      TYPE STANDARD TABLE OF ty_tblsheet1_line,
          ls_tblsheet2      TYPE ty_tblsheet2_line,
          lt_tblsheet2      TYPE STANDARD TABLE OF ty_tblsheet2_line,
          lt_field_catalog  TYPE zexcel_t_fieldcatalog,
          ls_catalog        TYPE zexcel_s_fieldcatalog,
          ls_table_settings TYPE zexcel_s_table_settings,
          lo_range          TYPE REF TO zcl_excel_range.
    FIELD-SYMBOLS: <ls_field_catalog> TYPE zexcel_s_fieldcatalog.

*** Initialization

    CREATE OBJECT lo_excel.

    " Sheet1
    lv_f1 = 'TblSheet1[[#This Row],[Airfare]]+100'. " [@Airfare]+100
    ls_tblsheet1-carrid = `AA`. ls_tblsheet1-connid = '0017'. ls_tblsheet1-fldate = '20180116'. ls_tblsheet1-price = '422.94'. ls_tblsheet1-formula = lv_f1. ls_tblsheet1-formula_2 = lv_f1.
    APPEND ls_tblsheet1 TO lt_tblsheet1.
    ls_tblsheet1-carrid = `AZ`. ls_tblsheet1-connid = '0555'. ls_tblsheet1-fldate = '20180116'. ls_tblsheet1-price = '185.00'. ls_tblsheet1-formula = lv_f1.
    APPEND ls_tblsheet1 TO lt_tblsheet1.
    ls_tblsheet1-carrid = `LH`. ls_tblsheet1-connid = '0400'. ls_tblsheet1-fldate = '20180119'. ls_tblsheet1-price = '666.00'. ls_tblsheet1-formula = lv_f1. ls_tblsheet1-formula_2 = lv_f1.
    APPEND ls_tblsheet1 TO lt_tblsheet1.
    ls_tblsheet1-carrid = `AA`. ls_tblsheet1-connid = '0941'. ls_tblsheet1-fldate = '20180117'. ls_tblsheet1-price = '879.82'. ls_tblsheet1-formula = lv_f1.
    APPEND ls_tblsheet1 TO lt_tblsheet1.

    " Sheet2
    ls_tblsheet2-carrid = `AA`. ls_tblsheet2-carrname = 'America Airlines'.
    APPEND ls_tblsheet2 TO lt_tblsheet2.
    ls_tblsheet2-carrid = `AZ`. ls_tblsheet2-carrname = 'Alitalia'.
    APPEND ls_tblsheet2 TO lt_tblsheet2.
    ls_tblsheet2-carrid = `LH`. ls_tblsheet2-carrname = 'Lufthansa'.
    APPEND ls_tblsheet2 TO lt_tblsheet2.

*** Sheet1
    lo_worksheet = lo_excel->get_active_worksheet( ).

    lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_tblsheet1 ).

    LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog>.
      CASE <ls_field_catalog>-fieldname.
        WHEN 'CARRID'.
          <ls_field_catalog>-scrtext_l       = 'Company ID'.
        WHEN 'AIRFARE'.
          <ls_field_catalog>-scrtext_l       = 'Airfare'.
        WHEN 'PRICE'.
          <ls_field_catalog>-totals_function = zcl_excel_table=>totals_function_average.
        WHEN 'FORMULA'.
          " Each cell may have a distinct formula, none formula is applied to future new rows
          <ls_field_catalog>-scrtext_l       = 'Formula and aggregate function'.
          <ls_field_catalog>-formula         = abap_true.
          <ls_field_catalog>-totals_function = zcl_excel_table=>totals_function_sum.
        WHEN 'FORMULA_2'.
          " each cell may have a distinct formula, a formula is applied to future new rows
          <ls_field_catalog>-scrtext_l       = 'Formula except 1 cell & aggregate fu.'.
          <ls_field_catalog>-formula         = abap_true.
          <ls_field_catalog>-column_formula  = lv_f1. " to apply to future rows
          <ls_field_catalog>-totals_function = zcl_excel_table=>totals_function_min.
        WHEN 'COLUMN_FORMULA'.
          " The column formula applies to all rows and to future new rows. Internally, the formula is NOT shared because a column name is used.
          <ls_field_catalog>-scrtext_l       = 'Column formula and aggregate function'.
          <ls_field_catalog>-column_formula  = 'TblSheet1[[#This Row],[Airfare]]+222'. " [@Airfare]+222
          <ls_field_catalog>-totals_function = zcl_excel_table=>totals_function_min.
        WHEN 'COLUMN_FORMULA_2'.
          " The column formula applies to all rows and to future new rows. Internally, the formula is shared.
          <ls_field_catalog>-scrtext_l       = 'C2. Column formula'.
          <ls_field_catalog>-column_formula  = 'D2+100'.
        WHEN 'COLUMN_FORMULA_3'.
          " The column formula applies to all rows and to future new rows. Internally, the formula is shared.
          <ls_field_catalog>-scrtext_l       = 'C3. Column formula & aggregate function'.
          <ls_field_catalog>-column_formula  = 'D2+100'.
          <ls_field_catalog>-totals_function = zcl_excel_table=>totals_function_max.
        WHEN 'COLUMN_FORMULA_4'.
          " The column formula applies to all rows and to future new rows. Internally, the formula is shared.
          <ls_field_catalog>-scrtext_l       = 'C4. Column formula array fu./named range'.
          <ls_field_catalog>-column_formula  = 'A1&";"&_xlfn.IFS(TRUE,NamedRange)'. " =A1&";"&@IFS(TRUE,NamedRange)
        WHEN 'COLUMN_FORMULA_5'.
          " The column formula applies to all rows and to future new rows. Internally, the formula is NOT shared because it refers to a different sheet.
          <ls_field_catalog>-scrtext_l       = 'C5. Column formula refers to other sheet'.
          <ls_field_catalog>-column_formula  = 'OtherSheet!A2'.
        WHEN 'COLUMN_FORMULA_6'.
          " The column formula applies to all rows and to future new rows. Internally, the formula is NOT shared.
          " The formula seen in Excel: =FILTER(TblSheet2[Company Name],TblSheet2[Airline ID]=[@Airline],"")
          <ls_field_catalog>-scrtext_l       = 'C6. Column formula array fu./other sheet'.
          <ls_field_catalog>-column_formula  = '_xlfn.FILTER(TblSheet2[Company Name],TblSheet2[Company ID]=TblSheet1[[#This Row],[Company ID]],"")'.
        WHEN 'COLUMN_FORMULA_7'.
          " The column formula applies to all rows and to future new rows. Internally, the formula is NOT shared.
          " The formula seen in Excel: =FILTER(Tbl2_Sheet1[Company Name],Tbl2_Sheet1[Airline ID]=[@Airline],"")
          <ls_field_catalog>-scrtext_l       = 'C7. Column formula array fu./same sheet'.
          <ls_field_catalog>-column_formula  = '_xlfn.FILTER(Tbl2_Sheet1[Company Name],Tbl2_Sheet1[Company ID]=TblSheet1[[#This Row],[Company ID]],"")'.
      ENDCASE.
    ENDLOOP.

    ls_table_settings-table_style         = zcl_excel_table=>builtinstyle_medium2.
    ls_table_settings-table_name          = 'TblSheet1'.
    ls_table_settings-top_left_column     = 'A'.
    ls_table_settings-top_left_row        = 1.
    ls_table_settings-show_row_stripes    = abap_true.

    lo_worksheet->bind_table(
        ip_table          = lt_tblsheet1
        it_field_catalog  = lt_field_catalog
        is_table_settings = ls_table_settings
        iv_default_descr  = 'L' ).

    " Named range for formula 4
    lo_range = lo_excel->add_new_range( ).
    lo_range->name = 'NamedRange'.
    lo_range->set_value( ip_sheet_name    = lo_worksheet->get_title( )
                         ip_start_column  = 'B'
                         ip_start_row     = 1
                         ip_stop_column   = 'B'
                         ip_stop_row      = 1 ).


    " Second table in same sheet
    lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_tblsheet2 ).

    LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog>.
      CASE <ls_field_catalog>-fieldname.
        WHEN 'CARRID'.
          <ls_field_catalog>-scrtext_l = 'Company ID'.
        WHEN 'CARRNAME'.
          <ls_field_catalog>-scrtext_l = 'Company Name'.
      ENDCASE.
    ENDLOOP.

    CLEAR ls_table_settings.
    ls_table_settings-table_style         = zcl_excel_table=>builtinstyle_medium2.
    ls_table_settings-table_name          = 'Tbl2_Sheet1'.
    ls_table_settings-top_left_column     = 'O'.
    ls_table_settings-top_left_row        = 1.
    ls_table_settings-show_row_stripes    = abap_true.

    lo_worksheet->bind_table(
        ip_table          = lt_tblsheet2
        it_field_catalog  = lt_field_catalog
        is_table_settings = ls_table_settings
        iv_default_descr  = 'L' ).

*** Sheet2
    lo_worksheet = lo_excel->add_new_worksheet( 'Sheet2' ).

    CLEAR ls_table_settings.
    ls_table_settings-table_style         = zcl_excel_table=>builtinstyle_medium2.
    ls_table_settings-table_name          = 'TblSheet2'.
    ls_table_settings-top_left_column     = 'A'.
    ls_table_settings-top_left_row        = 1.
    ls_table_settings-show_row_stripes    = abap_true.

    lo_worksheet->bind_table(
        ip_table          = lt_tblsheet2
        it_field_catalog  = lt_field_catalog
        is_table_settings = ls_table_settings
        iv_default_descr  = 'L' ).

*** OtherSheet
    lo_worksheet = lo_excel->add_new_worksheet( 'OtherSheet' ).
    lo_worksheet->set_cell( ip_column = 1 ip_row = 1 ip_value = 'Title' ).
    lo_worksheet->set_cell( ip_column = 1 ip_row = 2 ip_value = 'A2' ).
    lo_worksheet->set_cell( ip_column = 1 ip_row = 3 ip_value = 'A3' ).
    lo_worksheet->set_cell( ip_column = 1 ip_row = 4 ip_value = 'A4' ).
    lo_worksheet->set_cell( ip_column = 1 ip_row = 5 ip_value = 'A5' ).

*** Active sheet = Sheet1
    lo_excel->set_active_sheet_index_by_name( 'Sheet1' ).

*** Create output
    lcl_output=>output( cl_excel = lo_excel iv_info_message = abap_false ).

  ENDMETHOD.

ENDCLASS.

START-OF-SELECTION.
  DATA: go_app   TYPE REF TO lcl_app,
        go_error TYPE REF TO zcx_excel.
  TRY.
      CREATE OBJECT go_app.
      go_app->main( ).
    CATCH zcx_excel INTO go_error.
      MESSAGE go_error TYPE 'I' DISPLAY LIKE 'E'.
  ENDTRY.
