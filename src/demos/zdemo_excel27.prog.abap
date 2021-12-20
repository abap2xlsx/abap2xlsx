*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL27
*& Test Styles for ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel27.

CLASS lcl_app DEFINITION.
  PUBLIC SECTION.
    METHODS main
      RAISING
        zcx_excel.
  PRIVATE SECTION.
    METHODS conditional_formatting_cellis
      IMPORTING
        column TYPE simple
        row    TYPE zexcel_cell_row
        rule   TYPE zexcel_condition_rule
        op     TYPE zexcel_condition_operator
        f      TYPE zexcel_style_formula
        f2     TYPE zexcel_style_formula
        numfmt TYPE string
      RAISING
        zcx_excel.
    METHODS conditional_formatting_textfun
      IMPORTING
        column TYPE simple
        row    TYPE zexcel_cell_row
        txtfun TYPE zcl_excel_style_cond=>tv_textfunction
        text   TYPE string
        numfmt TYPE string
      RAISING
        zcx_excel.
ENDCLASS.

CONSTANTS: c_fish       TYPE string VALUE 'Fish'.

DATA: lo_app             TYPE REF TO lcl_app.
DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_worksheet       TYPE REF TO zcl_excel_worksheet,
      lo_range           TYPE REF TO zcl_excel_range,
      lo_data_validation TYPE REF TO zcl_excel_data_validation,
      lo_style_1         TYPE REF TO zcl_excel_style,
      lo_style_2         TYPE REF TO zcl_excel_style,
      lv_style_1_guid    TYPE zexcel_cell_style,
      lv_style_2_guid    TYPE zexcel_cell_style,
      lv_style_guid      TYPE zexcel_cell_style,
      ls_cellis          TYPE zexcel_conditional_cellis,
      ls_textfunction    TYPE zcl_excel_style_cond=>ts_conditional_textfunction.


DATA: lv_title          TYPE zexcel_sheet_title.

CONSTANTS: gc_save_file_name TYPE string VALUE '27_ConditionalFormatting.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.
  DATA: lo_error   TYPE REF TO zcx_excel,
        lv_message TYPE string.

  CREATE OBJECT lo_app.
  TRY.
      lo_app->main( ).
    CATCH zcx_excel INTO lo_error.
      lv_message = lo_error->get_text( ).
      MESSAGE lv_message TYPE 'I' DISPLAY LIKE 'E'.
  ENDTRY.


CLASS lcl_app IMPLEMENTATION.

  METHOD main.

    DATA:
      lo_style_cond TYPE REF TO zcl_excel_style_cond,
      lo_style      TYPE REF TO zcl_excel_style.

    " Creates active sheet
    CREATE OBJECT lo_excel.

    lo_style_1                        = lo_excel->add_new_style( ).
    lo_style_1->fill->filltype        = zcl_excel_style_fill=>c_fill_solid.
    lo_style_1->fill->bgcolor-rgb     = zcl_excel_style_color=>c_green.
    lv_style_1_guid                   = lo_style_1->get_guid( ).

    lo_style_2                        = lo_excel->add_new_style( ).
    lo_style_2->fill->filltype        = zcl_excel_style_fill=>c_fill_solid.
    lo_style_2->fill->bgcolor-rgb     = zcl_excel_style_color=>c_red.
    lv_style_2_guid                   = lo_style_2->get_guid( ).

    " Get active sheet
    lo_worksheet        = lo_excel->get_active_worksheet( ).
    lv_title = 'Conditional formatting'.
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

    " 1st validation
    lo_data_validation              = lo_worksheet->add_new_data_validation( ).
    lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
    lo_data_validation->formula1    = c_fish.
    lo_data_validation->cell_row    = 2.
    lo_data_validation->cell_column = 'C'.
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'C' ip_value = 'Select a value' ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule         = zcl_excel_style_cond=>c_rule_cellis.
    ls_cellis-formula           = '"Anchovy"'.
    ls_cellis-operator          = zcl_excel_style_cond=>c_operator_equal.
    ls_cellis-cell_style        = lv_style_1_guid.
    lo_style_cond->mode_cellis  = ls_cellis.
    lo_style_cond->priority     = 1.
    lo_style_cond->set_range( ip_start_column  = 'C'
                              ip_start_row     = 2
                              ip_stop_column   = 'C'
                              ip_stop_row      = 2 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule         = zcl_excel_style_cond=>c_rule_cellis.
    ls_cellis-formula           = '"Carp"'.
    ls_cellis-operator          = zcl_excel_style_cond=>c_operator_equal.
    ls_cellis-cell_style        = lv_style_2_guid.
    lo_style_cond->mode_cellis  = ls_cellis.
    lo_style_cond->priority     = 2.
    lo_style_cond->set_range( ip_start_column  = 'C'
                              ip_start_row     = 2
                              ip_stop_column   = 'C'
                              ip_stop_row      = 2 ).

    " Conditional formatting for all operators
    conditional_formatting_cellis( column = 'C' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_equal              f = '="Anchovy"' f2 = ''      numfmt = 'equal to Anchovy' ).
    conditional_formatting_cellis( column = 'C' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_equal              f = '="Anchovy"' f2 = ''      numfmt = 'equal to Anchovy' ).
    conditional_formatting_cellis( column = 'D' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_notequal           f = '="Anchovy"' f2 = ''      numfmt = 'not equal to Anchovy' ).
    conditional_formatting_cellis( column = 'E' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_between            f = '="B"'       f2 = '="CC"' numfmt = 'between B and CC' ).
    conditional_formatting_cellis( column = 'F' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_greaterthan        f = '="Catfish"' f2 = ''      numfmt = 'greater than Catfish' ).
    conditional_formatting_cellis( column = 'G' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_greaterthanorequal f = '="Catfish"' f2 = ''      numfmt = 'greater than or equal to Catfish' ).
    conditional_formatting_cellis( column = 'H' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_lessthan           f = '="Catfish"' f2 = ''      numfmt = 'less than Catfish' ).
    conditional_formatting_cellis( column = 'I' row = 4 rule = zcl_excel_style_cond=>c_rule_cellis op = zcl_excel_style_cond=>c_operator_lessthanorequal    f = '="Catfish"' f2 = ''      numfmt = 'less than or equal to Catfish' ).

    " Conditional formatting for all text functions
    conditional_formatting_textfun( column = 'C' row = 6 txtfun = zcl_excel_style_cond=>c_textfunction_beginswith   text = 'A' numfmt = 'begins with A' ).
    conditional_formatting_textfun( column = 'D' row = 6 txtfun = zcl_excel_style_cond=>c_textfunction_containstext text = 'h' numfmt = 'contains text h' ).
    conditional_formatting_textfun( column = 'E' row = 6 txtfun = zcl_excel_style_cond=>c_textfunction_endswith     text = 'p' numfmt = 'ends with p' ).
    conditional_formatting_textfun( column = 'F' row = 6 txtfun = zcl_excel_style_cond=>c_textfunction_notcontains  text = 'h' numfmt = 'not contains h' ).

*** Create output
    lcl_output=>output( lo_excel ).

  ENDMETHOD.


  METHOD conditional_formatting_cellis.

    DATA:
      lo_style      TYPE REF TO zcl_excel_style,
      lo_style_cond TYPE REF TO zcl_excel_style_cond.

    lo_style                             = lo_excel->add_new_style( ).
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_white.
    lo_style->number_format->format_code = '@\ "' && numfmt && '"'.
    lo_style->alignment->wraptext        = abap_true.
    lv_style_guid                        = lo_style->get_guid( ).

    lo_worksheet->set_cell( ip_row = row ip_column = column ip_formula = '$C$2' ip_style = lv_style_guid ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule         = rule.
    ls_cellis-operator          = op.
    ls_cellis-formula           = f.
    ls_cellis-formula2          = f2.
    ls_cellis-cell_style        = lv_style_1_guid.
    lo_style_cond->mode_cellis  = ls_cellis.
    lo_style_cond->priority     = 1.
    lo_style_cond->set_range( ip_start_column  = column
                              ip_start_row     = row
                              ip_stop_column   = column
                              ip_stop_row      = row ).

  ENDMETHOD.


  METHOD conditional_formatting_textfun.

    DATA:
      lo_style      TYPE REF TO zcl_excel_style,
      lo_style_cond TYPE REF TO zcl_excel_style_cond.

    lo_style                             = lo_excel->add_new_style( ).
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_white.
    lo_style->number_format->format_code = '@\ "' && numfmt && '"'.
    lo_style->alignment->wraptext        = abap_true.
    lv_style_guid                        = lo_style->get_guid( ).

    lo_worksheet->set_cell( ip_row = row ip_column = column ip_formula = '$C$2' ip_style = lv_style_guid ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule              = zcl_excel_style_cond=>c_rule_textfunction.
    ls_textfunction-textfunction     = txtfun.
    ls_textfunction-text             = text.
    ls_textfunction-cell_style       = lv_style_1_guid.
    lo_style_cond->mode_textfunction = ls_textfunction.
    lo_style_cond->priority          = 1.
    lo_style_cond->set_range( ip_start_column  = column
                              ip_start_row     = row
                              ip_stop_column   = column
                              ip_stop_row      = row ).

  ENDMETHOD.


ENDCLASS.
