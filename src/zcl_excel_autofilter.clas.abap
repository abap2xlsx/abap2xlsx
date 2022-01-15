CLASS zcl_excel_autofilter DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_AUTOFILTER
*"* do not include other source files here!!!
  PUBLIC SECTION.

    TYPES tv_filter_rule TYPE string .
    TYPES tv_logical_operator TYPE c LENGTH 3 .
    TYPES:
      BEGIN OF ts_filter,
        column           TYPE zexcel_cell_column,
        rule             TYPE tv_filter_rule,
        t_values         TYPE HASHED TABLE OF zexcel_cell_value WITH UNIQUE KEY table_line,
        tr_textfilter1   TYPE RANGE OF string,
        logical_operator TYPE tv_logical_operator,
        tr_textfilter2   TYPE RANGE OF string,
      END OF ts_filter .
    TYPES:
      tt_filters TYPE HASHED TABLE OF ts_filter WITH UNIQUE KEY column .

    DATA filter_area TYPE zexcel_s_autofilter_area .
    CONSTANTS mc_filter_rule_single_values TYPE tv_filter_rule VALUE 'single_values'. "#EC NOTEXT
    CONSTANTS mc_filter_rule_text_pattern TYPE tv_filter_rule VALUE 'text_pattern'. "#EC NOTEXT
    CONSTANTS mc_logical_operator_and TYPE tv_logical_operator VALUE 'and'. "#EC NOTEXT
    CONSTANTS mc_logical_operator_none TYPE tv_logical_operator VALUE space. "#EC NOTEXT
    CONSTANTS mc_logical_operator_or TYPE tv_logical_operator VALUE 'or'. "#EC NOTEXT

    METHODS constructor
      IMPORTING
        !io_sheet TYPE REF TO zcl_excel_worksheet .
    METHODS get_filter_area
      RETURNING
        VALUE(rs_area) TYPE zexcel_s_autofilter_area
      RAISING
        zcx_excel .
    METHODS get_filter_range
      RETURNING
        VALUE(r_range) TYPE zexcel_cell_value
      RAISING
        zcx_excel.
    METHODS get_filter_reference
      RETURNING
        VALUE(r_ref) TYPE zexcel_range_value
      RAISING
        zcx_excel .
    METHODS get_values
      RETURNING
        VALUE(rt_filter) TYPE zexcel_t_autofilter_values .
    METHODS is_row_hidden
      IMPORTING
        !iv_row             TYPE zexcel_cell_row
      RETURNING
        VALUE(rv_is_hidden) TYPE abap_bool .
    METHODS set_filter_area
      IMPORTING
        !is_area TYPE zexcel_s_autofilter_area .
    METHODS set_text_filter
      IMPORTING
        !i_column            TYPE zexcel_cell_column
        !iv_textfilter1      TYPE clike
        !iv_logical_operator TYPE tv_logical_operator DEFAULT mc_logical_operator_none
        !iv_textfilter2      TYPE clike OPTIONAL .
    METHODS set_value
      IMPORTING
        !i_column TYPE zexcel_cell_column
        !i_value  TYPE zexcel_cell_value .
    METHODS set_values
      IMPORTING
        !it_values TYPE zexcel_t_autofilter_values .
*"* protected components of class ZABAP_EXCEL_WORKSHEET
*"* do not include other source files here!!!
  PROTECTED SECTION.

    METHODS get_column_filter
      IMPORTING
        !i_column        TYPE zexcel_cell_column
      RETURNING
        VALUE(rr_filter) TYPE REF TO ts_filter .
    METHODS is_row_hidden_single_values
      IMPORTING
        !iv_row             TYPE zexcel_cell_row
        !iv_col             TYPE zexcel_cell_column
        !is_filter          TYPE ts_filter
      RETURNING
        VALUE(rv_is_hidden) TYPE abap_bool .
    METHODS is_row_hidden_text_pattern
      IMPORTING
        !iv_row             TYPE zexcel_cell_row
        !iv_col             TYPE zexcel_cell_column
        !is_filter          TYPE ts_filter
      RETURNING
        VALUE(rv_is_hidden) TYPE abap_bool .
*"* private components of class ZCL_EXCEL_AUTOFILTER
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA worksheet TYPE REF TO zcl_excel_worksheet .
    DATA mt_filters TYPE tt_filters .

    METHODS validate_area
      RAISING
        zcx_excel .
ENDCLASS.



CLASS zcl_excel_autofilter IMPLEMENTATION.


  METHOD constructor.
    worksheet = io_sheet.
  ENDMETHOD.


  METHOD get_column_filter.

    DATA: ls_filter LIKE LINE OF me->mt_filters.

    READ TABLE me->mt_filters REFERENCE INTO rr_filter WITH TABLE KEY column = i_column.
    IF sy-subrc <> 0.
      ls_filter-column = i_column.
      INSERT ls_filter INTO TABLE me->mt_filters REFERENCE INTO rr_filter.
    ENDIF.

  ENDMETHOD.


  METHOD get_filter_area.

    validate_area( ).

    rs_area = filter_area.

  ENDMETHOD.


  METHOD get_filter_range.
    DATA: l_row_start_c TYPE string,
          l_row_end_c   TYPE string,
          l_col_start_c TYPE string,
          l_col_end_c   TYPE string.

    validate_area( ).

    l_row_end_c = filter_area-row_end.
    CONDENSE l_row_end_c NO-GAPS.

    l_row_start_c = filter_area-row_start.
    CONDENSE l_row_start_c NO-GAPS.

    l_col_start_c = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_start ) .
    l_col_end_c   = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_end ) .

    CONCATENATE l_col_start_c l_row_start_c ':' l_col_end_c l_row_end_c INTO r_range.

  ENDMETHOD.


  METHOD get_filter_reference.
    DATA: l_row_start_c TYPE string,
          l_row_end_c   TYPE string,
          l_col_start_c TYPE string,
          l_col_end_c   TYPE string,
          l_value       TYPE string.

    validate_area( ).

    l_row_end_c = filter_area-row_end.
    CONDENSE l_row_end_c NO-GAPS.

    l_row_start_c = filter_area-row_start.
    CONDENSE l_row_start_c NO-GAPS.

    l_col_start_c = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_start ) .
    l_col_end_c   = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_end ) .
    l_value = worksheet->get_title( ) .

    r_ref = zcl_excel_common=>escape_string( ip_value = l_value ).

    CONCATENATE r_ref '!$' l_col_start_c '$' l_row_start_c ':$' l_col_end_c '$' l_row_end_c INTO r_ref.

  ENDMETHOD.


  METHOD get_values.

    FIELD-SYMBOLS: <ls_filter> LIKE LINE OF me->mt_filters,
                   <ls_value>  LIKE LINE OF <ls_filter>-t_values.

    DATA: ls_filter LIKE LINE OF rt_filter.

    LOOP AT me->mt_filters ASSIGNING <ls_filter> WHERE rule = mc_filter_rule_single_values.

      ls_filter-column = <ls_filter>-column.
      LOOP AT <ls_filter>-t_values ASSIGNING <ls_value>.
        ls_filter-value = <ls_value>.
        APPEND ls_filter TO rt_filter.
      ENDLOOP.

    ENDLOOP.

  ENDMETHOD.


  METHOD is_row_hidden.


    DATA: lr_filter TYPE REF TO ts_filter,
          lv_col    TYPE i.

    FIELD-SYMBOLS: <ls_filter> TYPE ts_filter.

    rv_is_hidden = abap_false.

*--------------------------------------------------------------------*
* 1st row of filter area is never hidden, because here the filter
* symbol is being shown
*--------------------------------------------------------------------*
    IF iv_row = me->filter_area-row_start.
      RETURN.
    ENDIF.


    lv_col = me->filter_area-col_start.


    WHILE lv_col <= me->filter_area-col_end.

      lr_filter = me->get_column_filter( lv_col ).
      ASSIGN lr_filter->* TO <ls_filter>.

      CASE <ls_filter>-rule.

        WHEN mc_filter_rule_single_values.
          rv_is_hidden = me->is_row_hidden_single_values( iv_row    = iv_row
                                                          iv_col    = lv_col
                                                          is_filter = <ls_filter> ).

        WHEN mc_filter_rule_text_pattern.
          rv_is_hidden = me->is_row_hidden_text_pattern(  iv_row    = iv_row
                                                          iv_col    = lv_col
                                                          is_filter = <ls_filter> ).

      ENDCASE.

      IF rv_is_hidden = abap_true.
        RETURN.
      ENDIF.


      ADD 1 TO lv_col.

    ENDWHILE.


  ENDMETHOD.


  METHOD is_row_hidden_single_values.


    DATA: lv_value TYPE string.

    FIELD-SYMBOLS: <ls_sheet_content> LIKE LINE OF me->worksheet->sheet_content.

    rv_is_hidden = abap_false.   " Default setting is NOT HIDDEN = is in filter range

*--------------------------------------------------------------------*
* No filter values --> only symbol should be shown but nothing is being hidden
*--------------------------------------------------------------------*
    IF is_filter-t_values IS INITIAL.
      RETURN.
    ENDIF.

*--------------------------------------------------------------------*
* Get value of cell
*--------------------------------------------------------------------*
    READ TABLE me->worksheet->sheet_content ASSIGNING <ls_sheet_content> WITH TABLE KEY cell_row    = iv_row
                                                                                        cell_column = iv_col.
    IF sy-subrc = 0.
      lv_value = <ls_sheet_content>-cell_value.
    ELSE.
      CLEAR lv_value.
    ENDIF.

*--------------------------------------------------------------------*
* Check whether it is affected by filter
* this needs to be extended if we support other filtertypes
* other than single values
*--------------------------------------------------------------------*
    READ TABLE is_filter-t_values TRANSPORTING NO FIELDS WITH TABLE KEY table_line =  lv_value.
    IF sy-subrc <> 0.
      rv_is_hidden = abap_true.
    ENDIF.

  ENDMETHOD.


  METHOD is_row_hidden_text_pattern.



    DATA: lv_value TYPE string.

    FIELD-SYMBOLS: <ls_sheet_content> LIKE LINE OF me->worksheet->sheet_content.

    rv_is_hidden = abap_false.   " Default setting is NOT HIDDEN = is in filter range

*--------------------------------------------------------------------*
* Get value of cell
*--------------------------------------------------------------------*
    READ TABLE me->worksheet->sheet_content ASSIGNING <ls_sheet_content> WITH TABLE KEY cell_row    = iv_row
                                                                                        cell_column = iv_col.
    IF sy-subrc = 0.
      lv_value = <ls_sheet_content>-cell_value.
    ELSE.
      CLEAR lv_value.
    ENDIF.

*--------------------------------------------------------------------*
* Check whether it is affected by filter
* this needs to be extended if we support other filtertypes
* other than single values
*--------------------------------------------------------------------*
    IF lv_value NOT IN is_filter-tr_textfilter1.
      rv_is_hidden = abap_true.
    ENDIF.

  ENDMETHOD.


  METHOD set_filter_area.

    filter_area = is_area.

  ENDMETHOD.


  METHOD set_text_filter.
*  see method documentation how to use this

    DATA: lr_filter TYPE REF TO ts_filter,
          ls_value1 TYPE LINE OF ts_filter-tr_textfilter1.

    FIELD-SYMBOLS: <ls_filter> TYPE ts_filter.


    lr_filter = me->get_column_filter(  i_column ).
    ASSIGN lr_filter->* TO <ls_filter>.

    <ls_filter>-rule     = mc_filter_rule_text_pattern.
    CLEAR <ls_filter>-tr_textfilter1.

    IF iv_textfilter1 CA '*+'. " Pattern
      ls_value1-sign   = 'I'.
      ls_value1-option = 'CP'.
      ls_value1-low    = iv_textfilter1.
    ELSE.
      ls_value1-sign   = 'I'.
      ls_value1-option = 'EQ'.
      ls_value1-low    = iv_textfilter1.
    ENDIF.
    APPEND ls_value1 TO <ls_filter>-tr_textfilter1.

  ENDMETHOD.


  METHOD set_value.

    DATA: lr_filter TYPE REF TO ts_filter.

    FIELD-SYMBOLS: <ls_filter> TYPE ts_filter.


    lr_filter = me->get_column_filter(  i_column ).
    ASSIGN lr_filter->* TO <ls_filter>.

    <ls_filter>-rule     = mc_filter_rule_single_values.

    INSERT i_value INTO TABLE <ls_filter>-t_values.

  ENDMETHOD.


  METHOD set_values.

    FIELD-SYMBOLS: <ls_value> LIKE LINE OF it_values.

    LOOP AT it_values ASSIGNING <ls_value>.

      me->set_value( i_column = <ls_value>-column
                     i_value  = <ls_value>-value ).

    ENDLOOP.

  ENDMETHOD.


  METHOD validate_area.
    DATA: l_col                   TYPE zexcel_cell_column,
          ls_original_filter_area TYPE zexcel_s_autofilter_area,
          l_row                   TYPE zexcel_cell_row.

    l_row = worksheet->get_highest_row( ) .
    l_col = worksheet->get_highest_column( ) .

    IF filter_area IS INITIAL.
      filter_area-row_start = 1.
      filter_area-col_start = 1.
      filter_area-row_end   = l_row .
      filter_area-col_end   = l_col .
    ENDIF.

    IF filter_area-row_start > filter_area-row_end.
      ls_original_filter_area = filter_area.
      filter_area-row_start = ls_original_filter_area-row_end.
      filter_area-row_end = ls_original_filter_area-row_start.
    ENDIF.
    IF filter_area-row_start < 1.
      filter_area-row_start = 1.
    ENDIF.
    IF filter_area-col_start < 1.
      filter_area-col_start = 1.
    ENDIF.
    IF filter_area-row_end > l_row OR
       filter_area-row_end < 1.
      filter_area-row_end = l_row.
    ENDIF.
    IF filter_area-col_end > l_col OR
       filter_area-col_end < 1.
      filter_area-col_end = l_col.
    ENDIF.
    IF filter_area-col_start > filter_area-col_end.
      filter_area-col_start = filter_area-col_end.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
