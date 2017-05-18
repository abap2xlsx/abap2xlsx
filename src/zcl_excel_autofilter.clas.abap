class ZCL_EXCEL_AUTOFILTER definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_AUTOFILTER
*"* do not include other source files here!!!
public section.

  types TV_FILTER_RULE type STRING .
  types TV_LOGICAL_OPERATOR type CHAR3 .
  types:
    BEGIN OF ts_filter,
        column            TYPE zexcel_cell_column,
        rule              TYPE tv_filter_rule,
        t_values          TYPE HASHED TABLE OF zexcel_cell_value WITH UNIQUE KEY table_line,
        tr_textfilter1    TYPE range of string,
        logical_operator  TYPE tv_logical_operator,
        tr_textfilter2   TYPE range of string,
      END OF ts_filter .
  types:
    tt_filters TYPE HASHED TABLE OF ts_filter WITH UNIQUE KEY column .

  data FILTER_AREA type ZEXCEL_S_AUTOFILTER_AREA .
  constants MC_FILTER_RULE_SINGLE_VALUES type TV_FILTER_RULE value 'single_values'. "#EC NOTEXT
  constants MC_FILTER_RULE_TEXT_PATTERN type TV_FILTER_RULE value 'text_pattern'. "#EC NOTEXT
  constants MC_LOGICAL_OPERATOR_AND type TV_LOGICAL_OPERATOR value 'and'. "#EC NOTEXT
  constants MC_LOGICAL_OPERATOR_NONE type TV_LOGICAL_OPERATOR value SPACE. "#EC NOTEXT
  constants MC_LOGICAL_OPERATOR_OR type TV_LOGICAL_OPERATOR value 'or'. "#EC NOTEXT

  methods CONSTRUCTOR
    importing
      !IO_SHEET type ref to ZCL_EXCEL_WORKSHEET .
  methods GET_FILTER_AREA
    returning
      value(RS_AREA) type ZEXCEL_S_AUTOFILTER_AREA .
  methods GET_FILTER_RANGE
    returning
      value(R_RANGE) type ZEXCEL_CELL_VALUE .
  methods GET_FILTER_REFERENCE
    returning
      value(R_REF) type ZEXCEL_RANGE_VALUE .
  methods GET_VALUES
    returning
      value(RT_FILTER) type ZEXCEL_T_AUTOFILTER_VALUES .
  type-pools ABAP .
  methods IS_ROW_HIDDEN
    importing
      !IV_ROW type ZEXCEL_CELL_ROW
    returning
      value(RV_IS_HIDDEN) type ABAP_BOOL .
  methods SET_FILTER_AREA
    importing
      !IS_AREA type ZEXCEL_S_AUTOFILTER_AREA .
  methods SET_TEXT_FILTER
    importing
      !I_COLUMN type ZEXCEL_CELL_COLUMN
      !IV_TEXTFILTER1 type CLIKE
      !IV_LOGICAL_OPERATOR type TV_LOGICAL_OPERATOR default MC_LOGICAL_OPERATOR_NONE
      !IV_TEXTFILTER2 type CLIKE optional .
  methods SET_VALUE
    importing
      !I_COLUMN type ZEXCEL_CELL_COLUMN
      !I_VALUE type ZEXCEL_CELL_VALUE .
  methods SET_VALUES
    importing
      !IT_VALUES type ZEXCEL_T_AUTOFILTER_VALUES .
*"* protected components of class ZABAP_EXCEL_WORKSHEET
*"* do not include other source files here!!!
protected section.

  methods GET_COLUMN_FILTER
    importing
      !I_COLUMN type ZEXCEL_CELL_COLUMN
    returning
      value(RR_FILTER) type ref to TS_FILTER .
  methods IS_ROW_HIDDEN_SINGLE_VALUES
    importing
      !IV_ROW type ZEXCEL_CELL_ROW
      !IV_COL type ZEXCEL_CELL_COLUMN
      !IS_FILTER type TS_FILTER
    returning
      value(RV_IS_HIDDEN) type ABAP_BOOL .
  methods IS_ROW_HIDDEN_TEXT_PATTERN
    importing
      !IV_ROW type ZEXCEL_CELL_ROW
      !IV_COL type ZEXCEL_CELL_COLUMN
      !IS_FILTER type TS_FILTER
    returning
      value(RV_IS_HIDDEN) type ABAP_BOOL .
*"* private components of class ZCL_EXCEL_AUTOFILTER
*"* do not include other source files here!!!
private section.

  data WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
  data MT_FILTERS type TT_FILTERS .

  methods VALIDATE_AREA .
ENDCLASS.



CLASS ZCL_EXCEL_AUTOFILTER IMPLEMENTATION.


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
  DATA: l_col TYPE zexcel_cell_column,
        l_row TYPE zexcel_cell_row.

  l_row = worksheet->get_highest_row( ) .
  l_col = worksheet->get_highest_column( ) .

  IF filter_area IS INITIAL.
    filter_area-row_start = 1.
    filter_area-col_start = 1.
    filter_area-row_end   = l_row .
    filter_area-col_end   = l_col .
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
  IF filter_area-row_start >= filter_area-row_end.
    filter_area-row_start = filter_area-row_end - 1.
    IF filter_area-row_start < 1.
      filter_area-row_start = 1.
      filter_area-row_end = 2.
    ENDIF.
  ENDIF.
  IF filter_area-col_start > filter_area-col_end.
    filter_area-col_start = filter_area-col_end.
  ENDIF.
ENDMETHOD.
ENDCLASS.
