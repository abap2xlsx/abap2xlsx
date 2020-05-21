class ZCL_EXCEL_FILL_TEMPLATE definition
  public
  final
  create public .

public section.

  data MT_SHEET type ZEXCEL_TEMPLATE_T_SHEET_TITLE .
  data MT_RANGE type ZEXCEL_TEMPLATE_T_RANGE .
  data MT_VAR type ZEXCEL_TEMPLATE_T_VAR .
  data MO_EXCEL type ref to ZCL_EXCEL .

  methods GET_RANGE
    importing
      !IO_EXCEL type ref to ZCL_EXCEL .
  methods VALIDATE_RANGE
    importing
      !IO_RANGE type ref to ZCL_EXCEL_RANGE .
  methods DISCARD_OVERLAPPED .
  methods SIGN_RANGE .
  methods FIND_VAR
    importing
      !IO_EXCEL type ref to ZCL_EXCEL .
  methods FILL_SHEET
    importing
      !IV_DATA type ZEXCEL_S_TEMPLATE_DATA .
  methods FILL_RANGE
    importing
      !IV_SHEET type ZEXCEL_SHEET_TITLE
      !IV_PARENT type ZEXCEL_CELL_ROW
      !IV_DATA type DATA
      value(IV_RANGE_LENGTH) type ZEXCEL_CELL_ROW
      !IO_SHEET type ref to ZCL_EXCEL_WORKSHEET
    changing
      !CT_CELLS type ZEXCEL_TEMPLATE_T_CELL_DATA
      !CH_DIFF type ZEXCEL_CELL_ROW
      !CT_MERGED_CELLS type ZCL_EXCEL_WORKSHEET=>MTY_TS_MERGE .
protected section.
private section.
ENDCLASS.



CLASS ZCL_EXCEL_FILL_TEMPLATE IMPLEMENTATION.


  method discard_overlapped.
    data
      : lt_range type zexcel_template_t_range
      .

    sort mt_range by sheet  start  stop.

    loop at mt_range assigning field-symbol(<fs_range>).

      loop at mt_range assigning field-symbol(<fs_tmp>) where sheet = <fs_range>-sheet and
                                                              name ne <fs_range>-name and
                                                              stop >= <fs_range>-start  and
                                                              start < <fs_range>-start  and
                                                              stop < <fs_range>-stop   .
        .

      endloop.

      if sy-subrc ne 0.
        append <fs_range> to lt_range.
      endif.

    endloop.

    mt_range = lt_range.

    sort mt_range by sheet  start  stop descending.

  endmethod.


  method fill_range.


    data
          : tmp_cells_template type zexcel_template_t_cell_data
          , lt_cells_result    type zexcel_template_t_cell_data
          , tmp_cells          type zexcel_template_t_cell_data
          , ls_cell            type zexcel_s_cell_data

          , tmp_merged_cells_template type zcl_excel_worksheet=>mty_ts_merge
          , lt_merged_cells_result    type zcl_excel_worksheet=>mty_ts_merge
          , lt_merged_cells_half      type zcl_excel_worksheet=>mty_ts_merge
          , tmp_merged_cells          type zcl_excel_worksheet=>mty_ts_merge
          , ls_merged_cell            like line of tmp_merged_cells

          , lv_start type i
          , lv_stop type i


          , col_str type string

          .
    field-symbols
                   : <fs_table> type any table
                   .



    ch_diff = ch_diff +  iv_range_length .

    lv_start = 1.


* recursive fill nested range

    loop at mt_range assigning field-symbol(<fs_range>) where sheet = iv_sheet and
                                                              parent = iv_parent.


      lv_stop = <fs_range>-start - 1.

*      update cells before any range

      loop at ct_cells into ls_cell  where cell_row >= lv_start and cell_row <= lv_stop .
        ls_cell-cell_row =  ls_cell-cell_row + ch_diff.
        col_str = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        concatenate col_str ls_cell-cell_coords into ls_cell-cell_coords.
        condense ls_cell-cell_coords no-gaps.

        append ls_cell to lt_cells_result.
      endloop.



*      update merged cells before range

      loop at ct_merged_cells into ls_merged_cell where row_from >=  lv_start and row_to <= lv_stop.
        ls_merged_cell-row_from = ls_merged_cell-row_from + ch_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + ch_diff.

        append ls_merged_cell to lt_merged_cells_result.

      endloop.



      lv_start = <fs_range>-stop + 1.



      clear
      : tmp_cells_template
      , tmp_merged_cells_template
      .

*copy cell template
      loop at ct_cells into ls_cell where cell_row >= <fs_range>-start and cell_row <= <fs_range>-stop.
        append ls_cell to tmp_cells_template.
      endloop.

      loop at ct_merged_cells into ls_merged_cell where row_from >= <fs_range>-start and row_to <= <fs_range>-stop.
        append ls_merged_cell to tmp_merged_cells_template.
      endloop.


      assign component <fs_range>-name of structure iv_data to <fs_table>.
      check sy-subrc = 0.

      ch_diff = ch_diff - <fs_range>-length.

*merge each line of data table with template
      loop at <fs_table> assigning field-symbol(<fs_line>) .
*        make local copy
        tmp_cells = tmp_cells_template.
        tmp_merged_cells = tmp_merged_cells_template.

*fill data

        fill_range(
          exporting
            io_sheet = io_sheet
            iv_sheet  = iv_sheet
            iv_parent = <fs_range>-id
            iv_data   = <fs_line>
            iv_range_length = <fs_range>-length
          changing
            ct_cells  = tmp_cells
            ct_merged_cells  = tmp_merged_cells
            ch_diff = ch_diff ).

*collect data

        append lines of tmp_cells to lt_cells_result.
        append lines of tmp_merged_cells to lt_merged_cells_result.

      endloop.

    endloop.


    if <fs_range> is assigned.

      loop at ct_cells into ls_cell where cell_row > <fs_range>-stop .
        ls_cell-cell_row =  ls_cell-cell_row + ch_diff.
        col_str = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        concatenate col_str ls_cell-cell_coords into ls_cell-cell_coords.
        condense ls_cell-cell_coords no-gaps.

        append ls_cell to lt_cells_result.
      endloop.

      ct_cells = lt_cells_result.

      loop at ct_merged_cells into ls_merged_cell where row_from > <fs_range>-stop.
        ls_merged_cell-row_from = ls_merged_cell-row_from + ch_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + ch_diff.

        append ls_merged_cell to lt_merged_cells_result.
      endloop.

      ct_merged_cells = lt_merged_cells_result.

    else.

      loop at ct_cells assigning field-symbol(<fs_cell>).
        <fs_cell>-cell_row =  <fs_cell>-cell_row + ch_diff.
        col_str = zcl_excel_common=>convert_column2alpha( <fs_cell>-cell_column ).

        <fs_cell>-cell_coords = <fs_cell>-cell_row.
        concatenate col_str <fs_cell>-cell_coords into <fs_cell>-cell_coords.
        condense <fs_cell>-cell_coords no-gaps.
      endloop.

      loop at ct_merged_cells into ls_merged_cell .
        ls_merged_cell-row_from = ls_merged_cell-row_from + ch_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + ch_diff.

        append ls_merged_cell to lt_merged_cells_result.
      endloop.

      ct_merged_cells = lt_merged_cells_result.

    endif.


*check if variables in this range
    loop at mt_var transporting no fields where sheet = iv_sheet and parent = iv_parent.
      exit.
    endloop.

    if sy-subrc = 0.
*      replace variables of current range with data
      loop at ct_cells assigning <fs_cell>.

        data
              : result_tab type match_result_tab
              , lv_search type string
              , lv_var_name type string
              , lo_value_new     type ref to data
              , lo_addit    type ref to cl_abap_elemdescr
              .

        field-symbols
                       : <fs_numeric>       type numeric
                       .

        refresh result_tab.

        find all occurrences of regex '\[[^\]]*\]' in <fs_cell>-cell_value  results result_tab.

        sort result_tab by offset descending .

        loop at result_tab assigning field-symbol(<fs_result>).
          lv_search = <fs_cell>-cell_value+<fs_result>-offset(<fs_result>-length).
          lv_var_name = lv_search.

          translate lv_var_name to upper case.
          translate lv_var_name using '[ ] '.
          condense lv_var_name .

          assign component lv_var_name of structure iv_data to field-symbol(<fs_var>).
          check sy-subrc = 0.

          data
                : lv_str type string
                .

          lv_str = <fs_var>.


          replace all occurrences of lv_search in <fs_cell>-cell_value  with lv_str.


        endloop.

        check  lines( result_tab )  = 1.

        check  <fs_cell>-cell_value co '1234567890. '.

        describe field <fs_var> type data(lv_value_type).

        check  lv_value_type = 'I'
            or lv_value_type = 'P'
            or lv_value_type = 's'
            or lv_value_type = '8'
            or lv_value_type = 'a'
            or lv_value_type = 'e'
            or lv_value_type = 'F'.

        case lv_value_type.
          when cl_abap_typedescr=>typekind_int or cl_abap_typedescr=>typekind_int1 or cl_abap_typedescr=>typekind_int2
            or cl_abap_typedescr=>typekind_int8. "Allow INT8 types columns
            lo_addit = cl_abap_elemdescr=>get_i( ).
            create data lo_value_new type handle lo_addit.
            assign lo_value_new->* to <fs_numeric>.
            if sy-subrc = 0.
              <fs_numeric> = <fs_var>.
              <fs_cell>-cell_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
              clear <fs_cell>-data_type.
            endif.

          when cl_abap_typedescr=>typekind_float or cl_abap_typedescr=>typekind_packed.
            lo_addit = cl_abap_elemdescr=>get_f( ).
            create data lo_value_new type handle lo_addit.
            assign lo_value_new->* to <fs_numeric>.
            if sy-subrc = 0.
              <fs_numeric> = <fs_var>.
              <fs_cell>-cell_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
              clear <fs_cell>-data_type.
            endif.
        endcase.

      endloop.
    endif.


  endmethod.


  method fill_sheet.

    data
          : lo_worksheet    type ref to zcl_excel_worksheet
          .

    lo_worksheet = mo_excel->get_worksheet_by_name( iv_data-sheet ).

    assign iv_data-data->* to field-symbol(<fs_data>).


    data
          : lt_sheet_content type  zexcel_template_t_cell_data
          , lt_merged_cells type zcl_excel_worksheet=>mty_ts_merge
          .

    lt_sheet_content = lo_worksheet->sheet_content.
    lt_merged_cells = lo_worksheet->mt_merged_cells.

    assign iv_data-data->* to field-symbol(<any_data>).

    data
          : lv_dif type i
          .

    fill_range(
      exporting
        io_sheet = lo_worksheet
        iv_range_length = 0
        iv_sheet  = iv_data-sheet
        iv_parent = 0
        iv_data   = <any_data>
      changing
        ct_cells  = lt_sheet_content
        ct_merged_cells = lt_merged_cells
        ch_diff   = lv_dif  ).


    refresh  lo_worksheet->sheet_content.

    loop at lt_sheet_content assigning field-symbol(<fs_sheet_content>).
      insert <fs_sheet_content> into table lo_worksheet->sheet_content.
    endloop.

    refresh  lo_worksheet->mt_merged_cells.

    loop at lt_merged_cells assigning field-symbol(<fs_merged_cell>).
      insert <fs_merged_cell> into table lo_worksheet->mt_merged_cells.
    endloop.



  endmethod.


  method find_var.

    data
          : row type i
          , column type i
          , col_str type string
          , value type string

          , lo_worksheet type ref to zcl_excel_worksheet
          , ls_var TYPE zexcel_template_s_var
          .

    loop at mt_sheet assigning field-symbol(<fs_sheet>).



      lo_worksheet ?= io_excel->get_worksheet_by_name(  <fs_sheet> ).
      row = 1.
      column = 1.

      data(highest_column) = lo_worksheet->get_highest_column( ).
      data(highest_row)    = lo_worksheet->get_highest_row( ).

      while row <= highest_row.
        while column <= highest_column.
          col_str = zcl_excel_common=>convert_column2alpha( column ).
          lo_worksheet->get_cell(
            exporting
              ip_column = col_str
              ip_row    = row
            importing
              ep_value = value
          ).


          data
                : result_tab type match_result_tab
                , lv_search type string
                , lv_replace type string
                .

          find all occurrences of regex '\[[^\]]*\]' in value results result_tab.

          loop at result_tab assigning field-symbol(<fs_result>).
            lv_search = value+<fs_result>-offset(<fs_result>-length).
            lv_replace = lv_search.

            translate lv_replace to upper case.

            clear ls_var.

            ls_var-sheet = <fs_sheet>.
            ls_var-name = lv_replace.
            translate ls_var-name using '[ ] '.
            condense ls_var-name .

            loop at mt_range assigning field-symbol(<fs_range>) where sheet = <fs_sheet>
                                                                    and start <= row
                                                                    and stop >= row.
              ls_var-parent = <fs_range>-id.
              exit.
            endloop.

            read table mt_var transporting no fields with key sheet = ls_var-sheet name = ls_var-name parent = ls_var-parent.
            if sy-subrc ne 0.
              append ls_var to mt_var.
            endif.


          endloop.
          column = column + 1.
        endwhile.
        column = 1.
        row = row + 1.
      endwhile.
    endloop.

    sort mt_range by id .
  endmethod.


  method get_range.
    data
          : lo_worksheets_iterator type ref to cl_object_collection_iterator
          , lo_worksheet type ref to zcl_excel_worksheet
          , lo_range_iterator TYPE REF TO cl_object_collection_iterator
          , lo_range type ref to zcl_excel_range
          .

    mo_excel = IO_EXCEL.

    lo_worksheets_iterator = io_excel->get_worksheets_iterator( ).


    while lo_worksheets_iterator->has_next( ) = 'X'.
      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      append lo_worksheet->get_title( ) to mt_sheet.
    endwhile.

    lo_range_iterator = io_excel->get_ranges_iterator( ).

    while lo_range_iterator->has_next(  ) = 'X'.
      lo_range ?= lo_range_iterator->get_next( ).
      validate_range( lo_range ).
    endwhile.

  endmethod.


  method sign_range.
    loop at mt_range assigning field-symbol(<fs_range>).
      <fs_range>-id = sy-tabix.
    endloop.

    loop at mt_range assigning  <fs_range>.
      data
            : lv_tabix type i
            .
      lv_tabix = sy-tabix + 1.
      loop at mt_range assigning  field-symbol(<fs_range_tmp>)
                                  from lv_tabix
                                  where sheet = <fs_range>-sheet.

        if <fs_range_tmp>-start >= <fs_range>-start and <fs_range_tmp>-stop <= <fs_range>-stop.
          <fs_range_tmp>-parent = <fs_range>-id.
        endif.

      endloop.

    endloop.

    sort mt_range by id descending.
  endmethod.


  method validate_range.

    data
          : name type string
          , value type string
          , start type string
          , stop  type string
          , lv_sheet type string
          , lv_value type string
          , lt_str type table of string
          , lv_start type string
          , lv_stop  type string
          .

    name = io_range->name.
    translate name TO UPPER CASE.
    value = io_range->get_value( ).

    split value at '!' into lv_sheet lv_value.

    split lv_value at ':' into start stop.

    split start at '$' into table lt_str.

    if lines( lt_str ) > 2.
      return.
    endif.

    read table lt_str into lv_start index 2.

    if lv_start co '0123456789'.
      append initial line to mt_range assigning field-symbol(<fs_range>).
      <fs_range>-sheet = lv_sheet.
      <fs_range>-name = name.
      <fs_range>-start = lv_start.

      split stop at '$' into table lt_str.
      read table lt_str into lv_stop index 2.
      <fs_range>-stop = lv_stop.

      <fs_range>-length = <fs_range>-stop - <fs_range>-start + 1.

    endif.

  endmethod.
ENDCLASS.
