CLASS zcl_excel_fill_template DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES: tv_range_name TYPE c LENGTH 30,
           BEGIN OF ts_range,
             sheet  TYPE zexcel_sheet_title,
             name   TYPE tv_range_name,
             start  TYPE zexcel_cell_row,
             stop   TYPE zexcel_cell_row,
             id     TYPE i,
             parent TYPE i,
             length TYPE i,
           END OF ts_range,
           tt_ranges TYPE STANDARD TABLE OF ts_range WITH DEFAULT KEY,
           BEGIN OF ts_variable,
             sheet  TYPE zexcel_sheet_title,
             parent TYPE i,
             name   TYPE tv_range_name,
           END OF ts_variable,
           BEGIN OF ts_name_style,
             sheet           TYPE zexcel_sheet_title,
             name            TYPE tv_range_name,
             parent          TYPE i,
             numeric_counter TYPE i,
             date_counter    TYPE i,
             time_counter    TYPE i,
             text_counter    TYPE i,
           END OF ts_name_style,
           tt_variables    TYPE STANDARD TABLE OF ts_variable WITH DEFAULT KEY,
           tt_sheet_titles TYPE STANDARD TABLE OF zexcel_sheet_title WITH DEFAULT KEY,
           tt_name_styles  TYPE STANDARD TABLE OF ts_name_style WITH DEFAULT KEY.

    DATA mt_sheet TYPE tt_sheet_titles READ-ONLY.
    DATA mt_range TYPE tt_ranges READ-ONLY.
    DATA mt_var TYPE tt_variables READ-ONLY.
    DATA mo_excel TYPE REF TO zcl_excel READ-ONLY.
    DATA mt_name_styles TYPE tt_name_styles READ-ONLY.

    CLASS-METHODS create
      IMPORTING
        !io_excel                 TYPE REF TO zcl_excel
      RETURNING
        VALUE(eo_template_filler) TYPE REF TO zcl_excel_fill_template
      RAISING
        zcx_excel.

    METHODS fill_sheet
      IMPORTING
        !iv_data TYPE zcl_excel_template_data=>ts_template_data_sheet
      RAISING
        zcx_excel .

  PROTECTED SECTION.
  PRIVATE SECTION.

    TYPES: tt_cell_data_no_key TYPE STANDARD TABLE OF zexcel_s_cell_data WITH DEFAULT KEY.

    METHODS fill_range
      IMPORTING
        !iv_sheet              TYPE zexcel_sheet_title
        !iv_parent             TYPE zexcel_cell_row
        !iv_data               TYPE data
        VALUE(iv_range_length) TYPE zexcel_cell_row
        !io_sheet              TYPE REF TO zcl_excel_worksheet
      CHANGING
        !ct_cells              TYPE tt_cell_data_no_key
        !cv_diff               TYPE zexcel_cell_row
        !ct_merged_cells       TYPE zcl_excel_worksheet=>mty_ts_merge
      RAISING
        zcx_excel .
    METHODS get_range .
    METHODS validate_range
      IMPORTING
        !io_range TYPE REF TO zcl_excel_range .
    METHODS discard_overlapped .
    METHODS sign_range .
    METHODS find_var
      RAISING
        zcx_excel .

ENDCLASS.



CLASS zcl_excel_fill_template IMPLEMENTATION.


  METHOD create.

    CREATE OBJECT eo_template_filler .

    eo_template_filler->mo_excel = io_excel.
    eo_template_filler->get_range( ).
    eo_template_filler->discard_overlapped( ).
    eo_template_filler->sign_range( ).
    eo_template_filler->find_var( ).

  ENDMETHOD.


  METHOD discard_overlapped.
    DATA:
       lt_range TYPE tt_ranges.
    FIELD-SYMBOLS:
      <ls_range>   TYPE ts_range,
      <ls_range_2> TYPE ts_range.

    SORT mt_range BY sheet  start  stop.

    LOOP AT mt_range ASSIGNING <ls_range>.

      LOOP AT mt_range ASSIGNING <ls_range_2> WHERE sheet =  <ls_range>-sheet
                                            AND name  <> <ls_range>-name
                                            AND stop  >= <ls_range>-start
                                            AND start <  <ls_range>-start
                                            AND stop  <  <ls_range>-stop.
        EXIT.
      ENDLOOP.

      IF sy-subrc NE 0.
        APPEND <ls_range> TO lt_range.
      ENDIF.

    ENDLOOP.

    mt_range = lt_range.

    SORT mt_range BY sheet  start  stop DESCENDING.

  ENDMETHOD.


  METHOD fill_range.

    DATA: lt_tmp_cells_template        TYPE tt_cell_data_no_key,
          lt_cells_result              TYPE tt_cell_data_no_key,
          lt_tmp_cells                 TYPE tt_cell_data_no_key,
          ls_cell                      TYPE zexcel_s_cell_data,
          lt_tmp_merged_cells_template TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_merged_cells_result       TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_tmp_merged_cells          TYPE zcl_excel_worksheet=>mty_ts_merge,
          ls_merged_cell               LIKE LINE OF lt_tmp_merged_cells,
          lv_start_row                 TYPE i,
          lv_stop_row                  TYPE i,
          lv_cell_row                  TYPE i,
          lv_column_alpha              TYPE string,
          lt_matches                   TYPE match_result_tab,
          lv_search                    TYPE string,
          lv_var_name                  TYPE string,
          lv_cell_value                TYPE string.

    FIELD-SYMBOLS:
      <table>     TYPE ANY TABLE,
      <line>      TYPE any,
      <ls_range>  TYPE ts_range,
      <ls_cell>   TYPE zexcel_s_cell_data,
      <ls_match>  TYPE match_result,
      <var_value> TYPE any.


    cv_diff = cv_diff +  iv_range_length .

    lv_start_row = 1.


* recursive fill nested range

    LOOP AT mt_range ASSIGNING <ls_range> WHERE sheet = iv_sheet
                                            AND parent = iv_parent.


      lv_stop_row = <ls_range>-start - 1.

*      update cells before any range

      LOOP AT ct_cells INTO ls_cell  WHERE cell_row >= lv_start_row AND cell_row <= lv_stop_row .
        ls_cell-cell_row =  ls_cell-cell_row + cv_diff.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        CONCATENATE lv_column_alpha ls_cell-cell_coords INTO ls_cell-cell_coords.
        CONDENSE ls_cell-cell_coords NO-GAPS.

        APPEND ls_cell TO lt_cells_result.
      ENDLOOP.



*      update merged cells before range

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from >=  lv_start_row AND row_to <= lv_stop_row.
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.

      ENDLOOP.



      lv_start_row = <ls_range>-stop + 1.



      CLEAR:
       lt_tmp_cells_template,
       lt_tmp_merged_cells_template.


*copy cell template
      LOOP AT ct_cells INTO ls_cell WHERE cell_row >= <ls_range>-start AND cell_row <= <ls_range>-stop.
        APPEND ls_cell TO lt_tmp_cells_template.
      ENDLOOP.

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from >= <ls_range>-start AND row_to <= <ls_range>-stop.
        APPEND ls_merged_cell TO lt_tmp_merged_cells_template.
      ENDLOOP.


      ASSIGN COMPONENT <ls_range>-name OF STRUCTURE iv_data TO <table>.
      CHECK sy-subrc = 0.

      cv_diff = cv_diff - <ls_range>-length.

*merge each line of data table with template
      LOOP AT <table> ASSIGNING <line>.
*        make local copy
        lt_tmp_cells = lt_tmp_cells_template.
        lt_tmp_merged_cells = lt_tmp_merged_cells_template.

*fill data

        fill_range(
          EXPORTING
            io_sheet        = io_sheet
            iv_sheet        = iv_sheet
            iv_parent       = <ls_range>-id
            iv_data         = <line>
            iv_range_length = <ls_range>-length
          CHANGING
            ct_cells        = lt_tmp_cells
            ct_merged_cells = lt_tmp_merged_cells
            cv_diff         = cv_diff ).

*collect data

        APPEND LINES OF lt_tmp_cells TO lt_cells_result.
        APPEND LINES OF lt_tmp_merged_cells TO lt_merged_cells_result.

      ENDLOOP.

    ENDLOOP.


    IF <ls_range> IS ASSIGNED.

      LOOP AT ct_cells INTO ls_cell WHERE cell_row > <ls_range>-stop .
        ls_cell-cell_row =  ls_cell-cell_row + cv_diff.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        CONCATENATE lv_column_alpha ls_cell-cell_coords INTO ls_cell-cell_coords.
        CONDENSE ls_cell-cell_coords NO-GAPS.

        APPEND ls_cell TO lt_cells_result.
      ENDLOOP.

      ct_cells = lt_cells_result.

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from > <ls_range>-stop.
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.
      ENDLOOP.

      ct_merged_cells = lt_merged_cells_result.

    ELSE.


      LOOP AT ct_cells ASSIGNING <ls_cell>.
        <ls_cell>-cell_row =  <ls_cell>-cell_row + cv_diff.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( <ls_cell>-cell_column ).

        <ls_cell>-cell_coords = <ls_cell>-cell_row.
        CONCATENATE lv_column_alpha <ls_cell>-cell_coords INTO <ls_cell>-cell_coords.
        CONDENSE <ls_cell>-cell_coords NO-GAPS.
      ENDLOOP.

      LOOP AT ct_merged_cells INTO ls_merged_cell .
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.
      ENDLOOP.

      ct_merged_cells = lt_merged_cells_result.

    ENDIF.


*check if variables in this range
    READ TABLE mt_var TRANSPORTING NO FIELDS WITH KEY sheet = iv_sheet parent = iv_parent.

    IF sy-subrc = 0.

*      replace variables of current range with data
      LOOP AT ct_cells ASSIGNING <ls_cell>.

        CLEAR lt_matches.

        lv_cell_value = <ls_cell>-cell_value.

        FIND ALL OCCURRENCES OF REGEX '\[[^\]]*\]' IN <ls_cell>-cell_value  RESULTS lt_matches.

        SORT lt_matches BY offset DESCENDING .

        LOOP AT lt_matches ASSIGNING <ls_match>.
          lv_search = <ls_cell>-cell_value+<ls_match>-offset(<ls_match>-length).
          lv_var_name = lv_search.

          TRANSLATE lv_var_name TO UPPER CASE.
          TRANSLATE lv_var_name USING '[ ] '.
          CONDENSE lv_var_name .

          ASSIGN COMPONENT lv_var_name OF STRUCTURE iv_data TO <var_value>.
          CHECK sy-subrc = 0.

          " Use SET_CELL to format correctly
          io_sheet->set_cell( ip_column = <ls_cell>-cell_column ip_row = <ls_cell>-cell_row - cv_diff ip_value = <var_value> ).
          lv_cell_row = <ls_cell>-cell_row - cv_diff.
          READ TABLE io_sheet->sheet_content INTO ls_cell
            WITH KEY cell_column = <ls_cell>-cell_column
                     cell_row    = lv_cell_row.
          REPLACE ALL OCCURRENCES OF lv_search IN <ls_cell>-cell_value WITH ls_cell-cell_value.
        ENDLOOP.

        IF lines( lt_matches ) = 1.
          lv_cell_row = <ls_cell>-cell_row - cv_diff.
          READ TABLE io_sheet->sheet_content INTO ls_cell
            WITH KEY cell_column = <ls_cell>-cell_column
                     cell_row    = lv_cell_row.
          <ls_cell>-data_type = ls_cell-data_type.
        ENDIF.

      ENDLOOP.
    ENDIF.

  ENDMETHOD.


  METHOD fill_sheet.

    DATA: lo_worksheet      TYPE REF TO zcl_excel_worksheet,
          lt_sheet_cells    TYPE tt_cell_data_no_key,
          lt_merged_cells   TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_merged_cells_2 TYPE zcl_excel_worksheet=>mty_ts_merge,
          lv_initial_diff   TYPE i.
    FIELD-SYMBOLS:
      <any_data>       TYPE any,
      <ls_sheet_cell>  TYPE zexcel_s_cell_data,
      <ls_merged_cell> TYPE zcl_excel_worksheet=>mty_merge.


    lo_worksheet = mo_excel->get_worksheet_by_name( iv_data-sheet ).

    lt_sheet_cells = lo_worksheet->sheet_content.
    lt_merged_cells = lo_worksheet->mt_merged_cells.

    ASSIGN iv_data-data->* TO <any_data>.

    fill_range(
      EXPORTING
        io_sheet        = lo_worksheet
        iv_range_length = 0
        iv_sheet        = iv_data-sheet
        iv_parent       = 0
        iv_data         = <any_data>
      CHANGING
        ct_cells        = lt_sheet_cells
        ct_merged_cells = lt_merged_cells
        cv_diff         = lv_initial_diff  ).


    CLEAR lo_worksheet->sheet_content.

    LOOP AT lt_sheet_cells ASSIGNING <ls_sheet_cell>.
      INSERT <ls_sheet_cell> INTO TABLE lo_worksheet->sheet_content.
    ENDLOOP.

    lt_merged_cells_2 = lo_worksheet->mt_merged_cells.
    LOOP AT lt_merged_cells_2 ASSIGNING <ls_merged_cell>.
      lo_worksheet->delete_merge( ip_cell_column = <ls_merged_cell>-col_from ip_cell_row = <ls_merged_cell>-row_from ).
    ENDLOOP.

    LOOP AT lt_merged_cells ASSIGNING <ls_merged_cell>.
      lo_worksheet->set_merge(
          ip_column_start = <ls_merged_cell>-col_from
          ip_column_end   = <ls_merged_cell>-col_to
          ip_row          = <ls_merged_cell>-row_from
          ip_row_to       = <ls_merged_cell>-row_to ).
    ENDLOOP.

  ENDMETHOD.


  METHOD find_var.

    DATA: lv_row            TYPE i,
          lv_column         TYPE i,
          lv_column_alpha   TYPE string,
          lv_value          TYPE string,
          ls_name_style     TYPE ts_name_style,
          lo_style          TYPE REF TO zcl_excel_style,
          lo_worksheet      TYPE REF TO zcl_excel_worksheet,
          ls_variable       TYPE ts_variable,
          lv_highest_column TYPE zexcel_cell_column,
          lv_highest_row    TYPE int4,
          lt_matches        TYPE match_result_tab,
          lv_search         TYPE string,
          lv_replace        TYPE string.

    FIELD-SYMBOLS:
      <ls_match>      TYPE match_result,
      <ls_range>      TYPE ts_range,
      <lv_sheet>      TYPE zexcel_sheet_title,
      <ls_name_style> TYPE ts_name_style.


    LOOP AT mt_sheet ASSIGNING <lv_sheet>.

      lo_worksheet ?= mo_excel->get_worksheet_by_name(  <lv_sheet> ).
      lv_row = 1.

      lv_highest_column = lo_worksheet->get_highest_column( ).
      lv_highest_row    = lo_worksheet->get_highest_row( ).

      WHILE lv_row <= lv_highest_row.
        lv_column = 1.
        WHILE lv_column <= lv_highest_column.
          lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column ).
          CLEAR lo_style.
          lo_worksheet->get_cell(
            EXPORTING
              ip_column = lv_column_alpha
              ip_row    = lv_row
            IMPORTING
              ep_value = lv_value
              ep_style = lo_style ).

          FIND ALL OCCURRENCES OF REGEX '\[[^\]]*\]' IN lv_value RESULTS lt_matches.

          LOOP AT lt_matches ASSIGNING <ls_match>.
            lv_search = lv_value+<ls_match>-offset(<ls_match>-length).
            lv_replace = lv_search.

            TRANSLATE lv_replace TO UPPER CASE.

            CLEAR ls_variable.

            ls_variable-sheet = <lv_sheet>.
            ls_variable-name = lv_replace.
            TRANSLATE ls_variable-name USING '[ ] '.
            CONDENSE ls_variable-name .

            LOOP AT mt_range ASSIGNING <ls_range> WHERE sheet = <lv_sheet>
                                                    AND start <= lv_row
                                                    AND stop >= lv_row.
              ls_variable-parent = <ls_range>-id.
              EXIT.
            ENDLOOP.

            READ TABLE mt_var TRANSPORTING NO FIELDS WITH KEY sheet = ls_variable-sheet name = ls_variable-name parent = ls_variable-parent.
            IF sy-subrc NE 0.
              APPEND ls_variable TO mt_var.
            ENDIF.

            READ TABLE mt_name_styles WITH KEY sheet = ls_variable-sheet name = ls_variable-name parent = ls_variable-parent ASSIGNING <ls_name_style>.
            IF sy-subrc NE 0.
              CLEAR ls_name_style.
              ls_name_style-sheet = <lv_sheet>.
              ls_name_style-name = ls_variable-name.
              ls_name_style-parent = ls_variable-parent.
              APPEND ls_name_style TO mt_name_styles ASSIGNING <ls_name_style>.
            ENDIF.
            IF lo_style IS NOT BOUND.
              <ls_name_style>-text_counter = <ls_name_style>-text_counter + 1.
            ELSE.
              IF lo_style->number_format->format_code CA '0'
                   AND lo_style->number_format->format_code NS '0]'.
                <ls_name_style>-numeric_counter = <ls_name_style>-numeric_counter + 1.
              ELSEIF lo_style->number_format->format_code CA 'm'
                 AND lo_style->number_format->format_code CA 'd'
                 AND lo_style->number_format->format_code NA 'h'.
                <ls_name_style>-date_counter = <ls_name_style>-date_counter + 1.
              ELSEIF ( lo_style->number_format->format_code CA 'h' OR lo_style->number_format->format_code CA 's' )
                 AND lo_style->number_format->format_code NA 'd'.
                <ls_name_style>-time_counter = <ls_name_style>-time_counter + 1.
              ELSE.
                <ls_name_style>-text_counter = <ls_name_style>-text_counter + 1.
              ENDIF.
            ENDIF.

          ENDLOOP.
          lv_column = lv_column + 1.
        ENDWHILE.
        lv_row = lv_row + 1.
      ENDWHILE.
    ENDLOOP.

    SORT mt_range BY id .
  ENDMETHOD.


  METHOD get_range.

    DATA:
      lo_worksheets_iterator TYPE REF TO zcl_excel_collection_iterator,
      lo_worksheet           TYPE REF TO zcl_excel_worksheet,
      lo_range_iterator      TYPE REF TO zcl_excel_collection_iterator,
      lo_range               TYPE REF TO zcl_excel_range.


    lo_worksheets_iterator = mo_excel->get_worksheets_iterator( ).


    WHILE lo_worksheets_iterator->has_next( ) = abap_true.
      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      APPEND lo_worksheet->get_title( ) TO mt_sheet.
    ENDWHILE.

    lo_range_iterator = mo_excel->get_ranges_iterator( ).

    WHILE lo_range_iterator->has_next(  ) = abap_true.
      lo_range ?= lo_range_iterator->get_next( ).
      validate_range( lo_range ).
    ENDWHILE.

  ENDMETHOD.


  METHOD sign_range.

    DATA: lv_tabix TYPE i.
    FIELD-SYMBOLS:
      <ls_range>   TYPE ts_range,
      <ls_range_2> TYPE ts_range.

    LOOP AT mt_range ASSIGNING <ls_range>.
      <ls_range>-id = sy-tabix.
    ENDLOOP.

    LOOP AT mt_range ASSIGNING  <ls_range>.
      lv_tabix = sy-tabix + 1.

      LOOP AT mt_range ASSIGNING  <ls_range_2>
                                  FROM lv_tabix
                                  WHERE sheet = <ls_range>-sheet.

        IF <ls_range_2>-start >= <ls_range>-start AND <ls_range_2>-stop <= <ls_range>-stop.
          <ls_range_2>-parent = <ls_range>-id.
        ENDIF.

      ENDLOOP.

    ENDLOOP.

    SORT mt_range BY id DESCENDING.
  ENDMETHOD.


  METHOD validate_range.

    DATA: lv_range_name       TYPE string,
          lv_range_address    TYPE string,
          lv_range_start      TYPE string,
          lv_range_stop       TYPE string,
          lv_range_sheet      TYPE string,
          lv_tmp_value        TYPE string,
          lt_cell_coord_parts TYPE TABLE OF string,
          lv_cell_coord_start TYPE string,
          lv_cell_coord_stop  TYPE string,
          lv_column_start     TYPE zexcel_cell_column_alpha,
          lv_column_end       TYPE zexcel_cell_column_alpha,
          lv_row_start        TYPE zexcel_cell_row,
          lv_row_end          TYPE zexcel_cell_row.

    FIELD-SYMBOLS: <ls_range> TYPE ts_range.


    lv_range_name = io_range->name.
    TRANSLATE lv_range_name TO UPPER CASE.
    lv_range_address = io_range->get_value( ).

    SPLIT lv_range_address AT '!' INTO lv_range_sheet lv_tmp_value.

    SPLIT lv_tmp_value AT ':' INTO lv_range_start lv_range_stop.

    SPLIT lv_range_start AT '$' INTO TABLE lt_cell_coord_parts.

    IF lines( lt_cell_coord_parts ) > 2.
      TRY.
          zcl_excel_common=>convert_range2column_a_row(
            EXPORTING
              i_range        = lv_range_address
            IMPORTING
              e_column_start = lv_column_start
              e_column_end   = lv_column_end
              e_row_start    = lv_row_start
              e_row_end      = lv_row_end
          ).
        CATCH zcx_excel.    "
          RETURN.
      ENDTRY.
      IF lv_column_start = 'A' AND lv_column_end = 'XFD'.
        lv_cell_coord_start = |{ lv_row_start }|.
        lv_cell_coord_stop = |{ lv_row_end }|.
        CLEAR lt_cell_coord_parts.
        CLEAR lv_range_stop.
      ELSE.
        RETURN.
      ENDIF.
    ENDIF.

    IF lines( lt_cell_coord_parts ) >= 2.
      READ TABLE lt_cell_coord_parts INTO lv_cell_coord_start INDEX 2.
    ENDIF.

    IF lv_cell_coord_start CO '0123456789'.
      APPEND INITIAL LINE TO mt_range ASSIGNING <ls_range>.
      <ls_range>-sheet = lv_range_sheet.
      <ls_range>-name = lv_range_name.
      <ls_range>-start = lv_cell_coord_start.

      SPLIT lv_range_stop AT '$' INTO TABLE lt_cell_coord_parts.
      READ TABLE lt_cell_coord_parts INTO lv_cell_coord_stop INDEX 2.
      <ls_range>-stop = lv_cell_coord_stop.

      <ls_range>-length = <ls_range>-stop - <ls_range>-start + 1.

    ENDIF.

  ENDMETHOD.
ENDCLASS.
