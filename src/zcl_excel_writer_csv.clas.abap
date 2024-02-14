class zcl_excel_writer_csv definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
  public section.

    interfaces zif_excel_writer .

    class-methods set_delimiter
      importing
        value(ip_value) type c default ';' .
    class-methods set_enclosure
      importing
        value(ip_value) type c default '"' .
    class-methods set_endofline
      importing
        value(ip_value) type any default cl_abap_char_utilities=>cr_lf .
    class-methods set_active_sheet_index
      importing
        !i_active_worksheet type zexcel_active_worksheet .
    class-methods set_active_sheet_index_by_name
      importing
        !i_worksheet_name type zexcel_worksheets_name .
    class-methods set_initial_ext_date
      importing
        !ip_value type char10 default 'DEFAULT' .
  protected section.
*"* private components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  private section.

    data excel type ref to zcl_excel .
    class-data delimiter type c value ';' ##NO_TEXT.
    class-data enclosure type c value '"' ##NO_TEXT.
    class-data:
      eol type c length 2 value cl_abap_char_utilities=>cr_lf ##NO_TEXT.
    class-data worksheet_name type zexcel_worksheets_name .
    class-data worksheet_index type zexcel_active_worksheet .
    class-data initial_ext_date type char10 value 'DEFAULT' ##NO_TEXT.
    constants c_default type char10 value 'DEFAULT' ##NO_TEXT.

    methods create
      returning
        value(ep_excel) type xstring
      raising
        zcx_excel .
    methods create_csv
      returning
        value(ep_content) type xstring
      raising
        zcx_excel .
ENDCLASS.



CLASS ZCL_EXCEL_WRITER_CSV IMPLEMENTATION.


  method create.

* .csv format with ; delimiter

* Start of insertion # issue 1134 - Dateretention of cellstyles(issue #139)
    me->excel->add_static_styles( ).
* End of insertion # issue 1134 - Dateretention of cellstyles(issue #139)

    ep_excel = me->create_csv( ).

  endmethod.


  method create_csv.

    types: begin of lty_format,
             cmpname  type seocmpname,
             attvalue type seovalue,
           end of lty_format.
    data: lt_format type standard table of lty_format,
          ls_format like line of lt_format,
          lv_date   type d,
          lv_tmp    type string,
          lv_time   type c length 8.

    data: lo_iterator  type ref to zcl_excel_collection_iterator,
          lo_worksheet type ref to zcl_excel_worksheet.

    data: lt_cell_data type zexcel_t_cell_data_unsorted,
          lv_row       type i,
          lv_col       type i,
          lv_string    type string,
          lc_value     type string,
          lv_attrname  type seocmpname.

    data: ls_numfmt type zexcel_s_style_numfmt,
          lo_style  type ref to zcl_excel_style.

    field-symbols: <fs_sheet_content> type zexcel_s_cell_data.

* --- Retrieve supported cell format
    select * into corresponding fields of table lt_format
      from seocompodf
     where clsname  = 'ZCL_EXCEL_STYLE_NUMBER_FORMAT'
       and typtype  = 1
       and type     = 'ZEXCEL_NUMBER_FORMAT'.

* --- Retrieve SAP date format
    clear ls_format.
    select ddtext into ls_format-attvalue from dd07t where domname    = 'XUDATFM'
                                                       and ddlanguage = sy-langu.
      ls_format-cmpname = 'DATE'.
      condense ls_format-attvalue.
      concatenate '''' ls_format-attvalue '''' into ls_format-attvalue.
      append ls_format to lt_format.
    endselect.


    loop at lt_format into ls_format.
      translate ls_format-attvalue to upper case.
      modify lt_format from ls_format.
    endloop.


* STEP 1: Collect strings from the first worksheet
    lo_iterator = excel->get_worksheets_iterator( ).
    data: current_worksheet_title type zexcel_sheet_title.

    while lo_iterator->has_next( ) eq abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      if worksheet_name is not initial.
        current_worksheet_title = lo_worksheet->get_title( ).
        check current_worksheet_title = worksheet_name.
      else.
        if worksheet_index is initial.
          worksheet_index = 1.
        endif.
        check worksheet_index = sy-index.
      endif.
      append lines of lo_worksheet->sheet_content to lt_cell_data.
      exit. " Take first worksheet only
    endwhile.

    delete lt_cell_data where cell_formula is not initial. " delete formula content

    sort lt_cell_data by cell_row
                         cell_column.
    lv_row = 1.
    lv_col = 1.
    clear lv_string.
    loop at lt_cell_data assigning <fs_sheet_content>.

*   --- Retrieve Cell Style format and data type
      clear ls_numfmt.
      if <fs_sheet_content>-data_type is initial and <fs_sheet_content>-cell_style is not initial.
        lo_iterator = excel->get_styles_iterator( ).
        while lo_iterator->has_next( ) eq abap_true.
          lo_style ?= lo_iterator->get_next( ).
          check lo_style->get_guid( ) = <fs_sheet_content>-cell_style.
          ls_numfmt     = lo_style->number_format->get_structure( ).
          exit.
        endwhile.
      endif.
      if <fs_sheet_content>-data_type is initial and ls_numfmt is not initial.
        " determine data-type
        clear lv_attrname.
        concatenate '''' ls_numfmt-numfmt '''' into ls_numfmt-numfmt.
        translate ls_numfmt-numfmt to upper case.
        read table lt_format into ls_format with key attvalue = ls_numfmt-numfmt.
        if sy-subrc = 0.
          lv_attrname = ls_format-cmpname.
        endif.

        if lv_attrname is not initial.
          find first occurrence of 'DATETIME' in lv_attrname.
          if sy-subrc = 0.
            <fs_sheet_content>-data_type = 'd'.
          else.
            find first occurrence of 'TIME' in lv_attrname.
            if sy-subrc = 0.
              <fs_sheet_content>-data_type = 't'.
            else.
              find first occurrence of 'DATE' in lv_attrname.
              if sy-subrc = 0.
                <fs_sheet_content>-data_type = 'd'.
              else.
                find first occurrence of 'CURRENCY' in lv_attrname.
                if sy-subrc = 0.
                  <fs_sheet_content>-data_type = 'n'.
                else.
                  find first occurrence of 'NUMBER' in lv_attrname.
                  if sy-subrc = 0.
                    <fs_sheet_content>-data_type = 'n'.
                  else.
                    find first occurrence of 'PERCENTAGE' in lv_attrname.
                    if sy-subrc = 0.
                      <fs_sheet_content>-data_type = 'n'.
                    endif. " Purcentage
                  endif. " Number
                endif. " Currency
              endif. " Date
            endif. " TIME
          endif. " DATETIME
        endif. " lv_attrname IS NOT INITIAL.
      endif. " <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.

* --- Add empty rows
      while lv_row < <fs_sheet_content>-cell_row.
        concatenate lv_string zcl_excel_writer_csv=>eol into lv_string.
        lv_row = lv_row + 1.
        lv_col = 1.
      endwhile.

* --- Add empty columns
      while lv_col < <fs_sheet_content>-cell_column.
        concatenate lv_string zcl_excel_writer_csv=>delimiter into lv_string.
        lv_col = lv_col + 1.
      endwhile.

* ----- Use format to determine the data type and display format.
      case <fs_sheet_content>-data_type.

        when 'd' or 'D'.
          if <fs_sheet_content>-cell_value is initial and initial_ext_date <> c_default.
            lc_value = initial_ext_date.
          else.
            lc_value = zcl_excel_common=>excel_string_to_date( ip_value = <fs_sheet_content>-cell_value ).
            try.
                lv_date = lc_value.
                call function 'CONVERT_DATE_TO_EXTERNAL'
                  exporting
                    date_internal            = lv_date
                  importing
                    date_external            = lv_tmp
                  exceptions
                    date_internal_is_invalid = 1
                    others                   = 2.
                if sy-subrc = 0.
                  lc_value = lv_tmp.
                endif.

              catch cx_sy_conversion_no_number.

            endtry.
          endif.

        when 't' or 'T'.
          lc_value = zcl_excel_common=>excel_string_to_time( ip_value = <fs_sheet_content>-cell_value ).
          write lc_value to lv_time using edit mask '__:__:__'.
          lc_value = lv_time.
        when others.
          lc_value = <fs_sheet_content>-cell_value.

      endcase.

      concatenate zcl_excel_writer_csv=>enclosure zcl_excel_writer_csv=>enclosure into lv_tmp.
      condense lv_tmp.
      replace all occurrences of zcl_excel_writer_csv=>enclosure in lc_value with lv_tmp.

      find first occurrence of zcl_excel_writer_csv=>delimiter in lc_value.
      if sy-subrc = 0.
        concatenate lv_string zcl_excel_writer_csv=>enclosure lc_value zcl_excel_writer_csv=>enclosure into lv_string.
      else.
        concatenate lv_string lc_value into lv_string.
      endif.

    endloop.

    clear ep_content.

    call function 'SCMS_STRING_TO_XSTRING'
      exporting
        text   = lv_string
      importing
        buffer = ep_content
      exceptions
        failed = 1
        others = 2.

  endmethod.


  method set_active_sheet_index.
    clear worksheet_name.
    worksheet_index = i_active_worksheet.
  endmethod.


  method set_active_sheet_index_by_name.
    clear worksheet_index.
    worksheet_name = i_worksheet_name.
  endmethod.


  method set_delimiter.
    delimiter = ip_value.
  endmethod.


  method set_enclosure.
    zcl_excel_writer_csv=>enclosure = ip_value.
  endmethod.


  method set_endofline.
    zcl_excel_writer_csv=>eol = ip_value.
  endmethod.


  method zif_excel_writer~write_file.
    me->excel = io_excel.
    ep_file = me->create( ).
  endmethod.


  method set_initial_ext_date.
    initial_ext_date = ip_value.
  endmethod.
ENDCLASS.
