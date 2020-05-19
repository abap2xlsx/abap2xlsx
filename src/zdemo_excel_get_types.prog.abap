*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL_GET_TYPES
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel_get_types.

types
: tt_text type table of text80 with empty key
.

data  : lo_excel     type ref to zcl_excel
      , reader          type ref to zif_excel_reader
      , lo_template_filler type ref to zcl_excel_fill_template
      .


parameters: p_fpath type string obligatory lower case default 'C:\Users\sadfasdf\Desktop\abap2xlsx\ZABAP2XLSX_EXAMPLE.xlsx'.

parameters: p_normal radiobutton group rad1 default 'X'
          , p_other radiobutton group rad1
          .

at selection-screen on value-request for p_fpath.
  perform get_file_path changing p_fpath.


start-of-selection.

  create object reader type zcl_excel_reader_2007.
  lo_excel = reader->load_file( p_fpath ).

    create object lo_template_filler .

    lo_template_filler->get_range( lo_excel ).
    lo_template_filler->discard_overlapped( ).
    lo_template_filler->sign_range( ).
    lo_template_filler->find_var( lo_excel ).

  perform get_types.


*&---------------------------------------------------------------------*
*&      Form  Get_file_path
*&---------------------------------------------------------------------*
form get_file_path changing cv_path type string.
  clear cv_path.

  data:
    lv_rc          type  i,
    lv_user_action type  i,
    lt_file_table  type  filetable,
    ls_file_table  like line of lt_file_table.

  cl_gui_frontend_services=>file_open_dialog(
  exporting
    window_title        = 'select template  xlsx'
    multiselection      = ''
    default_extension   = '*.xlsx'
    file_filter         = 'Text file (*.xlsx)|*.xlsx|All (*.*)|*.*'
  changing
    file_table          = lt_file_table
    rc                  = lv_rc
    user_action         = lv_user_action
  exceptions
    others              = 1
    ).
  if sy-subrc = 0.
    if lv_user_action = cl_gui_frontend_services=>action_ok.
      if lt_file_table is not initial.
        read table lt_file_table into ls_file_table index 1.
        if sy-subrc = 0.
          cv_path = ls_file_table-filename.
        endif.
      endif.
    endif.
  endif.
endform.                    " Get_file_path

form get_types .

  data
        : lv_sum type i
        , lt_res type tt_text
        , lt_buf type tt_text
        .
  loop at lo_template_filler->MT_SHEET assigning field-symbol(<fs_sheet>).

    clear lv_sum.

    read table lo_template_filler->MT_RANGE transporting no fields with key sheet = <fs_sheet>.

    add sy-subrc to lv_sum.

    read table lo_template_filler->MT_VAR transporting no fields with key sheet = <fs_sheet>.

    add sy-subrc to lv_sum.

    check lv_sum <= 4.

    perform get_type_r using <fs_sheet>   0    changing lt_buf.

    append lines of lt_buf to lt_res.


  endloop.

    data
          : lv_lines type i
          .


  if p_normal is initial.
    read table lt_res assigning field-symbol(<fs_res>) index 1.
    translate <fs_res> using ',:'.
    insert initial line into lt_res assigning <fs_res> index 1.
    <fs_res> = 'TYPES'.
    append initial line to lt_res assigning <fs_res>.
    <fs_res> = '.'.

    lv_lines = lines( lt_res ) - 2.
    delete lt_res index lv_lines.
    delete lt_res index lv_lines.

  else.
    insert initial line into lt_res assigning <fs_res> index 1.
    <fs_res> = 'TYPES:'.

    lv_lines = lines( lt_res )  - 2.

    read table lt_res assigning <fs_res> index lv_lines.
    translate <fs_res> using ',.'.
    add 1 to lv_lines.

  endif.

  if p_normal is initial.
    append initial line to lt_res assigning <fs_res>.
    <fs_res> = 'DATA'.
    append initial line to lt_res assigning <fs_res>.
    <fs_res> = ': lo_data type ref to ZCL_EXCEL_TEMPLATE_DATA'.
    append initial line to lt_res assigning <fs_res>.
    <fs_res> = '.'.

  else.
    append initial line to lt_res assigning <fs_res>.
    <fs_res> = 'DATA: lo_data type ref to ZCL_EXCEL_TEMPLATE_DATA.'.
  endif.

  cl_demo_output=>new( 'TEXT' )->display( lt_res ).
endform.


form get_type_r using p_sheet type ZEXCEL_TEMPLATE_SHEET_TITLE
                      p_parent  type i
                changing ct_result type tt_text.

  clear ct_result.

  data
        : lt_buf type tt_text
        , lt_tmp type tt_text
        , lv_sum type i
        .


  loop at lo_template_filler->MT_RANGE assigning field-symbol(<fs_range>) where sheet = p_sheet
                                                        and parent = p_parent.

    perform get_type_r
                using
                   p_sheet
                   <fs_range>-id
                changing
                   lt_tmp.

    append lines of lt_tmp to lt_buf.


  endloop.

  add sy-subrc to lv_sum.

  data
        : lv_name type string
        .

  if p_parent = 0.
    lv_name = p_sheet.
  else.
    read table lo_template_filler->MT_RANGE assigning <fs_range> with key sheet = p_sheet
                                                      id = p_parent.
    lv_name = <fs_range>-name.
  endif.

  append initial line to lt_buf assigning field-symbol(<fs_buf>).
  if p_normal is initial.
    <fs_buf> = |, begin of t_{ lv_name }|.
  else.
    <fs_buf> = | begin of t_{ lv_name },|.
  endif.


  loop at lo_template_filler->MT_VAR assigning field-symbol(<fs_var>) where sheet = p_sheet
                                                    and parent = p_parent.

    append initial line to lt_buf assigning <fs_buf>.
    if p_normal is initial.
      <fs_buf> = |,     { <fs_var>-name } type string|.
    else.
      <fs_buf> = |     { <fs_var>-name } type string,|.
    endif.


  endloop.

  add sy-subrc to lv_sum.

  loop at lo_template_filler->MT_RANGE assigning <fs_range> where sheet = p_sheet
                                                        and parent = p_parent.

    append initial line to lt_buf assigning <fs_buf>.
    if p_normal is initial.
      <fs_buf> = |,     { <fs_range>-name } type tt_{ <fs_range>-name }|.
    else.
      <fs_buf> = |     { <fs_range>-name } type tt_{ <fs_range>-name },|.
    endif.


  endloop.

  if lv_sum > 4.
    append initial line to lt_buf assigning <fs_buf>.
    if p_normal is initial.
      <fs_buf> = |,     xz type i|.
    else.
      <fs_buf> = |     xz type i,|.
    endif.

  endif.


  append initial line to lt_buf assigning <fs_buf>.
  if p_normal is initial.
    <fs_buf> = |, end of t_{ lv_name }|.
  else.
    <fs_buf> = | end of t_{ lv_name },|.
  endif.

  append initial line to lt_buf assigning <fs_buf>.

  if p_parent ne 0.
    append initial line to lt_buf assigning <fs_buf>.
    if p_normal is initial.
      <fs_buf> = |, tt_{ lv_name } type table of  t_{ lv_name } with empty key|.
    else.
      <fs_buf> = | tt_{ lv_name } type table of  t_{ lv_name } with empty key,|.
    endif.

  endif.

  append initial line to lt_buf assigning <fs_buf>.


  ct_result = lt_buf.
endform.
