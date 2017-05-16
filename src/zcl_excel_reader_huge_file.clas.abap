class ZCL_EXCEL_READER_HUGE_FILE definition
  public
  inheriting from ZCL_EXCEL_READER_2007
  create public .

public section.
protected section.

  methods LOAD_SHARED_STRINGS
    redefinition .
  methods LOAD_WORKSHEET
    redefinition .
private section.

  types:
    begin of t_cell_content,
      datatype type zexcel_cell_data_type,
      value    type zexcel_cell_value,
      formula  type zexcel_cell_formula,
      style    type zexcel_cell_style,
    end of t_cell_content .
  types:
    begin of t_cell_coord,
      row      type zexcel_cell_row,
      column   type zexcel_cell_column_alpha,
    end of t_cell_coord .
  types:
    begin of t_cell.
          include type t_cell_coord as coord.
          include type t_cell_content as content.
  types: end of t_cell .

  interface IF_SXML_NODE load .
  constants C_END_OF_STREAM type IF_SXML_NODE=>NODE_TYPE value IF_SXML_NODE=>CO_NT_FINAL. "#EC NOTEXT
  constants C_ELEMENT_OPEN type IF_SXML_NODE=>NODE_TYPE value IF_SXML_NODE=>CO_NT_ELEMENT_OPEN. "#EC NOTEXT
  constants C_ELEMENT_CLOSE type IF_SXML_NODE=>NODE_TYPE value IF_SXML_NODE=>CO_NT_ELEMENT_CLOSE. "#EC NOTEXT
  data:
    begin of gs_buffer_style,
    index type i value -1,
    guid type zexcel_cell_style,
    end of gs_buffer_style .
  constants C_ATTRIBUTE type IF_SXML_NODE=>NODE_TYPE value IF_SXML_NODE=>CO_NT_ATTRIBUTE. "#EC NOTEXT
  constants C_NODE_VALUE type IF_SXML_NODE=>NODE_TYPE value IF_SXML_NODE=>CO_NT_VALUE. "#EC NOTEXT

  methods SKIP_TO
    importing
      !IV_ELEMENT_NAME type STRING
      !IO_READER type ref to IF_SXML_READER
    raising
      LCX_NOT_FOUND .
  methods FILL_CELL_FROM_ATTRIBUTES
    importing
      !IO_READER type ref to IF_SXML_READER
    returning
      value(ES_CELL) type T_CELL
    raising
      LCX_NOT_FOUND .
  methods READ_SHARED_STRINGS
    importing
      !IO_READER type ref to IF_SXML_READER
    returning
      value(ET_SHARED_STRINGS) type STRINGTAB .
  methods GET_CELL_COORD
    importing
      !IV_COORD type STRING
    returning
      value(ES_COORD) type T_CELL_COORD .
  methods PUT_CELL_TO_WORKSHEET
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
      !IS_CELL type T_CELL .
  methods GET_SHARED_STRING
    importing
      !IV_INDEX type ANY
    returning
      value(EV_VALUE) type STRING
    raising
      LCX_NOT_FOUND .
  methods GET_STYLE
    importing
      !IV_INDEX type ANY
    returning
      value(EV_STYLE_GUID) type ZEXCEL_CELL_STYLE
    raising
      LCX_NOT_FOUND .
  methods READ_WORKSHEET_DATA
    importing
      !IO_READER type ref to IF_SXML_READER
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      LCX_NOT_FOUND .
  methods GET_SXML_READER
    importing
      !IV_PATH type STRING
    returning
      value(EO_READER) type ref to IF_SXML_READER
    raising
      ZCX_EXCEL .
ENDCLASS.



CLASS ZCL_EXCEL_READER_HUGE_FILE IMPLEMENTATION.


method FILL_CELL_FROM_ATTRIBUTES.

  while io_reader->node_type ne c_end_of_stream.
    io_reader->next_attribute( ).
    if io_reader->node_type ne c_attribute.
      exit.
    endif.
    case io_reader->name.
      when `t`.
        es_cell-datatype = io_reader->value.
      when `s`.
        if io_reader->value is not initial.
          es_cell-style = get_style( io_reader->value ).
        endif.
      when `r`.
        es_cell-coord = get_cell_coord( io_reader->value ).
    endcase.
  endwhile.

endmethod.


method GET_CELL_COORD.

  zcl_excel_common=>convert_columnrow2column_a_row(
    exporting
      i_columnrow = iv_coord
    importing
      e_column    = es_coord-column
      e_row       = es_coord-row
    ).

endmethod.


method GET_SHARED_STRING.
  data: lv_tabix type i,
        lv_error type string.
  lv_tabix = iv_index + 1.
  read table shared_strings into ev_value index lv_tabix.
  if sy-subrc ne 0.
    concatenate 'Entry ' iv_index ' not found in Shared String Table' into lv_error.
    raise exception type lcx_not_found
      exporting
        error = lv_error.
  endif.
endmethod.


method GET_STYLE.

  data: lv_tabix type i,
        lo_style type ref to zcl_excel_style,
        lv_error type string.

  if gs_buffer_style-index ne iv_index.
    lv_tabix = iv_index + 1.
    read table styles into lo_style index lv_tabix.
    if sy-subrc ne 0.
      concatenate 'Entry ' iv_index ' not found in Style Table' into lv_error.
      raise exception type lcx_not_found
        exporting
          error = lv_error.
    else.
      gs_buffer_style-index = iv_index.
      gs_buffer_style-guid  = lo_style->get_guid( ).
    endif.
  endif.

  ev_style_guid = gs_buffer_style-guid.

endmethod.


method GET_SXML_READER.

  data: lv_xml type xstring.

  lv_xml = get_from_zip_archive( iv_path ).
  eo_reader = cl_sxml_string_reader=>create( lv_xml ).

endmethod.


method LOAD_SHARED_STRINGS.

  data: lo_reader type ref to if_sxml_reader.

  lo_reader = get_sxml_reader( ip_path ).

  shared_strings = read_shared_strings( lo_reader ).

endmethod.


method LOAD_WORKSHEET.

  data: lo_reader type ref to if_sxml_reader.

  lo_reader = get_sxml_reader( ip_path ).

  read_worksheet_data( io_reader    = lo_reader
                       io_worksheet = io_worksheet ).

endmethod.


method PUT_CELL_TO_WORKSHEET.
  check is_cell-value is not initial
     or is_cell-formula is not initial
     or is_cell-style is not initial.
  call method io_worksheet->set_cell
    exporting
      ip_column    = is_cell-column
      ip_row       = is_cell-row
      ip_value     = is_cell-value
      ip_formula   = is_cell-formula
      ip_data_type = is_cell-datatype
      ip_style     = is_cell-style.
endmethod.


method read_shared_strings.

  data lv_value type string.

  while io_reader->node_type ne c_end_of_stream.
    io_reader->next_node( ).
    if io_reader->name eq `t`.
      case io_reader->node_type .
        when c_element_open .
          clear lv_value .
        when c_node_value .
          lv_value = lv_value && io_reader->value .
        when c_element_close .
          append lv_value to et_shared_strings.
      endcase .
    endif.
  endwhile.

endmethod.


method READ_WORKSHEET_DATA.

  data: ls_cell   type t_cell.

* Skip to <sheetData> element
  skip_to(  iv_element_name = `sheetData`  io_reader = io_reader ).

* Main loop: Evaluate the <c> elements and its children
  while io_reader->node_type ne c_end_of_stream.
    io_reader->next_node( ).
    case io_reader->node_type.
      when c_element_open.
        if io_reader->name eq `c`.
          ls_cell = fill_cell_from_attributes( io_reader ).
        endif.
      when c_node_value.
        case io_reader->name.
          when `f`.
            ls_cell-formula = io_reader->value.
          when `v`.
            if ls_cell-datatype eq `s`.
              ls_cell-value = get_shared_string( io_reader->value ).
            else.
              ls_cell-value = io_reader->value.
            endif.
          when `t` or `is`.
            ls_cell-value = io_reader->value.
        endcase.
      when c_element_close.
        case io_reader->name.
          when `c`.
            put_cell_to_worksheet( is_cell = ls_cell io_worksheet = io_worksheet ).
          when `sheetData`.
            exit.
        endcase.
    endcase.
  endwhile.

endmethod.


method SKIP_TO.

  data: lv_error type string.

* Skip forward to given element
  while io_reader->name ne iv_element_name or
        io_reader->node_type ne c_element_open.
    io_reader->next_node( ).
    if io_reader->node_type = c_end_of_stream.
      concatenate 'XML error: Didn''t find element <' iv_element_name '>' into lv_error.
      raise exception type lcx_not_found
        exporting
          error = lv_error.
    endif.
  endwhile.


endmethod.
ENDCLASS.
