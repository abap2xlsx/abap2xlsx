CLASS zcl_excel_reader_huge_file DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_reader_2007
  CREATE PUBLIC .

  PUBLIC SECTION.
  PROTECTED SECTION.

    METHODS load_shared_strings
        REDEFINITION .
    METHODS load_worksheet
        REDEFINITION .
  PRIVATE SECTION.

    TYPES:
      BEGIN OF t_cell_content,
        datatype TYPE zexcel_cell_data_type,
        value    TYPE zexcel_cell_value,
        formula  TYPE zexcel_cell_formula,
        style    TYPE zexcel_cell_style,
      END OF t_cell_content .
    TYPES:
      BEGIN OF t_cell_coord,
        row    TYPE zexcel_cell_row,
        column TYPE zexcel_cell_column_alpha,
      END OF t_cell_coord .
    TYPES:
      BEGIN OF t_cell.
        INCLUDE TYPE t_cell_coord AS coord.
        INCLUDE TYPE t_cell_content AS content.
    TYPES: END OF t_cell .

    INTERFACE if_sxml_node LOAD .
    CONSTANTS c_end_of_stream TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_final. "#EC NOTEXT
    CONSTANTS c_element_open TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_element_open. "#EC NOTEXT
    CONSTANTS c_element_close TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_element_close. "#EC NOTEXT
    DATA:
      BEGIN OF gs_buffer_style,
        index TYPE i VALUE -1,
        guid  TYPE zexcel_cell_style,
      END OF gs_buffer_style .
    CONSTANTS c_attribute TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_attribute. "#EC NOTEXT
    CONSTANTS c_node_value TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_value. "#EC NOTEXT

    METHODS skip_to
      IMPORTING
        !iv_element_name TYPE string
        !io_reader       TYPE REF TO if_sxml_reader
      RAISING
        lcx_not_found .
    METHODS fill_cell_from_attributes
      IMPORTING
        !io_reader     TYPE REF TO if_sxml_reader
      RETURNING
        VALUE(es_cell) TYPE t_cell
      RAISING
        lcx_not_found
        zcx_excel.
    METHODS read_shared_strings
      IMPORTING
        !io_reader               TYPE REF TO if_sxml_reader
      RETURNING
        VALUE(et_shared_strings) TYPE stringtab .
    METHODS get_cell_coord
      IMPORTING
        !iv_coord       TYPE string
      RETURNING
        VALUE(es_coord) TYPE t_cell_coord
      RAISING
        zcx_excel.
    METHODS put_cell_to_worksheet
      IMPORTING
        !io_worksheet TYPE REF TO zcl_excel_worksheet
        !is_cell      TYPE t_cell
      RAISING
        zcx_excel.
    METHODS get_shared_string
      IMPORTING
        !iv_index       TYPE any
      RETURNING
        VALUE(ev_value) TYPE string
      RAISING
        lcx_not_found .
    METHODS get_style
      IMPORTING
        !iv_index            TYPE any
      RETURNING
        VALUE(ev_style_guid) TYPE zexcel_cell_style
      RAISING
        lcx_not_found .
    METHODS read_worksheet_data
      IMPORTING
        !io_reader    TYPE REF TO if_sxml_reader
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        lcx_not_found
        zcx_excel .
    METHODS get_sxml_reader
      IMPORTING
        !iv_path         TYPE string
      RETURNING
        VALUE(eo_reader) TYPE REF TO if_sxml_reader
      RAISING
        zcx_excel .
ENDCLASS.



CLASS zcl_excel_reader_huge_file IMPLEMENTATION.


  METHOD fill_cell_from_attributes.

    WHILE io_reader->node_type NE c_end_of_stream.
      io_reader->next_attribute( ).
      IF io_reader->node_type NE c_attribute.
        EXIT.
      ENDIF.
      CASE io_reader->name.
        WHEN `t`.
          es_cell-datatype = io_reader->value.
        WHEN `s`.
          IF io_reader->value IS NOT INITIAL.
            es_cell-style = get_style( io_reader->value ).
          ENDIF.
        WHEN `r`.
          es_cell-coord = get_cell_coord( io_reader->value ).
      ENDCASE.
    ENDWHILE.

  ENDMETHOD.


  METHOD get_cell_coord.

    zcl_excel_common=>convert_columnrow2column_a_row(
      EXPORTING
        i_columnrow = iv_coord
      IMPORTING
        e_column    = es_coord-column
        e_row       = es_coord-row
      ).

  ENDMETHOD.


  METHOD get_shared_string.
    DATA: lv_tabix TYPE i,
          lv_error TYPE string.
    FIELD-SYMBOLS: <ls_shared_string> TYPE t_shared_string.
    lv_tabix = iv_index + 1.
    READ TABLE shared_strings ASSIGNING <ls_shared_string> INDEX lv_tabix.
    IF sy-subrc NE 0.
      CONCATENATE 'Entry ' iv_index ' not found in Shared String Table' INTO lv_error.
      RAISE EXCEPTION TYPE lcx_not_found
        EXPORTING
          error = lv_error.
    ENDIF.
    ev_value = <ls_shared_string>-value.
  ENDMETHOD.


  METHOD get_style.

    DATA: lv_tabix TYPE i,
          lo_style TYPE REF TO zcl_excel_style,
          lv_error TYPE string.

    IF gs_buffer_style-index NE iv_index.
      lv_tabix = iv_index + 1.
      READ TABLE styles INTO lo_style INDEX lv_tabix.
      IF sy-subrc NE 0.
        CONCATENATE 'Entry ' iv_index ' not found in Style Table' INTO lv_error.
        RAISE EXCEPTION TYPE lcx_not_found
          EXPORTING
            error = lv_error.
      ELSE.
        gs_buffer_style-index = iv_index.
        gs_buffer_style-guid  = lo_style->get_guid( ).
      ENDIF.
    ENDIF.

    ev_style_guid = gs_buffer_style-guid.

  ENDMETHOD.


  METHOD get_sxml_reader.

    DATA: lv_xml TYPE xstring.

    lv_xml = get_from_zip_archive( iv_path ).
    eo_reader = cl_sxml_string_reader=>create( lv_xml ).

  ENDMETHOD.


  METHOD load_shared_strings.

    DATA: lo_reader TYPE REF TO if_sxml_reader.
    DATA: lt_shared_strings TYPE TABLE OF string,
          ls_shared_string  TYPE t_shared_string.
    FIELD-SYMBOLS: <lv_shared_string> TYPE string.

    lo_reader = get_sxml_reader( ip_path ).

    lt_shared_strings = read_shared_strings( lo_reader ).
    LOOP AT lt_shared_strings ASSIGNING <lv_shared_string>.
      ls_shared_string-value = <lv_shared_string>.
      APPEND ls_shared_string TO shared_strings.
    ENDLOOP.

  ENDMETHOD.


  METHOD load_worksheet.

    DATA: lo_reader TYPE REF TO if_sxml_reader.
    DATA: lx_not_found TYPE REF TO lcx_not_found.

    lo_reader = get_sxml_reader( ip_path ).

    TRY.

        read_worksheet_data( io_reader    = lo_reader
                             io_worksheet = io_worksheet ).

      CATCH lcx_not_found INTO lx_not_found.
        zcx_excel=>raise_text( lx_not_found->error ).
    ENDTRY.
  ENDMETHOD.


  METHOD put_cell_to_worksheet.
    CHECK is_cell-value IS NOT INITIAL
       OR is_cell-formula IS NOT INITIAL
       OR is_cell-style IS NOT INITIAL.
    CALL METHOD io_worksheet->set_cell
      EXPORTING
        ip_column    = is_cell-column
        ip_row       = is_cell-row
        ip_value     = is_cell-value
        ip_formula   = is_cell-formula
        ip_data_type = is_cell-datatype
        ip_style     = is_cell-style.
  ENDMETHOD.


  METHOD read_shared_strings.

    DATA lv_value TYPE string.

    WHILE io_reader->node_type NE c_end_of_stream.
      io_reader->next_node( ).
      CASE io_reader->name.
        WHEN 'si'.
          CASE io_reader->node_type .
            WHEN c_element_open .
              CLEAR lv_value .
            WHEN c_element_close .
              APPEND lv_value TO et_shared_strings.
          ENDCASE .
        WHEN 't'.
          CASE io_reader->node_type .
            WHEN c_node_value .
              lv_value = lv_value && io_reader->value .
          ENDCASE .
      ENDCASE .
    ENDWHILE.

  ENDMETHOD.


  METHOD read_worksheet_data.

    DATA: ls_cell   TYPE t_cell.

* Skip to <sheetData> element
    skip_to(  iv_element_name = `sheetData`  io_reader = io_reader ).

* Main loop: Evaluate the <c> elements and its children
    WHILE io_reader->node_type NE c_end_of_stream.
      io_reader->next_node( ).
      CASE io_reader->node_type.
        WHEN c_element_open.
          IF io_reader->name EQ `c`.
            ls_cell = fill_cell_from_attributes( io_reader ).
          ENDIF.
        WHEN c_node_value.
          CASE io_reader->name.
            WHEN `f`.
              ls_cell-formula = io_reader->value.
            WHEN `v`.
              IF ls_cell-datatype EQ `s`.
                ls_cell-value = get_shared_string( io_reader->value ).
              ELSE.
                ls_cell-value = io_reader->value.
              ENDIF.
            WHEN `t` OR `is`.
              ls_cell-value = io_reader->value.
          ENDCASE.
        WHEN c_element_close.
          CASE io_reader->name.
            WHEN `c`.
              put_cell_to_worksheet( is_cell = ls_cell io_worksheet = io_worksheet ).
            WHEN `sheetData`.
              EXIT.
          ENDCASE.
      ENDCASE.
    ENDWHILE.

  ENDMETHOD.


  METHOD skip_to.

    DATA: lv_error TYPE string.

* Skip forward to given element
    WHILE io_reader->name NE iv_element_name OR
          io_reader->node_type NE c_element_open.
      io_reader->next_node( ).
      IF io_reader->node_type = c_end_of_stream.
        CONCATENATE 'XML error: Didn''t find element <' iv_element_name '>' INTO lv_error.
        RAISE EXCEPTION TYPE lcx_not_found
          EXPORTING
            error = lv_error.
      ENDIF.
    ENDWHILE.


  ENDMETHOD.
ENDCLASS.
